import JSZip from 'jszip';
import { ImportError, type ImportErrorCode } from '../common/import-error';
import { detectDocxLoss } from './loss';
import { importDocx } from './reader';
import { parseXml } from './xml-tree';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const MB = 1024 * 1024;

/** A minimal package whose main part is `documentXml` (a body paragraph per string when it is an array). */
async function pack(documentXml: string, extra: Record<string, string | Uint8Array> = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types/>');
  zip.file('word/document.xml', documentXml);
  for (const [name, content] of Object.entries(extra)) zip.file(name, content);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 9 } });
}
const document = (body: string, doctype = '') => `<?xml version="1.0"?>${doctype}<w:document ${NS}><w:body>${body}</w:body></w:document>`;
const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

async function failure(promise: Promise<unknown>): Promise<ImportError> {
  try {
    await promise;
  } catch (e) {
    return e as ImportError;
  }
  throw new Error('expected the import to fail');
}

describe('importDocx: only an ImportError escapes', () => {
  const cases: Array<[string, () => Promise<Uint8Array>, ImportErrorCode]> = [
    ['an empty file', async () => new Uint8Array(), 'not-a-package'],
    ['text', async () => new TextEncoder().encode('hello'), 'not-a-package'],
    ['a legacy OLE file', async () => Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]), 'legacy-binary'],
    ['a truncated zip', async () => (await pack(document(para('x')))).slice(0, 60), 'corrupt'],
    ['a zip without [Content_Types].xml', async () => {
      const zip = new JSZip();
      zip.file('word/document.xml', document(para('x')));
      return zip.generateAsync({ type: 'uint8array' });
    }, 'not-a-package'],
    ['an OpenDocument text file', async () => {
      const zip = new JSZip();
      zip.file('mimetype', 'application/vnd.oasis.opendocument.text');
      return zip.generateAsync({ type: 'uint8array' });
    }, 'not-a-package'],
    ['a part with a path that leaves the package', async () => pack(document(para('x')), { '../evil.xml': '<a/>' }), 'corrupt'],
    ['a main part that is not XML', async () => pack('not xml <<<'), 'corrupt'],
    ['a main part nested past the depth cap', async () => pack(document('<w:p>'.repeat(1000) + '</w:p>'.repeat(1000))), 'corrupt'],
    ['a DOCTYPE that is never closed', async () => pack('<!DOCTYPE w:document [' + document(para('x'))), 'corrupt'],
  ];
  it.each(cases)('%s', async (_name, make, code) => {
    const error = await failure(importDocx(await make()));
    expect(error).toBeInstanceOf(ImportError);
    expect(error.code).toBe(code);
    expect(error.message.length).toBeGreaterThan(0);
  });

  it('keeps the warning (not an error) for a package without a main document', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types/>');
    const result = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(result.sections).toEqual([]);
    expect(result.warnings[0]).toBe('word/document.xml is missing — this file may not be a valid .docx.');
  });
});

describe('importDocx: limits', () => {
  it('refuses a part that inflates over the compression ratio, before reading it', async () => {
    const bytes = await pack(document(para('x')), { 'word/media/bomb.bin': new Uint8Array(24 * MB) });
    expect(bytes.length).toBeLessThan(100 * 1024);
    const error = await failure(importDocx(bytes));
    expect(error.code).toBe('size-limit');
    expect(error.message).toContain('compression ratio');
  });

  it('takes the limits as an option', async () => {
    const bytes = await pack(document(para('x')), { 'word/a.bin': new Uint8Array(4096) });
    expect((await importDocx(bytes)).sections).toHaveLength(1);
    expect((await failure(importDocx(bytes, { limits: { maxPartBytes: 1024 } }))).code).toBe('size-limit');
    expect((await failure(importDocx(bytes, { limits: { maxEntries: 1 } }))).code).toBe('size-limit');
    expect((await failure(importDocx(bytes, { limits: { maxPackageBytes: 10 } }))).code).toBe('size-limit');
  });

  it('does not refuse a legitimate large document (many paragraphs, a big picture)', async () => {
    const picture = new Uint8Array(8 * MB).map((_, i) => (i * 2654435761) >>> 24); // poorly compressible
    const bytes = await pack(document(para('x').repeat(20_000)), { 'word/media/big.png': picture });
    const result = await importDocx(bytes);
    expect(result.sections[0].blocks.length).toBeGreaterThanOrEqual(20_000);
  });
});

describe('importDocx: XML entities', () => {
  const doctype =
    '<!DOCTYPE w:document [<!ENTITY lol "lol"><!ENTITY lol1 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">' +
    '<!ENTITY lol2 "&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;"><!ENTITY xxe SYSTEM "file:///etc/passwd">]>';

  const textOf = (result: Awaited<ReturnType<typeof importDocx>>) =>
    JSON.stringify(result.sections[0].blocks);

  it('never expands declared entities: the references stay as written and the document still imports', async () => {
    const result = await importDocx(await pack(document(para('a &lol2; b &xxe; c') + para('after'), doctype)));
    const text = textOf(result);
    expect(text).toContain('a &lol2; b &xxe; c');
    expect(text).toContain('after');
    expect(text).not.toContain('lollol');
    expect(text).not.toContain('root:');
  });

  it('still decodes the predefined entities in text and attributes', async () => {
    const result = await importDocx(await pack(document(`<w:p><w:r><w:t>a &amp; b &lt;c&gt; &quot;d&quot; &apos;e&apos;</w:t></w:r></w:p>`)));
    expect(textOf(result)).toContain('a & b <c> \\"d\\" \'e\'');
  });

  it('parseXml removes the DOCTYPE and throws where the XML cannot be read', () => {
    expect(JSON.stringify(parseXml(`<!DOCTYPE a [<!ENTITY x "y">]><a>&x;</a>`))).toBe('[{"a":[{"#text":"&x;"}]}]');
    expect(() => parseXml('<!DOCTYPE a [<a/>')).toThrow('unterminated DOCTYPE');
    expect(() => parseXml('<a>'.repeat(1000) + '</a>'.repeat(1000))).toThrow();
  });
});

describe('detectDocxLoss: bounded like the import', () => {
  it('does not inflate a bomb: it answers an empty report', async () => {
    const bytes = await pack(document(para('x') + '<w:txbxContent/>'), { 'word/media/bomb.bin': new Uint8Array(24 * MB) });
    expect((await detectDocxLoss(bytes)).items).toEqual([]);
  });

  it('takes the limits as an option', async () => {
    const bytes = await pack(document('<w:txbxContent/>'), { 'word/a.bin': new Uint8Array(4096) });
    expect((await detectDocxLoss(bytes)).items.length).toBeGreaterThan(0);
    expect((await detectDocxLoss(bytes, { limits: { maxPartBytes: 1024 } })).items).toEqual([]);
  });
});
