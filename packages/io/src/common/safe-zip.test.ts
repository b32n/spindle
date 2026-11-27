import JSZip from 'jszip';
import { ImportError, type ImportErrorCode } from './import-error';
import { declaredEntryCount, DEFAULT_IMPORT_LIMITS, isUnsafePartName, openSafeZip, resolveLimits, type ImportLimits } from './safe-zip';

const KB = 1024;
const MB = KB * KB;

async function build(files: Record<string, string | Uint8Array>, withTypes = true): Promise<Uint8Array> {
  const zip = new JSZip();
  if (withTypes) zip.file('[Content_Types].xml', '<Types/>');
  for (const [name, content] of Object.entries(files)) zip.file(name, content);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

async function expectRefused(bytes: Uint8Array, code: ImportErrorCode, message: RegExp, limits?: Partial<ImportLimits>): Promise<void> {
  const error = await openSafeZip(bytes, { limits }).then(
    () => undefined,
    (e: unknown) => e
  );
  expect(error).toBeInstanceOf(ImportError);
  expect((error as ImportError).code).toBe(code);
  expect((error as ImportError).message).toMatch(message);
}

/** Overwrite the declared size in the local and central directory headers of the entries named in `names`: forged headers. */
function forgeSizes(zip: Uint8Array, uncompressedSize: number, names: string[]): Uint8Array {
  const out = zip.slice();
  const view = new DataView(out.buffer);
  const nameAt = (at: number, length: number) => new TextDecoder().decode(out.subarray(at, at + length));
  for (let i = 0; i + 4 < out.length; i++) {
    const sig = view.getUint32(i, true);
    if (sig === 0x04034b50 && names.includes(nameAt(i + 30, view.getUint16(i + 26, true)))) view.setUint32(i + 22, uncompressedSize, true);
    else if (sig === 0x02014b50 && names.includes(nameAt(i + 46, view.getUint16(i + 28, true)))) view.setUint32(i + 24, uncompressedSize, true);
  }
  return out;
}

/** Overwrite the compressed data of the entry `name` with bytes that are not valid deflate: corrupt data under a good header. */
function corruptData(zip: Uint8Array, name: string): Uint8Array {
  const out = zip.slice();
  const view = new DataView(out.buffer);
  for (let i = 0; i + 30 < out.length; i++) {
    if (view.getUint32(i, true) !== 0x04034b50) continue;
    const nameLength = view.getUint16(i + 26, true);
    if (new TextDecoder().decode(out.subarray(i + 30, i + 30 + nameLength)) !== name) continue;
    const at = i + 30 + nameLength + view.getUint16(i + 28, true);
    out.fill(0xff, at, at + view.getUint32(i + 18, true));
  }
  return out;
}

const utf16 = (s: string) => Array.from(s).flatMap((c) => [c.charCodeAt(0), 0]);
const ole = (...streams: string[]) => Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...Array(512).fill(0), ...streams.flatMap((s) => [...utf16(s), 0, 0])]);

describe('limits', () => {
  it('has the documented defaults and merges overrides over them', () => {
    expect(DEFAULT_IMPORT_LIMITS).toEqual({
      maxPackageBytes: 100 * MB,
      maxEntries: 20_000,
      maxTotalBytes: 1024 * MB,
      maxPartBytes: 500 * MB,
      maxXmlPartBytes: 256 * MB,
      maxRatio: 200,
      ratioFloorBytes: 16 * MB,
      maxSheetFormulaChars: 32 * MB,
      maxFormulaChars: 64 * MB,
      maxTotalCells: 5_000_000,
      maxSheets: 1_000,
    });
    expect(resolveLimits()).toEqual(DEFAULT_IMPORT_LIMITS);
    expect(resolveLimits({ maxEntries: 3 })).toEqual({ ...DEFAULT_IMPORT_LIMITS, maxEntries: 3 });
  });
});

describe('isUnsafePartName', () => {
  it.each(['../evil.xml', 'a/../../evil.xml', '/etc/passwd', 'C:/windows/x', 'c:evil', 'a\\b.xml', 'a\0b.xml', 'xl/..'])('refuses %p', (name) => {
    expect(isUnsafePartName(name)).toBe(true);
  });
  it.each(['xl/workbook.xml', '[Content_Types].xml', 'xl/media/image1.png', 'a/b..c/d.xml', 'xl/worksheets/', 'x:y/z'.replace('x:y/z', 'xx/y:z')])('accepts %p', (name) => {
    expect(isUnsafePartName(name)).toBe(false);
  });
});

describe('declaredEntryCount', () => {
  it('reads the count from the end record', async () => {
    expect(declaredEntryCount(await build({ 'a.xml': '1', 'b.xml': '2' }))).toBe(3);
  });

  it('answers undefined without an end record, also for short input', () => {
    expect(declaredEntryCount(new Uint8Array(10))).toBeUndefined();
    expect(declaredEntryCount(new Uint8Array(1000))).toBeUndefined();
  });

  /** An end record (22 bytes) preceded by a zip64 locator (20 bytes) and, at `zip64At`, a zip64 end record claiming `count`. */
  function zip64(count: number, { sigAtOffset = true, offsetValue = 0, locator = true } = {}): Uint8Array {
    const bytes = new Uint8Array(56 + (locator ? 20 : 0) + 22);
    const view = new DataView(bytes.buffer);
    if (sigAtOffset) view.setUint32(0, 0x06064b50, true);
    view.setBigUint64(32, BigInt(count), true);
    let at = 56;
    if (locator) {
      view.setUint32(at, 0x07064b50, true);
      view.setBigUint64(at + 8, BigInt(offsetValue), true);
      at += 20;
    }
    view.setUint32(at, 0x06054b50, true);
    view.setUint16(at + 10, 0xffff, true);
    return bytes;
  }

  it('reads a zip64 count from the zip64 end record', () => {
    expect(declaredEntryCount(zip64(123_456))).toBe(123_456);
  });
  it('falls back to 0xFFFF when the zip64 record is not where the locator says, or there is no locator', () => {
    expect(declaredEntryCount(zip64(5, { sigAtOffset: false }))).toBe(0xffff);
    expect(declaredEntryCount(zip64(5, { offsetValue: 1_000_000 }))).toBe(0xffff);
    expect(declaredEntryCount(zip64(5, { locator: false }))).toBe(0xffff);
  });
  it('does not read a locator before the start of the file', () => {
    const bytes = new Uint8Array(22);
    const view = new DataView(bytes.buffer);
    view.setUint32(0, 0x06054b50, true);
    view.setUint16(10, 0xffff, true);
    expect(declaredEntryCount(bytes)).toBe(0xffff);
  });
});

describe('openSafeZip: what the file is', () => {
  it('opens an ordinary package and returns the zip', async () => {
    const zip = await openSafeZip(await build({ 'xl/workbook.xml': '<workbook/>' }));
    expect(await zip.file('xl/workbook.xml')!.async('string')).toBe('<workbook/>');
  });

  it('opens a package with directory entries', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types/>');
    zip.folder('xl')!.folder('worksheets');
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    await expect(openSafeZip(bytes)).resolves.toBeDefined();
  });

  it('refuses an empty file', async () => {
    await expectRefused(new Uint8Array(0), 'not-a-package', /empty/);
  });

  it.each([
    ['random bytes', Uint8Array.from({ length: 4096 }, (_, i) => (i * 7919 + 13) & 0xff)],
    ['text', new TextEncoder().encode('hello, world')],
    ['a PNG', Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3])],
    ['PK with a bad record type', Uint8Array.from([0x50, 0x4b, 9, 9, 0, 0])],
    ['PK with a bad second marker', Uint8Array.from([0x50, 0x4b, 3, 9, 0, 0])],
  ])('refuses %s as not a package', async (_name, bytes) => {
    await expectRefused(bytes, 'not-a-package', /not an Excel workbook|This is not/);
  });

  it('names the wanted format in the message', async () => {
    await expect(openSafeZip(new Uint8Array([1, 2, 3]), { target: 'a Word document (.docx)' })).rejects.toThrow('This is not a Word document (.docx).');
  });

  it('recognises an encrypted package (OLE container with EncryptedPackage or EncryptionInfo)', async () => {
    await expectRefused(ole('EncryptionInfo', 'EncryptedPackage'), 'encrypted', /password-protected/);
    await expectRefused(ole('EncryptionInfo'), 'encrypted', /password-protected/);
  });

  it('recognises a legacy binary file (OLE container without encryption streams)', async () => {
    await expectRefused(ole('Workbook'), 'legacy-binary', /legacy binary Office file/);
  });

  it('recognises a password-protected zip entry', async () => {
    const zip = await build({ 'xl/workbook.xml': '<workbook/>' });
    const view = new DataView(zip.buffer);
    for (let i = 0; i + 8 < zip.length; i++) {
      if (view.getUint32(i, true) === 0x04034b50) view.setUint16(i + 6, view.getUint16(i + 6, true) | 1, true);
      if (view.getUint32(i, true) === 0x02014b50) view.setUint16(i + 8, view.getUint16(i + 8, true) | 1, true);
    }
    await expectRefused(zip, 'encrypted', /password-protected/);
  });

  it('refuses a truncated zip as corrupt', async () => {
    const full = await build({ 'xl/workbook.xml': '<workbook/>'.repeat(100) });
    await expectRefused(full.subarray(0, Math.floor(full.length * 0.6)), 'corrupt', /damaged/);
  });

  it('recognises OpenDocument by its mimetype entry', async () => {
    const bytes = await build({ mimetype: 'application/vnd.oasis.opendocument.spreadsheet', 'content.xml': '<x/>' }, false);
    await expectRefused(bytes, 'not-a-package', /OpenDocument/);
  });

  it('opens a package that merely has a mimetype entry of another kind', async () => {
    await expect(openSafeZip(await build({ mimetype: 'image/png' }))).resolves.toBeDefined();
  });

  it('does not read a mimetype entry that is large (not an ODF marker)', async () => {
    await expect(openSafeZip(await build({ mimetype: 'application/vnd.oasis.opendocument.text'.padEnd(1000, ' ') }))).resolves.toBeDefined();
  });

  it('does not trust the declared size of a mimetype entry: a forged header cannot make it inflate unbounded', async () => {
    const bomb = await build({ mimetype: new Uint8Array(3 * MB) }); // really 3 MB of zeros, declared as 10 bytes
    const forged = forgeSizes(bomb, 10, ['mimetype']);
    // the entry is read through the bounded stream (stopped at 255 bytes), then the guard finds the mismatch
    await expectRefused(forged, 'corrupt', /damaged/);
  });

  it('refuses a forged mimetype header that really holds an ODF marker', async () => {
    const forged = forgeSizes(await build({ mimetype: 'application/vnd.oasis.opendocument.text' }), 10, ['mimetype']);
    await expectRefused(forged, 'corrupt', /damaged/);
  });

  it('reports corrupt data in the mimetype entry as damaged (an ImportError, not a plain Error)', async () => {
    await expectRefused(corruptData(await build({ mimetype: 'application/vnd.oasis.opendocument.text' }), 'mimetype'), 'corrupt', /damaged/);
  });

  it.each([['Index/Document.iwa'], ['Index.zip'], ['index.apxl'], ['index.xml.gz']])('recognises an iWork package by %s', async (part) => {
    await expectRefused(await build({ [part]: 'x' }, false), 'not-a-package', /Apple Numbers, Keynote or Pages/);
  });

  it('recognises a binary workbook (.xlsb)', async () => {
    await expectRefused(await build({ 'xl/workbook.bin': 'x' }), 'legacy-binary', /\.xlsb/);
  });

  it('refuses a zip that has no [Content_Types].xml', async () => {
    await expectRefused(await build({ 'hello.txt': 'hi' }, false), 'not-a-package', /no \[Content_Types\]\.xml/);
  });
});

describe('openSafeZip: names and counts', () => {
  it.each([['../evil.xml'], ['/abs.xml'], ['a\\b.xml'], ['C:/x.xml'], ['a/../../b.xml']])('refuses the part name %p', async (name) => {
    await expectRefused(await build({ [name]: 'x' }), 'corrupt', /invalid path/);
  });

  it('refuses a package over the size limit before reading it', async () => {
    await expectRefused(await build({ 'a.xml': 'x' }), 'size-limit', /the file \(\d+ bytes\) is over the limit of 100 bytes/, { maxPackageBytes: 100 });
    await expectRefused(new Uint8Array(3 * MB), 'size-limit', /the file \(3 MB\) is over the limit of 2 MB/, { maxPackageBytes: 2 * MB });
  });

  it('refuses too many entries, by the end record before JSZip builds anything', async () => {
    const files = Object.fromEntries(Array.from({ length: 10 }, (_, i) => [`p${i}.xml`, 'x']));
    await expectRefused(await build(files), 'size-limit', /number of entries \(11\) is over the limit of 5/, { maxEntries: 5 });
  });

  it('counts entries again after loading, in case the end record understates them', async () => {
    const files = Object.fromEntries(Array.from({ length: 10 }, (_, i) => [`p${i}.xml`, 'x']));
    const zip = await build(files);
    const view = new DataView(zip.buffer);
    for (let i = zip.length - 22; i >= 0; i--) {
      if (view.getUint32(i, true) === 0x06054b50) view.setUint16(i + 10, 1, true);
    }
    expect(declaredEntryCount(zip)).toBe(1);
    await expectRefused(zip, 'size-limit', /number of entries \(11\) is over the limit of 5/, { maxEntries: 5 });
  });
});

describe('openSafeZip: sizes', () => {
  const zeros = (n: number) => new Uint8Array(n);

  it('refuses a part over the part limit, and an XML part over the XML limit', async () => {
    await expectRefused(await build({ 'xl/media/a.bin': zeros(2000) }), 'size-limit', /a part \(\d+ bytes\) is over the limit of 1000 bytes/, { maxPartBytes: 1000, ratioFloorBytes: 1e9 });
    await expectRefused(await build({ 'xl/sheet.xml': zeros(2000) }), 'size-limit', /a part \(\d+ bytes\) is over the limit of 1500 bytes/, { maxXmlPartBytes: 1500, ratioFloorBytes: 1e9 });
    // a binary part may be larger than an XML part may
    await expect(openSafeZip(await build({ 'xl/media/a.bin': zeros(2000) }), { limits: { maxXmlPartBytes: 1500, ratioFloorBytes: 1e9 } })).resolves.toBeDefined();
  });

  it('refuses a total over the total limit', async () => {
    const bytes = await build({ 'a.bin': zeros(800), 'b.bin': zeros(800) });
    await expectRefused(bytes, 'size-limit', /the content \(\d+ bytes\) is over the limit of 1000 bytes/, { maxTotalBytes: 1000, ratioFloorBytes: 1e9 });
  });

  it('refuses a part whose compression ratio is over the limit, but only above the floor', async () => {
    const bytes = await build({ 'xl/media/a.bin': zeros(100 * KB) });
    await expectRefused(bytes, 'size-limit', /the compression ratio of a part is over the limit of 200:1/, { ratioFloorBytes: 50 * KB });
    // a small part may compress as well as it likes: below the floor it is not judged
    await expect(openSafeZip(bytes, { limits: { ratioFloorBytes: 200 * KB } })).resolves.toBeDefined();
  });

  it('refuses a package whose overall ratio is over the limit when no single part is judged', async () => {
    const bytes = await build({ 'a.bin': zeros(60 * KB), 'b.bin': zeros(60 * KB), 'c.bin': zeros(60 * KB) });
    await expectRefused(bytes, 'size-limit', /the compression ratio is over the limit of 200:1/, { ratioFloorBytes: 100 * KB, maxRatio: 200 });
  });

  it('a real zip bomb: 30 MB of zeros in a few KB is refused with the default limits', async () => {
    const bomb = await build({ 'xl/worksheets/sheet1.xml': zeros(30 * MB) });
    expect(bomb.length).toBeLessThan(100 * KB);
    await expectRefused(bomb, 'size-limit', /compression ratio/);
  });

  it('does not refuse a legitimate large, repetitive but not bomb-like part', async () => {
    // 20 MB of pseudo-random digits compresses about 2:1: nothing to see
    let seed = 1;
    const digits = new TextEncoder().encode(Array.from({ length: 2 * MB }, () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) % 10).toString()).join(''));
    await expect(openSafeZip(await build({ 'xl/worksheets/sheet1.xml': digits }))).resolves.toBeDefined();
  });
});

describe('openSafeZip: streaming guard (forged headers)', () => {
  it('stops a bomb whose header declares a small size', async () => {
    const forged = forgeSizes(await build({ 'xl/media/a.bin': new Uint8Array(200 * KB) }), 10, ['xl/media/a.bin']);
    await expectRefused(forged, 'size-limit', /a part is over the limit of 1000 bytes/, { maxPartBytes: 1000, ratioFloorBytes: 1e9 });
  });

  it('reports a forged header that understates a small part as corrupt (the size does not match what inflates)', async () => {
    const forged = forgeSizes(await build({ 'a.bin': new Uint8Array(800) }), 10, ['a.bin']);
    await expectRefused(forged, 'corrupt', /a part cannot be read/);
  });

  it('stops a forged part that takes the total past the total limit', async () => {
    const forged = forgeSizes(await build({ 'a.bin': new Uint8Array(800), 'b.bin': new Uint8Array(800) }), 10, ['b.bin']);
    await expectRefused(forged, 'size-limit', /the content is over the limit of 1000 bytes/, { maxTotalBytes: 1000, ratioFloorBytes: 1e9 });
  });

  it('stops a forged part by its real compression ratio', async () => {
    const forged = forgeSizes(await build({ 'xl/media/a.bin': new Uint8Array(100 * KB) }), 10, ['xl/media/a.bin']);
    await expectRefused(forged, 'size-limit', /the compression ratio of a part is over the limit of 200:1/, { ratioFloorBytes: 50 * KB });
  });

  it('reports a part that cannot be inflated as corrupt', async () => {
    const payload = new TextEncoder().encode('0123456789'.repeat(1000));
    const zip = await build({ 'xl/sheet.xml': payload });
    const view = new DataView(zip.buffer);
    // the entry data follows its local header; replace it with an invalid deflate block (type 3)
    for (let i = 0; i + 30 < zip.length; i++) {
      if (view.getUint32(i, true) !== 0x04034b50) continue;
      const nameLength = view.getUint16(i + 26, true);
      const extraLength = view.getUint16(i + 28, true);
      const dataAt = i + 30 + nameLength + extraLength;
      const compressed = view.getUint32(i + 18, true);
      if (view.getUint16(i + 8, true) === 8 && compressed > 4) zip.fill(0xff, dataAt, dataAt + compressed);
    }
    await expectRefused(zip, 'corrupt', /a part cannot be read/);
  });
});
