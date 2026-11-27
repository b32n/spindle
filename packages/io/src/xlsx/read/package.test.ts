import { ImportError } from '../../common/import-error';
import { decodePart, XlsxPackage } from './package';
import { corpus, relsXml, zipOf } from './test-support';

const code = async (bytes: Uint8Array): Promise<string | undefined> => {
  try {
    await XlsxPackage.open(bytes);
    return undefined;
  } catch (e) {
    expect(e).toBeInstanceOf(ImportError);
    return (e as ImportError).code;
  }
};

describe('XlsxPackage.open', () => {
  it.each([
    ['edge-encrypted.xlsx', 'encrypted'],
    ['edge-legacy-xls.xls', 'legacy-binary'],
    ['edge-truncated-zip.xlsx', 'corrupt'],
  ])('refuses %s with the named error (%s)', async (file, expected) => expect(await code(corpus(file))).toBe(expected));
  it('refuses empty bytes and bytes that are not a zip', async () => {
    expect(await code(new Uint8Array())).toBe('not-a-package');
    expect(await code(new TextEncoder().encode('hello'))).toBe('not-a-package');
  });
  it('applies the limits it is given', async () => {
    await expect(XlsxPackage.open(await zipOf({ 'a.xml': '<a/>' }), { limits: { maxEntries: 1 } })).rejects.toMatchObject({ code: 'size-limit' });
  });
  it('opens a package whose content types are not well-formed, with a warning', async () => {
    const pkg = await XlsxPackage.open(await zipOf({ '[Content_Types].xml': '<Types><oops' }));
    expect(pkg.contentTypes.defaults.size).toBe(0);
    expect(pkg.warnings).toEqual(['The part [Content_Types].xml is not well-formed XML and was skipped.']);
  });
  it('reads the content types', async () => {
    const pkg = await XlsxPackage.open(corpus('openpyxl-comment.xlsx'));
    expect(pkg.contentType('xl/workbook.xml')).toMatch(/spreadsheetml\.sheet\.main\+xml$/);
    expect(pkg.warnings).toEqual([]);
  });
});

describe('XlsxPackage parts', () => {
  const open = async () =>
    XlsxPackage.open(
      await zipOf({
        'xl/Worksheets/Sheet1.xml': '<a><b/></a>',
        'xl/bad.xml': '<a><b',
        'xl/text.txt': 'plain',
        'dir/': '',
        '_rels/.rels': relsXml([['rId1', 'officeDocument', 'xl/workbook.xml']]),
        'xl/_rels/book.xml.rels': relsXml([['rId1', 'worksheet', 'Worksheets/Sheet1.xml']]),
        'xl/book.xml': '<book/>',
      }),
    );

  it('lists file entries without directories', async () => {
    const names = (await open()).partNames;
    expect(names).toContain('xl/text.txt');
    expect(names).not.toContain('dir/');
  });
  it('finds a part by exact name, with a leading slash, or in another case', async () => {
    const pkg = await open();
    expect(pkg.resolve('xl/Worksheets/Sheet1.xml')).toBe('xl/Worksheets/Sheet1.xml');
    expect(pkg.resolve('/xl/Worksheets/Sheet1.xml')).toBe('xl/Worksheets/Sheet1.xml');
    expect(pkg.resolve('XL/worksheets/sheet1.XML')).toBe('xl/Worksheets/Sheet1.xml');
    expect(pkg.resolve('xl/none.xml')).toBeUndefined();
    expect(pkg.has('xl/text.txt')).toBe(true);
    expect(pkg.has('nope')).toBe(false);
  });
  it('reads text, bytes and parsed XML; absent parts are undefined', async () => {
    const pkg = await open();
    expect(await pkg.text('xl/text.txt')).toBe('plain');
    expect(Array.from((await pkg.bytes('xl/text.txt')) ?? [])).toEqual([112, 108, 97, 105, 110]);
    expect((await pkg.xml('xl/Worksheets/Sheet1.xml'))?.children[0].name).toBe('b');
    expect(await pkg.text('xl/none.xml')).toBeUndefined();
    expect(await pkg.bytes('xl/none.xml')).toBeUndefined();
    expect(await pkg.xml('xl/none.xml')).toBeUndefined();
    expect(pkg.warnings).toEqual([]);
  });
  it('warns and answers undefined for XML that is not well-formed, never throws', async () => {
    const pkg = await open();
    expect(await pkg.xml('xl/bad.xml')).toBeUndefined();
    expect(pkg.warnings).toEqual(['The part xl/bad.xml is not well-formed XML and was skipped.']);
  });
  it('warns when an entry cannot be inflated after the open check', async () => {
    const pkg = await open();
    // the safe zip inflated every entry once at open, so this is a guard for a failure after that
    const file = (pkg as unknown as { zip: { file(n: string): { async: () => Promise<never> } } }).zip.file('xl/text.txt');
    file.async = () => Promise.reject(new Error('boom'));
    expect(await pkg.bytes('xl/text.txt')).toBeUndefined();
    expect(pkg.warnings).toEqual(['The part xl/text.txt could not be read.']);
  });
  it('reads relationships of the package, of a part, and caches them; none for a part without a rels file', async () => {
    const pkg = await open();
    expect((await pkg.rels('')).map((r) => r.target)).toEqual(['xl/workbook.xml']);
    expect((await pkg.rels('/')).map((r) => r.target)).toEqual(['xl/workbook.xml']);
    const first = await pkg.rels('xl/book.xml');
    expect(first.map((r) => r.target)).toEqual(['xl/Worksheets/Sheet1.xml']);
    expect(await pkg.rels('xl/book.xml')).toBe(first);
    expect(await pkg.rels('xl/text.txt')).toEqual([]);
  });
  it('keeps the first of two entries whose names differ only by case', async () => {
    const pkg = await XlsxPackage.open(await zipOf({ 'xl/A.xml': '<first/>', 'xl/a.xml': '<second/>' }));
    expect((await pkg.xml('XL/A.XML'))?.name).toBe('first');
  });
});

describe('decodePart', () => {
  it('decodes UTF-8 with or without a byte order mark', () => {
    expect(decodePart(new TextEncoder().encode('héllo'))).toBe('héllo');
    expect(decodePart(Uint8Array.from([0xef, 0xbb, 0xbf, 0x61]))).toBe('a');
    expect(decodePart(new Uint8Array())).toBe('');
  });
  it('decodes UTF-16 in either byte order when it has a mark', () => {
    expect(decodePart(Uint8Array.from([0xff, 0xfe, 0x61, 0x00, 0xe9, 0x00]))).toBe('aé');
    expect(decodePart(Uint8Array.from([0xfe, 0xff, 0x00, 0x61, 0x00, 0xe9]))).toBe('aé');
  });
  it('reads a UTF-16 part as XML through the package', async () => {
    const xml = '<?xml version="1.0" encoding="UTF-16"?><a><b/></a>';
    const bytes = new Uint8Array(2 + xml.length * 2);
    bytes.set([0xff, 0xfe]);
    for (let i = 0; i < xml.length; i++) {
      bytes[2 + i * 2] = xml.charCodeAt(i);
    }
    const pkg = await XlsxPackage.open(await zipOf({ 'u.xml': bytes }));
    expect((await pkg.xml('u.xml'))?.name).toBe('a');
  });
});
