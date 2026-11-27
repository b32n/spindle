import JSZip from 'jszip';
import { corpus, el, MAIN_NS, numberFormatCases, relsXml, workbookZip, zipOf } from './test-support';

describe('the helpers of the reader tests', () => {
  it('reads a corpus file as bytes, and the number format table as data', () => {
    expect(corpus('number-formats.xlsx').length).toBeGreaterThan(0);
    const cases = numberFormatCases();
    expect(cases.length).toBeGreaterThan(10);
    for (const [raw, editor, isDate, why] of cases) expect([typeof raw, typeof editor, typeof isDate, typeof why]).toEqual(['string', 'string', 'boolean', 'string']);
  });

  it('writes relationships with the Office kind spelled short or in full, and a target mode when given', () => {
    expect(relsXml([['a', 'worksheet', 'w.xml'], ['b', 'http://x/y', 'https://e.com', 'External']])).toBe(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="a" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="w.xml"/>' +
        '<Relationship Id="b" Type="http://x/y" Target="https://e.com" TargetMode="External"/></Relationships>'
    );
  });

  it('zips parts (text or bytes), adding the content types part unless it is given', async () => {
    const zip = await JSZip.loadAsync(await zipOf({ 'a.xml': '<a/>', 'b.bin': new Uint8Array([1, 2]) }));
    expect(Object.keys(zip.files).sort()).toEqual(['[Content_Types].xml', 'a.xml', 'b.bin']);
    expect(await zip.file('b.bin')?.async('uint8array')).toEqual(new Uint8Array([1, 2]));
    const own = await JSZip.loadAsync(await zipOf({ '[Content_Types].xml': '<Types/>' }));
    expect(await own.file('[Content_Types].xml')?.async('string')).toBe('<Types/>');
  });

  it('builds a workbook package of a workbook body, its relationships and extra parts', async () => {
    const zip = await JSZip.loadAsync(await workbookZip('<sheets/>', [['rId1', 'worksheet', 'worksheets/sheet1.xml']], { 'xl/worksheets/sheet1.xml': '<worksheet/>' }));
    expect(await zip.file('xl/workbook.xml')?.async('string')).toContain(`<workbook ${MAIN_NS}`);
    expect(await zip.file('xl/_rels/workbook.xml.rels')?.async('string')).toContain('Id="rId1"');
    expect(zip.file('xl/worksheets/sheet1.xml')).not.toBeNull();
    // the extra parts are optional
    expect((await JSZip.loadAsync(await workbookZip('<sheets/>', []))).file('xl/workbook.xml')).not.toBeNull();
  });

  it('parses XML into a tree, and throws for text that is not XML', () => {
    expect(el('<a x="1"><b/></a>').children[0].name).toBe('b');
    expect(() => el('not xml')).toThrow('not XML: not xml');
  });
});
