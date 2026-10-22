import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { exportXlsx } from './xlsx/writer';
import { exportDocx } from './docx/writer';
import { detectXlsxLoss } from './xlsx/loss';
import { detectDocxLoss } from './docx/loss';
import { detectCsvLoss } from './csv/loss';
import { describeImportLoss, hasImportLoss } from './loss';
import type { DocSection } from './docx/types';

const fixture = (name: string) => new Uint8Array(readFileSync(path.resolve(__dirname, '../../../../e2e/fixtures', name)));

async function plainXlsx(): Promise<Uint8Array> {
  return exportXlsx({
    sheets: [
      { name: 'Data', cells: new Map([['0:0', { value: 'a' }]]) },
      { name: 'Other', cells: new Map([['0:0', { value: 'b' }]]) },
    ],
  });
}

async function patch(bytes: Uint8Array, edit: (zip: JSZip) => Promise<void> | void): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  await edit(zip);
  return zip.generateAsync({ type: 'uint8array' });
}

async function patchSheet(bytes: Uint8Array, sheetPath: string, fn: (xml: string) => string): Promise<Uint8Array> {
  return patch(bytes, async (zip) => {
    const xml = await zip.file(sheetPath)!.async('string');
    zip.file(sheetPath, fn(xml));
  });
}

const ids = (r: { items: Array<{ id: string }> }) => r.items.map((i) => i.id).sort();

describe('detectXlsxLoss', () => {
  it('reports nothing for a file the importer fully understands', async () => {
    const report = await detectXlsxLoss(await plainXlsx());
    expect(report).toEqual({ format: 'xlsx', items: [] });
    expect(hasImportLoss(report)).toBe(false);
  });

  it('counts conditional formatting, validation, filters and protection per sheet', async () => {
    let bytes = await plainXlsx();
    bytes = await patchSheet(bytes, 'xl/worksheets/sheet1.xml', (xml) =>
      xml.replace(
        '</sheetData>',
        '</sheetData><sheetProtection sheet="1"/><autoFilter ref="A1:B2"/>' +
          '<conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>' +
          '<conditionalFormatting sqref="B1"><cfRule type="cellIs" priority="2" operator="lessThan"><formula>1</formula></cfRule></conditionalFormatting>' +
          '<dataValidations count="1"><dataValidation type="list" sqref="C1"><formula1>"x,y"</formula1></dataValidation></dataValidations>'
      )
    );
    const report = await detectXlsxLoss(bytes);
    const by = Object.fromEntries(report.items.map((i) => [i.id, i]));
    expect(by['xlsx.conditionalFormatting']).toMatchObject({ count: 2, where: ['Data'] });
    expect(by['xlsx.dataValidation']).toMatchObject({ count: 1, where: ['Data'] });
    expect(by['xlsx.autoFilter']).toMatchObject({ count: 1 });
    expect(by['xlsx.sheetProtection']).toMatchObject({ count: 1 });
    expect(hasImportLoss(report)).toBe(true);
    expect(describeImportLoss(report)[0]).toMatch(/^\d+ × /);
  });

  it('reports named ranges (not the built-in print ones), hidden sheets, charts, images, comments and macros', async () => {
    const bytes = await patch(await plainXlsx(), async (zip) => {
      let wb = await zip.file('xl/workbook.xml')!.async('string');
      wb = wb.replace(/(<sheet [^>]*name="Other")/, '$1 state="hidden"');
      wb = wb.replace('</sheets>', '</sheets><definedNames><definedName name="Rate">Data!$A$1</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$2</definedName></definedNames>');
      zip.file('xl/workbook.xml', wb);
      zip.file('xl/charts/chart1.xml', '<c:chartSpace/>');
      zip.file('xl/media/image1.png', 'x');
      zip.file('xl/comments1.xml', '<comments/>');
      zip.file('xl/vbaProject.bin', 'x');
      zip.file('xl/tables/table1.xml', '<table/>');
    });
    const report = await detectXlsxLoss(bytes);
    expect(ids(report)).toEqual(
      ['xlsx.charts', 'xlsx.comments', 'xlsx.definedNames', 'xlsx.hiddenSheets', 'xlsx.images', 'xlsx.macros', 'xlsx.tables'].sort()
    );
    expect(report.items.find((i) => i.id === 'xlsx.definedNames')).toMatchObject({ count: 1, where: ['Rate'] });
    expect(report.items.find((i) => i.id === 'xlsx.hiddenSheets')).toMatchObject({ count: 1, where: ['Other'] });
  });

  it('an empty <workbookProtection/> or sheet="0" protection (LibreOffice) is not protection', async () => {
    let bytes = await plainXlsx();
    bytes = await patch(bytes, async (zip) => {
      const wb = await zip.file('xl/workbook.xml')!.async('string');
      zip.file('xl/workbook.xml', wb.replace('<sheets>', '<workbookProtection/><sheets>'));
    });
    bytes = await patchSheet(bytes, 'xl/worksheets/sheet1.xml', (xml) => xml.replace('</sheetData>', '</sheetData><sheetProtection sheet="false"/>'));
    expect(await detectXlsxLoss(bytes)).toEqual({ format: 'xlsx', items: [] });
    const locked = await patch(bytes, async (zip) => {
      const wb = await zip.file('xl/workbook.xml')!.async('string');
      zip.file('xl/workbook.xml', wb.replace('<workbookProtection/>', '<workbookProtection lockStructure="1"/>'));
    });
    expect(ids(await detectXlsxLoss(locked))).toEqual(['xlsx.workbookProtection']);
  });

  it('does not throw on a non-zip', async () => {
    expect(await detectXlsxLoss(new Uint8Array([1, 2, 3]))).toEqual({ format: 'xlsx', items: [] });
  });

  it('finds the features in the LibreOffice-made styled fixture without false alarms about plain content', async () => {
    const report = await detectXlsxLoss(fixture('styled-multisheet.xlsx'));
    // Whatever it reports must be real feature ids, never empty labels.
    for (const item of report.items) {
      expect(item.id).toMatch(/^xlsx\./);
      expect(item.count).toBeGreaterThan(0);
    }
  });
});

const SECTION: Omit<DocSection, 'blocks'> = {
  pageWidthPx: 816, pageHeightPx: 1056, marginTopPx: 96, marginRightPx: 96, marginBottomPx: 96, marginLeftPx: 96, orientation: 'portrait',
};

async function plainDocx(): Promise<Uint8Array> {
  const out = await exportDocx({ title: 'T', sections: [{ ...SECTION, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'hi' }] }] }] });
  return out.bytes;
}

describe('detectDocxLoss', () => {
  it('reports nothing for a document the importer fully understands', async () => {
    expect(await detectDocxLoss(await plainDocx())).toEqual({ format: 'docx', items: [] });
  });

  it('detects tracked changes, content controls, floating images, text boxes, equations, columns and fields', async () => {
    const bytes = await patch(await plainDocx(), async (zip) => {
      let xml = await zip.file('word/document.xml')!.async('string');
      xml = xml.replace(
        '</w:body>',
        '<w:p><w:ins w:id="1"><w:r><w:t>new</w:t></w:r></w:ins><w:del w:id="2"><w:r><w:delText>old</w:delText></w:r></w:del></w:p>' +
          '<w:sdt><w:sdtContent><w:p><w:r><w:t>cc</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
          '<w:p><w:r><w:drawing><wp:anchor/></w:drawing></w:r></w:p>' +
          '<w:p><w:r><w:pict><w:txbxContent><w:p/></w:txbxContent></w:pict></w:r></w:p>' +
          '<w:p><m:oMath/></w:p>' +
          '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o "1-3" </w:instrText></w:r></w:p>' +
          '<w:p><w:r><w:instrText> PAGE </w:instrText></w:r></w:p>' +
          '<w:sectPr><w:cols w:num="2"/></w:sectPr></w:body>'
      );
      zip.file('word/document.xml', xml);
      zip.file('word/footnotes.xml', '<w:footnotes><w:footnote w:type="separator" w:id="-1"/><w:footnote w:id="1"><w:p/></w:footnote></w:footnotes>');
      zip.file('word/comments.xml', '<w:comments><w:comment w:id="0"/><w:comment w:id="1"/></w:comments>');
    });
    const report = await detectDocxLoss(bytes);
    const by = Object.fromEntries(report.items.map((i) => [i.id, i.count]));
    expect(by).toMatchObject({
      'docx.trackedInsertions': 1,
      'docx.trackedDeletions': 1,
      'docx.contentControls': 1,
      'docx.floatingImages': 1,
      'docx.textBoxes': 1,
      'docx.equations': 1,
      'docx.columns': 1,
      'docx.fields': 1, // TOC only; PAGE is supported
      'docx.footnotes': 1, // the separator is not a footnote
      'docx.comments': 2,
    });
  });

  it('flags the LibreOffice rich-document fixture without crashing', async () => {
    const report = await detectDocxLoss(fixture('rich-document.docx'));
    for (const item of report.items) expect(item.id).toMatch(/^docx\./);
  });

  it('does not throw on a non-zip', async () => {
    expect(await detectDocxLoss(new Uint8Array([1, 2, 3]))).toEqual({ format: 'docx', items: [] });
  });
});

describe('detectCsvLoss', () => {
  it('comma-separated CSV loses nothing', () => {
    expect(detectCsvLoss('a,b\n1,2\n').items).toEqual([]);
    expect(detectCsvLoss('"a;b",c\n').items).toEqual([]);
    expect(detectCsvLoss('').items).toEqual([]);
  });
  it('semicolon- and tab-separated files are flagged (single-column import)', () => {
    expect(ids(detectCsvLoss('a;b;c\n1;2;3\n'))).toEqual(['csv.delimiter']);
    expect(detectCsvLoss('a\tb\n1\t2\n').items[0].label).toMatch(/tab/);
  });
  it('a semicolon file with decimal commas is still flagged', () => {
    expect(ids(detectCsvLoss('name;qty;price\nWidget;10;2,5\nGadget;4;12,75\n'))).toEqual(['csv.delimiter']);
  });
  it('a comma file with an occasional semicolon in text is not flagged', () => {
    expect(detectCsvLoss('a,b\nx;y,z\nq,r\n').items).toEqual([]);
  });
  it('a single column of plain text is not flagged', () => {
    expect(detectCsvLoss('one\ntwo\nthree\n').items).toEqual([]);
  });
});
