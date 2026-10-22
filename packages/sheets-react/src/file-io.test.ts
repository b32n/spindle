import { WorkbookImpl } from '@b32nio/spindle-sheets-core';
import JSZip from 'jszip';
import { exportWorkbookFile, importWorkbookFile, inspectWorkbookFile, SHEETS_EXPORT_FORMATS } from './file-io';

function makeWorkbook(): WorkbookImpl {
  const wb = new WorkbookImpl('wb1', 'Budget');
  const sheet = wb.getSheet();
  sheet.setCellValue(0, 0, 'Item');
  sheet.setCellValue(0, 1, 'Cost');
  sheet.setCellValue(1, 0, 'Pens');
  sheet.setCellValue(1, 1, 12);
  return wb;
}

function asFile(data: Uint8Array | string, name: string): File {
  return new File([typeof data === 'string' ? data : new Uint8Array(data)], name);
}

describe('sheets file I/O', () => {
  it('offers XLSX and CSV downloads', () => {
    expect(SHEETS_EXPORT_FORMATS.map((f) => f.id)).toEqual(['xlsx', 'csv']);
  });

  it('exports CSV with a sanitized file name', async () => {
    const out = await exportWorkbookFile(makeWorkbook(), 'csv', 'Q1/Budget');
    expect(out.fileName).toBe('Q1_Budget.csv');
    expect(out.mimeType).toContain('text/csv');
    expect(String(out.data)).toContain('Pens,12');
  });

  it('defaults the file name to the workbook name', async () => {
    const out = await exportWorkbookFile(makeWorkbook(), 'xlsx');
    expect(out.fileName).toBe('Budget.xlsx');
    expect(out.data).toBeInstanceOf(Uint8Array);
  });

  it('round-trips through XLSX', async () => {
    const out = await exportWorkbookFile(makeWorkbook(), 'xlsx');
    const target = new WorkbookImpl('wb2', 'Empty');
    await importWorkbookFile(target, asFile(out.data, out.fileName));
    expect(target.getSheet().getCell(1, 0)?.value).toBe('Pens');
    expect(target.getSheet().getCell(1, 1)?.value).toBe(12);
  });

  it('opens a CSV as a fresh workbook, replacing existing content', async () => {
    const target = makeWorkbook();
    await importWorkbookFile(target, asFile('x,y\n1,2\n', 'data.csv'));
    const sheet = target.getSheet();
    expect(sheet.getCell(0, 0)?.value).toBe('x');
    expect(sheet.getCell(1, 1)?.value).toBe(2);
    // Old content (row 1 "Pens") is gone.
    expect(sheet.getCell(1, 0)?.value).not.toBe('Pens');
    expect(target.sheets.size).toBe(1);
  });

  it('rejects unsupported file types without touching the workbook', async () => {
    const target = makeWorkbook();
    await expect(importWorkbookFile(target, asFile('hi', 'notes.txt'))).rejects.toThrow(/Unsupported file type/);
    expect(target.getSheet().getCell(1, 0)?.value).toBe('Pens');
  });
});

describe('sheets import loss report', () => {
  async function xlsxWithConditionalFormatting(): Promise<File> {
    const out = await exportWorkbookFile(makeWorkbook(), 'xlsx');
    const zip = await JSZip.loadAsync(out.data as Uint8Array);
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    zip.file(
      'xl/worksheets/sheet1.xml',
      xml.replace('</sheetData>', '</sheetData><conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="equal"><formula>1</formula></cfRule></conditionalFormatting>')
    );
    return asFile(await zip.generateAsync({ type: 'uint8array' }), 'cf.xlsx');
  }

  it('inspectWorkbookFile reports losses without touching the workbook', async () => {
    const wb = makeWorkbook();
    const report = await inspectWorkbookFile(await xlsxWithConditionalFormatting());
    expect(report.items.map((i) => i.id)).toEqual(['xlsx.conditionalFormatting']);
    expect(wb.getCellValue(undefined, 1, 1)).toBe(12);
  });

  it('importWorkbookFile returns the same report alongside the warnings', async () => {
    const wb = new WorkbookImpl('wb2', 'Other');
    const result = await importWorkbookFile(wb, await xlsxWithConditionalFormatting());
    expect(result.loss.items.map((i) => i.id)).toEqual(['xlsx.conditionalFormatting']);
    expect(wb.getCellValue(undefined, 1, 1)).toBe(12);
  });

  it('a clean file and a comma CSV report no loss; a semicolon CSV is flagged', async () => {
    const clean = await exportWorkbookFile(makeWorkbook(), 'xlsx');
    expect((await inspectWorkbookFile(asFile(clean.data, 'ok.xlsx'))).items).toEqual([]);
    expect((await inspectWorkbookFile(asFile('a,b\n1,2', 'ok.csv'))).items).toEqual([]);
    expect((await inspectWorkbookFile(asFile('a;b\n1;2', 'semi.csv'))).items.map((i) => i.id)).toEqual(['csv.delimiter']);
    const wb = new WorkbookImpl('wb3', 'X');
    expect((await importWorkbookFile(wb, asFile('a;b\n1;2', 'semi.csv'))).loss.items).toHaveLength(1);
  });

  it('rejects unsupported file types', async () => {
    await expect(inspectWorkbookFile(asFile('x', 'a.pdf'))).rejects.toThrow(/Unsupported/);
  });
});
