import { WorkbookImpl } from '@b32nio/spindle-sheets-core';
import { exportWorkbookFile, importWorkbookFile, SHEETS_EXPORT_FORMATS } from './file-io';

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
