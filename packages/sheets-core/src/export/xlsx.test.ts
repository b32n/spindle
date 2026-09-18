import ExcelJS from 'exceljs';
import { exportToXlsx, importFromXlsx } from './xlsx';
import { WorkbookImpl } from '../workbook';

async function readBack(workbook: WorkbookImpl): Promise<ExcelJS.Workbook> {
  const bytes = await exportToXlsx(workbook);
  const out = new ExcelJS.Workbook();
  // @ts-expect-error -- Buffer<ArrayBuffer> vs exceljs's plain Buffer type
  await out.xlsx.load(Buffer.from(bytes));
  return out;
}

function freshWorkbook(): WorkbookImpl {
  return new WorkbookImpl('wb', 'WB');
}

describe('exportToXlsx', () => {
  it('writes plain values', async () => {
    const wb = freshWorkbook();
    wb.setCellValue(undefined, 0, 0, 'hello');
    wb.setCellValue(undefined, 0, 1, 42);
    wb.setCellValue(undefined, 0, 2, true);

    const out = await readBack(wb);
    const ws = out.worksheets[0];
    expect(ws.name).toBe(wb.getSheet().name);
    expect(ws.getCell(1, 1).value).toBe('hello');
    expect(ws.getCell(1, 2).value).toBe(42);
    expect(ws.getCell(1, 3).value).toBe(true);
  });

  it('writes a formula with its calculated result', async () => {
    const wb = freshWorkbook();
    wb.setCellValue(undefined, 0, 0, 2);
    wb.setCellValue(undefined, 0, 1, 3);
    wb.setFormula(undefined, 0, 2, '=SUM(A1:B1)');

    const out = await readBack(wb);
    const cell = out.worksheets[0].getCell(1, 3);
    expect(cell.formula).toBe('SUM(A1:B1)'); // no leading '=' in xlsx's own convention
    expect(cell.result).toBe(5);
  });

  it('translates cell style to xlsx font/fill/alignment', async () => {
    const wb = freshWorkbook();
    wb.setCellValue(undefined, 0, 0, 'styled');
    const stylePool = wb.getStylePool();
    const styleId = stylePool.getOrCreate({
      bold: true,
      textDecoration: 'underline',
      fontColor: '#ff0000',
      backgroundColor: '#00ff00',
      textAlign: 'center',
      textWrap: true,
    });
    wb.setCell(undefined, 0, 0, { styleId });

    const out = await readBack(wb);
    const cell = out.worksheets[0].getCell(1, 1);
    expect(cell.font?.bold).toBe(true);
    expect(cell.font?.underline).toBe(true);
    expect(cell.font?.color?.argb).toBe('FFFF0000');
    expect((cell.fill as ExcelJS.FillPattern).fgColor?.argb).toBe('FF00FF00');
    expect(cell.alignment?.horizontal).toBe('center');
    expect(cell.alignment?.wrapText).toBe(true);
  });

  it('translates a CSS border shorthand to an xlsx border', async () => {
    const wb = freshWorkbook();
    wb.setCellValue(undefined, 0, 0, 'bordered');
    const styleId = wb.getStylePool().getOrCreate({ borderTop: '2px dashed #123456' });
    wb.setCell(undefined, 0, 0, { styleId });

    const out = await readBack(wb);
    const border = out.worksheets[0].getCell(1, 1).border?.top;
    expect(border?.style).toBe('dashed');
    expect(border?.color?.argb).toBe('FF123456');
  });

  it('translates a currency format to an Excel number format code', async () => {
    const wb = freshWorkbook();
    wb.setCellValue(undefined, 0, 0, 1234.5);
    const formatId = wb.getFormatPool().getOrCreate({
      type: 'currency',
      currencyCode: 'USD',
      decimalPlaces: 2,
      useThousandsSeparator: true,
    });
    wb.setCell(undefined, 0, 0, { formatId });

    const out = await readBack(wb);
    expect(out.worksheets[0].getCell(1, 1).numFmt).toBe('"$"#,##0.00');
  });

  it('carries merged regions, frozen panes, and hidden rows/columns', async () => {
    const wb = freshWorkbook();
    wb.setCellValue(undefined, 0, 0, 'merged');
    wb.mergeCells({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
    wb.getSheet().setFreeze(1, 1);
    wb.getSheet().hideRow(2);
    wb.getSheet().hideCol(2);
    // Cells past the merge/freeze/hide so the bounded export loop reaches them.
    wb.setCellValue(undefined, 2, 2, 'x');

    const out = await readBack(wb);
    const ws = out.worksheets[0];
    expect(ws.getCell(2, 2).isMerged).toBe(true);
    expect(ws.views[0]).toMatchObject({ state: 'frozen', xSplit: 1, ySplit: 1 });
    expect(ws.getRow(3).hidden).toBe(true);
    expect(ws.getColumn(3).hidden).toBe(true);
  });

  it('exports every sheet, in order', async () => {
    const wb = freshWorkbook();
    wb.addSheet('Second');
    wb.addSheet('Third');

    const out = await readBack(wb);
    expect(out.worksheets.map((ws) => ws.name)).toEqual(['Sheet1', 'Second', 'Third']);
  });
});

describe('exportToXlsx -> importFromXlsx round-trips', () => {
  it('round-trips plain values, formulas, and every sheet', async () => {
    const wb = freshWorkbook();
    wb.setCellValue(undefined, 0, 0, 'hello');
    wb.setCellValue(undefined, 0, 1, 42);
    wb.setCellValue(undefined, 0, 2, true);
    wb.setFormula(undefined, 1, 0, '=SUM(B1)');
    wb.addSheet('Second');

    const bytes = await exportToXlsx(wb);
    const wb2 = new WorkbookImpl('wb2', 'WB2');
    const { warnings } = await importFromXlsx(bytes, wb2);

    expect(warnings).toEqual([]);
    expect(Array.from(wb2.sheets.values()).map((s) => s.name)).toEqual(['Sheet1', 'Second']);
    expect(wb2.getCellValue(undefined, 0, 0)).toBe('hello');
    expect(wb2.getCellValue(undefined, 0, 1)).toBe(42);
    expect(wb2.getCellValue(undefined, 0, 2)).toBe(true);
    expect(wb2.getCell(undefined, 1, 0)?.formula).toBe('=SUM(B1)');
  });

  it('round-trips style and a currency format', async () => {
    const wb = freshWorkbook();
    wb.setCellValue(undefined, 0, 0, 1234.5);
    const styleId = wb.getStylePool().getOrCreate({ bold: true, fontColor: '#ff0000' });
    const formatId = wb.getFormatPool().getOrCreate({
      type: 'currency', currencyCode: 'USD', decimalPlaces: 2, useThousandsSeparator: true,
    });
    wb.setCell(undefined, 0, 0, { styleId, formatId });

    const bytes = await exportToXlsx(wb);
    const wb2 = new WorkbookImpl('wb2', 'WB2');
    await importFromXlsx(bytes, wb2);

    const cell = wb2.getCell(undefined, 0, 0);
    const style = wb2.getStylePool().get(cell!.styleId!);
    const format = wb2.getFormatPool().get(cell!.formatId!);
    expect(style).toMatchObject({ bold: true, fontColor: '#ff0000' });
    expect(format).toMatchObject({ type: 'currency', currencyCode: 'USD', decimalPlaces: 2, useThousandsSeparator: true });
  });

  it('round-trips merged regions, frozen panes, and hidden rows/columns', async () => {
    const wb = freshWorkbook();
    wb.setCellValue(undefined, 0, 0, 'merged');
    wb.mergeCells({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
    wb.getSheet().setFreeze(1, 1);
    wb.getSheet().hideRow(2);
    wb.getSheet().hideCol(2);
    wb.setCellValue(undefined, 2, 2, 'x');

    const bytes = await exportToXlsx(wb);
    const wb2 = new WorkbookImpl('wb2', 'WB2');
    await importFromXlsx(bytes, wb2);

    const sheet2 = wb2.getSheet();
    expect(sheet2.getMergeAt(0, 0)).toEqual({ startRow: 0, startCol: 0, endRow: 1, endCol: 1 });
    expect(sheet2.getFrozenRows()).toBe(1);
    expect(sheet2.getFrozenCols()).toBe(1);
    expect(sheet2.isRowHidden(2)).toBe(true);
    expect(sheet2.isColHidden(2)).toBe(true);
  });

  it('replaces the target workbook entirely rather than merging into it', async () => {
    const wb = freshWorkbook();
    wb.setCellValue(undefined, 0, 0, 'from file');

    const bytes = await exportToXlsx(wb);
    const target = new WorkbookImpl('target', 'Target');
    target.addSheet('PreExisting');
    target.setCellValue(undefined, 5, 5, 'should be gone');

    await importFromXlsx(bytes, target);

    expect(Array.from(target.sheets.values()).map((s) => s.name)).toEqual(['Sheet1']);
    expect(target.getCellValue(undefined, 0, 0)).toBe('from file');
    expect(target.getCellValue(undefined, 5, 5)).toBeNull();
  });
});
