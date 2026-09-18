import ExcelJS from 'exceljs';
import { exportXlsx } from './writer';
import type { SpreadsheetExportInput } from './types';

async function readBack(input: SpreadsheetExportInput): Promise<ExcelJS.Workbook> {
  const bytes = await exportXlsx(input);
  const workbook = new ExcelJS.Workbook();
  // @types/node's newer generic Buffer<T> clashes with exceljs's older
  // non-generic Buffer typing here; it's a real Buffer at runtime either way.
  // @ts-expect-error -- Buffer<ArrayBuffer> vs exceljs's plain Buffer type
  await workbook.xlsx.load(Buffer.from(bytes));
  return workbook;
}

describe('exportXlsx', () => {
  it('writes plain values of every supported type', async () => {
    const workbook = await readBack({
      sheets: [
        {
          name: 'Sheet1',
          cells: new Map([
            ['0:0', { value: 'hello' }],
            ['0:1', { value: 42 }],
            ['0:2', { value: true }],
            ['1:0', { value: null }],
          ]),
        },
      ],
    });
    const ws = workbook.worksheets[0];
    expect(ws.name).toBe('Sheet1');
    expect(ws.getCell(1, 1).value).toBe('hello');
    expect(ws.getCell(1, 2).value).toBe(42);
    expect(ws.getCell(1, 3).value).toBe(true);
  });

  it('writes a formula with its cached result', async () => {
    const workbook = await readBack({
      sheets: [{ name: 'S', cells: new Map([['0:0', { value: 3, formula: 'SUM(A2:A3)' }]]) }],
    });
    const cell = workbook.worksheets[0].getCell(1, 1);
    expect(cell.formula).toBe('SUM(A2:A3)');
    expect(cell.result).toBe(3);
  });

  it('writes a number format code', async () => {
    const workbook = await readBack({
      sheets: [{ name: 'S', cells: new Map([['0:0', { value: 1234.5, numberFormat: { code: '#,##0.00' } }]]) }],
    });
    expect(workbook.worksheets[0].getCell(1, 1).numFmt).toBe('#,##0.00');
  });

  it('writes font/fill/alignment style', async () => {
    const workbook = await readBack({
      sheets: [
        {
          name: 'S',
          cells: new Map([
            [
              '0:0',
              {
                value: 'styled',
                style: {
                  bold: true,
                  italic: true,
                  fontColor: '#ff0000',
                  backgroundColor: '#00ff00',
                  horizontalAlign: 'center',
                  wrapText: true,
                },
              },
            ],
          ]),
        },
      ],
    });
    const cell = workbook.worksheets[0].getCell(1, 1);
    expect(cell.font?.bold).toBe(true);
    expect(cell.font?.italic).toBe(true);
    expect(cell.font?.color?.argb).toBe('FFFF0000');
    expect((cell.fill as ExcelJS.FillPattern).fgColor?.argb).toBe('FF00FF00');
    expect(cell.alignment?.horizontal).toBe('center');
    expect(cell.alignment?.wrapText).toBe(true);
  });

  it('writes borders', async () => {
    const workbook = await readBack({
      sheets: [
        {
          name: 'S',
          cells: new Map([
            ['0:0', { value: 1, style: { border: { top: { style: 'thick', color: '#000000' } } } }],
          ]),
        },
      ],
    });
    const border = workbook.worksheets[0].getCell(1, 1).border?.top;
    expect(border?.style).toBe('thick');
    expect(border?.color?.argb).toBe('FF000000');
  });

  it('writes merged ranges', async () => {
    const workbook = await readBack({
      sheets: [
        {
          name: 'S',
          cells: new Map([['0:0', { value: 'merged' }]]),
          mergedRanges: [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }],
        },
      ],
    });
    expect(workbook.worksheets[0].getCell(2, 2).isMerged).toBe(true);
  });

  it('writes frozen panes', async () => {
    const workbook = await readBack({
      sheets: [{ name: 'S', cells: new Map(), frozenRows: 1, frozenCols: 2 }],
    });
    const view = workbook.worksheets[0].views[0];
    expect(view).toMatchObject({ state: 'frozen', xSplit: 2, ySplit: 1 });
  });

  it('writes hidden rows/columns', async () => {
    const workbook = await readBack({
      sheets: [
        {
          name: 'S',
          cells: new Map([['0:0', { value: 1 }]]),
          hiddenRows: new Set([0]),
          hiddenCols: new Set([0]),
        },
      ],
    });
    const ws = workbook.worksheets[0];
    expect(ws.getRow(1).hidden).toBe(true);
    expect(ws.getColumn(1).hidden).toBe(true);
  });

  it('writes multiple sheets in order', async () => {
    const workbook = await readBack({
      sheets: [
        { name: 'First', cells: new Map() },
        { name: 'Second', cells: new Map() },
      ],
    });
    expect(workbook.worksheets.map((ws) => ws.name)).toEqual(['First', 'Second']);
  });

  it('writes a hyperlink', async () => {
    const workbook = await readBack({
      sheets: [{ name: 'S', cells: new Map([['0:0', { value: 'Spindle', hyperlink: 'https://example.com' }]]) }],
    });
    const cell = workbook.worksheets[0].getCell(1, 1);
    expect(cell.text).toBe('Spindle');
    expect((cell.value as ExcelJS.CellHyperlinkValue).hyperlink).toBe('https://example.com');
  });
});
