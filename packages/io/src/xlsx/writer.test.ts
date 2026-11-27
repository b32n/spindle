import ExcelJS from 'exceljs';
import { exportXlsx } from './writer';
import type { SpreadsheetExportInput } from '@b32nio/spindle-shared/export/xlsx';

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
  it('opens on the active sheet', async () => {
    const sheets = ['A', 'B', 'C'].map((name) => ({ name, cells: new Map([['0:0', { value: name }]]) }));
    const out = await readBack({ sheets, activeSheetIndex: 2 });
    expect(out.views[0].activeTab).toBe(2);
  });

  it('defaults to the first sheet and clamps an out-of-range index', async () => {
    const sheets = ['A', 'B'].map((name) => ({ name, cells: new Map([['0:0', { value: name }]]) }));
    expect((await readBack({ sheets })).views[0].activeTab).toBe(0);
    expect((await readBack({ sheets, activeSheetIndex: 9 })).views[0].activeTab).toBe(1);
  });

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

  it('writes the prefixes the file format needs on a formula, unless it is cached-only (written as it came)', async () => {
    const workbook = await readBack({
      sheets: [
        {
          name: 'S',
          cells: new Map([
            ['0:0', { value: 'x', formula: 'IFS(A2>1,"x",TRUE,"y")' }],
            ['0:1', { value: 'x', formula: '@A1:A2' }],
            ['0:2', { value: 1, formula: 'LET(x,1,x)' }],
            ['0:3', { value: 1, formula: '_xlfn.LET(_xlpm.x,1,_xlpm.x)', cachedOnly: true as const }],
            ['0:4', { value: 1, formula: 'IFS(1,2)', cachedOnly: true as const }],
          ]),
        },
      ],
    });
    const ws = workbook.worksheets[0];
    expect(ws.getCell(1, 1).formula).toBe('_xlfn.IFS(A2>1,"x",TRUE,"y")');
    expect(ws.getCell(1, 2).formula).toBe('_xlfn.SINGLE(A1:A2)');
    expect(ws.getCell(1, 3).formula).toBe('_xlfn.LET(_xlpm.x,1,_xlpm.x)');
    expect(ws.getCell(1, 4).formula).toBe('_xlfn.LET(_xlpm.x,1,_xlpm.x)');
    expect(ws.getCell(1, 5).formula).toBe('IFS(1,2)');
  });

  it('writes an array formula over its range (and ignores an extent that is not an A1 range)', async () => {
    const bytes = await exportXlsx({
      sheets: [
        {
          name: 'S',
          cells: new Map([
            ['0:0', { value: 1, formula: 'SORT(B1:B3)', arrayRef: 'A1:A3' }],
            ['1:0', { value: 2 }],
            ['0:2', { value: 3, formula: 'SUM(B1:B2*B1:B2)', arrayRef: 'C1' }],
            ['0:3', { value: 4, formula: 'A1+1', arrayRef: 'A1:A2"><x' }],
          ]),
        },
      ],
    });
    const zip = await (await import('jszip')).default.loadAsync(bytes);
    const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(xml).toContain('<f t="array" ref="A1:A3">_xlfn._xlws.SORT(B1:B3)</f>');
    expect(xml).toContain('<f t="array" ref="C1">SUM(B1:B2*B1:B2)</f>');
    expect(xml).toContain('<f>A1+1</f>');
    expect(xml).not.toContain(' cm=');
    expect(zip.file('xl/metadata.xml')).toBeNull();
  });

  it('writes a dynamic-array formula with the cell metadata Excel marks it with, and reads it back as one', async () => {
    const JSZip = (await import('jszip')).default;
    const bytes = await exportXlsx({
      sheets: [
        { name: 'A', cells: new Map([['0:0', { value: 1, formula: 'SORT(B1:B3)', arrayRef: 'A1:A3', dynamicArray: true as const }], ['0:1', { value: 2, formula: 'SUM(1,1)', arrayRef: 'B1' }]]) },
        { name: 'B', cells: new Map([['0:0', { value: 5, formula: 'SEQUENCE(2)', arrayRef: 'A1:A2', dynamicArray: true as const }]]) },
      ],
    });
    const zip = await JSZip.loadAsync(bytes);
    const sheet1 = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    const sheet2 = await zip.file('xl/worksheets/sheet2.xml')!.async('string');
    expect(sheet1).toMatch(/<c r="A1" cm="1"[^>]*><f t="array" ref="A1:A3">_xlfn\._xlws\.SORT\(B1:B3\)<\/f>/);
    expect(sheet1).not.toMatch(/<c r="B1" cm=/);
    expect(sheet2).toMatch(/<c r="A1" cm="1"/);
    expect(await zip.file('xl/metadata.xml')!.async('string')).toContain('XLDAPR');
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('/xl/metadata.xml');
    expect(await zip.file('xl/_rels/workbook.xml.rels')!.async('string')).toContain('sheetMetadata');
    const { importXlsx } = await import('./reader');
    const back = await importXlsx(bytes);
    expect(back.sheets[0].cells.get('0:0')?.formulaInfo).toEqual({ type: 'array', ref: 'A1:A3', dynamic: true });
    expect(back.sheets[0].cells.get('0:1')?.formulaInfo).toEqual({ type: 'array', ref: 'B1' });
    expect(back.sheets[1].cells.get('0:0')?.formulaInfo).toEqual({ type: 'array', ref: 'A1:A2', dynamic: true });
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
