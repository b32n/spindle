import { exportXlsx } from './writer';
import { importXlsx } from './reader';
import type { SpreadsheetExportInput } from './types';

async function roundTrip(input: SpreadsheetExportInput) {
  const bytes = await exportXlsx(input);
  return importXlsx(bytes);
}

describe('importXlsx', () => {
  it('round-trips plain values of every supported type', async () => {
    const result = await roundTrip({
      sheets: [
        {
          name: 'Sheet1',
          cells: new Map([
            ['0:0', { value: 'hello' }],
            ['0:1', { value: 42 }],
            ['0:2', { value: true }],
          ]),
        },
      ],
    });
    expect(result.warnings).toEqual([]);
    const cells = result.sheets[0].cells;
    expect(cells.get('0:0')?.value).toBe('hello');
    expect(cells.get('0:1')?.value).toBe(42);
    expect(cells.get('0:2')?.value).toBe(true);
  });

  it('round-trips a formula with its cached result', async () => {
    const result = await roundTrip({
      sheets: [{ name: 'S', cells: new Map([['0:0', { value: 5, formula: 'SUM(A2:A3)' }]]) }],
    });
    const cell = result.sheets[0].cells.get('0:0');
    expect(cell?.formula).toBe('SUM(A2:A3)');
    expect(cell?.value).toBe(5);
  });

  it('round-trips a number format code', async () => {
    const result = await roundTrip({
      sheets: [{ name: 'S', cells: new Map([['0:0', { value: 1234.5, numberFormat: { code: '#,##0.00' } }]]) }],
    });
    expect(result.sheets[0].cells.get('0:0')?.numberFormat?.code).toBe('#,##0.00');
  });

  it('does not set a numberFormat for General (default) cells', async () => {
    const result = await roundTrip({ sheets: [{ name: 'S', cells: new Map([['0:0', { value: 1 }]]) }] });
    expect(result.sheets[0].cells.get('0:0')?.numberFormat).toBeUndefined();
  });

  it('round-trips font/fill/alignment style', async () => {
    const result = await roundTrip({
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
                  underline: true,
                  strikethrough: true,
                  fontColor: '#ff0000',
                  backgroundColor: '#00ff00',
                  horizontalAlign: 'center',
                  verticalAlign: 'middle',
                  wrapText: true,
                },
              },
            ],
          ]),
        },
      ],
    });
    expect(result.sheets[0].cells.get('0:0')?.style).toMatchObject({
      bold: true,
      italic: true,
      underline: true,
      strikethrough: true,
      fontColor: '#ff0000',
      backgroundColor: '#00ff00',
      horizontalAlign: 'center',
      verticalAlign: 'middle',
      wrapText: true,
    });
  });

  it('round-trips a border', async () => {
    const result = await roundTrip({
      sheets: [
        { name: 'S', cells: new Map([['0:0', { value: 1, style: { border: { top: { style: 'thick', color: '#123456' } } } }]]) },
      ],
    });
    expect(result.sheets[0].cells.get('0:0')?.style?.border?.top).toEqual({ style: 'thick', color: '#123456' });
  });

  it('round-trips merged ranges', async () => {
    const result = await roundTrip({
      sheets: [
        {
          name: 'S',
          cells: new Map([['0:0', { value: 'merged' }]]),
          mergedRanges: [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }],
        },
      ],
    });
    expect(result.sheets[0].mergedRanges).toEqual([{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }]);
  });

  it('round-trips frozen panes', async () => {
    const result = await roundTrip({ sheets: [{ name: 'S', cells: new Map([['0:0', { value: 1 }]]), frozenRows: 2, frozenCols: 1 }] });
    expect(result.sheets[0].frozenRows).toBe(2);
    expect(result.sheets[0].frozenCols).toBe(1);
  });

  it('round-trips hidden rows/columns', async () => {
    const result = await roundTrip({
      sheets: [
        {
          name: 'S',
          cells: new Map([['0:0', { value: 1 }]]),
          hiddenRows: new Set([0]),
          hiddenCols: new Set([0]),
        },
      ],
    });
    expect(result.sheets[0].hiddenRows?.has(0)).toBe(true);
    expect(result.sheets[0].hiddenCols?.has(0)).toBe(true);
  });

  it('round-trips multiple sheets, in order', async () => {
    const result = await roundTrip({
      sheets: [
        { name: 'First', cells: new Map() },
        { name: 'Second', cells: new Map() },
      ],
    });
    expect(result.sheets.map((s) => s.name)).toEqual(['First', 'Second']);
  });

  it('round-trips a hyperlink', async () => {
    const result = await roundTrip({
      sheets: [{ name: 'S', cells: new Map([['0:0', { value: 'Spindle', hyperlink: 'https://example.com' }]]) }],
    });
    const cell = result.sheets[0].cells.get('0:0');
    expect(cell?.value).toBe('Spindle');
    expect(cell?.hyperlink).toBe('https://example.com');
  });

  it('surfaces a warning list even when empty', async () => {
    const result = await roundTrip({ sheets: [{ name: 'S', cells: new Map() }] });
    expect(result.warnings).toEqual([]);
  });
});
