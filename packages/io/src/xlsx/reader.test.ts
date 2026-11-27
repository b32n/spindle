import { exportXlsx } from './writer';
import JSZip from 'jszip';
import { importXlsx } from './reader';
import { ImportError, type ImportErrorCode } from '../common/import-error';
import type { SpreadsheetExportInput } from '@b32nio/spindle-shared/export/xlsx';

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

describe('importXlsx: the only error is an ImportError', () => {
  const ole = (...streams: string[]) =>
    Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...Array(512).fill(0), ...streams.flatMap((s) => [...Array.from(s).flatMap((c) => [c.charCodeAt(0), 0]), 0, 0])]);
  const zipOf = async (files: Record<string, string>) => {
    const zip = new JSZip();
    for (const [name, content] of Object.entries(files)) zip.file(name, content);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  };
  const codeOf = async (bytes: Uint8Array, options?: Parameters<typeof importXlsx>[1]): Promise<ImportErrorCode | 'no error'> =>
    importXlsx(bytes, options).then(
      () => 'no error',
      (e: unknown) => {
        expect(e).toBeInstanceOf(ImportError);
        return (e as ImportError).code;
      }
    );

  it.each<[string, () => Promise<Uint8Array> | Uint8Array, ImportErrorCode]>([
    ['an empty file', () => new Uint8Array(0), 'not-a-package'],
    ['text', () => new TextEncoder().encode('a,b,c'), 'not-a-package'],
    ['an encrypted package', () => ole('EncryptionInfo', 'EncryptedPackage'), 'encrypted'],
    ['a legacy .xls', () => ole('Workbook'), 'legacy-binary'],
    ['a zip that is not a workbook', () => zipOf({ 'hello.txt': 'hi' }), 'not-a-package'],
    ['a package with an unreadable workbook part', () => zipOf({ '[Content_Types].xml': '<Types/>', 'xl/workbook.xml': '<<<' }), 'corrupt'],
  ])('%s', async (_name, make, code) => {
    expect(await codeOf(await make())).toBe(code);
  });

  it('refuses a forged or corrupt mimetype entry as corrupt (an ImportError, not a plain Error)', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types/>');
    zip.file('mimetype', new Uint8Array(1024 * 1024));
    const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    const forged = bytes.slice();
    const view = new DataView(forged.buffer);
    for (let i = 0; i + 30 < forged.length; i++) {
      if (view.getUint32(i, true) === 0x04034b50 && new TextDecoder().decode(forged.subarray(i + 30, i + 38)) === 'mimetype') view.setUint32(i + 22, 10, true);
      if (view.getUint32(i, true) === 0x02014b50 && new TextDecoder().decode(forged.subarray(i + 46, i + 54)) === 'mimetype') view.setUint32(i + 24, 10, true);
    }
    expect(await codeOf(forged)).toBe('corrupt');
  });

  it('refuses a truncated workbook as corrupt', async () => {
    const bytes = await exportXlsx({ sheets: [{ name: 'S', cells: new Map([['0:0', { value: 'x'.repeat(5000) }]]) }] });
    expect(await codeOf(bytes.subarray(0, Math.floor(bytes.length / 2)))).toBe('corrupt');
  });

  it('applies the limits it is given, and the defaults otherwise', async () => {
    const bytes = await exportXlsx({ sheets: [{ name: 'S', cells: new Map([['0:0', { value: 1 }]]) }] });
    expect(await codeOf(bytes)).toBe('no error');
    expect(await codeOf(bytes, { limits: { maxEntries: 2 } })).toBe('size-limit');
    expect(await codeOf(bytes, { limits: { maxPackageBytes: 10 } })).toBe('size-limit');
  });

  it('refuses a zip bomb with a small header declaration without inflating it', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types/>');
    zip.file('xl/worksheets/sheet1.xml', new Uint8Array(40 * 1024 * 1024));
    const bomb = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    expect(bomb.length).toBeLessThan(200 * 1024);
    expect(await codeOf(bomb)).toBe('size-limit');
  });
});
