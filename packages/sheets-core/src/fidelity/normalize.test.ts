import { createHash } from 'node:crypto';
import {
  cellAddress,
  colName,
  countWorkbook,
  diffField,
  digestBinaries,
  stableNumber,
  normalizeWorkbook,
  resolvePools,
  scrubMessage,
} from '../../../../scripts/xlsx-fidelity/normalize';
import type { Json } from '../../../../scripts/fidelity-core/types';
import { WorkbookImpl } from '../workbook';

const sha = (s: string) => createHash('sha256').update(s).digest('hex');

describe('addresses', () => {
  it.each([
    [0, 'A'],
    [25, 'Z'],
    [26, 'AA'],
    [51, 'AZ'],
    [701, 'ZZ'],
    [702, 'AAA'],
    [16383, 'XFD'],
  ])('column %i is %s', (index, name) => {
    expect(colName(index)).toBe(name);
  });
  it('joins column and 1-based row', () => {
    expect(cellAddress(0, 0)).toBe('A1');
    expect(cellAddress(9, 27)).toBe('AB10');
  });
});

describe('stableNumber', () => {
  it('keeps 13 significant digits so last-digit differences between Node versions vanish', () => {
    // financial results as computed on Node 24 and Node 22
    expect(stableNumber(772.1734929184817)).toBe(stableNumber(772.1734929184818));
    expect(stableNumber(0.08896339469334992)).toBe(stableNumber(0.08896339469334985));
    expect(stableNumber(1257.789253554884)).toBe(stableNumber(1257.7892535548842));
    expect(stableNumber(0.04277497803511183)).toBe(stableNumber(0.042774978035111176));
    expect(stableNumber(0.1 + 0.2)).toBe(0.3);
    expect(stableNumber(1.5)).toBe(1.5);
    expect(stableNumber(0)).toBe(0);
  });
  it('does not make values that differ in the first 13 digits equal', () => {
    expect(stableNumber(1.0000001)).not.toBe(stableNumber(1.0000002));
    expect(stableNumber(1234567890123)).not.toBe(stableNumber(1234567890124));
    expect(stableNumber(0.1234567890123)).not.toBe(stableNumber(0.1234567890124));
  });
});

describe('digestBinaries', () => {
  it('replaces bytes and data URLs by size and digest', () => {
    const bytes = new Uint8Array([1, 2, 3]);
    expect(digestBinaries(bytes)).toEqual({ bytes: 3, sha256: sha('\u0001\u0002\u0003') });
    expect(digestBinaries(bytes.buffer)).toEqual({ bytes: 3, sha256: sha('\u0001\u0002\u0003') });
    expect(digestBinaries(Buffer.from('hi'))).toEqual({ bytes: 2, sha256: sha('hi') });
    const url = `data:image/png;base64,${Buffer.from('hi').toString('base64')}`;
    expect(digestBinaries(url)).toEqual({ dataUrl: 'image/png', bytes: 2, sha256: sha('hi') });
    expect(digestBinaries(`data:;base64,${Buffer.from('hi').toString('base64')}`)).toMatchObject({
      dataUrl: '',
      bytes: 2,
    });
    expect(
      digestBinaries(`data:text/plain;charset=utf-8;base64,${Buffer.from('hi').toString('base64')}`)
    ).toMatchObject({ dataUrl: 'text/plain' });
  });

  it('keeps everything else, recursing into arrays and objects', () => {
    expect(digestBinaries('data:text/plain,not base64')).toBe('data:text/plain,not base64');
    expect(
      digestBinaries({ a: [1, 'x', null, true], b: undefined, c: { d: new Uint8Array(0) } })
    ).toEqual({
      a: [1, 'x', null, true],
      c: { d: { bytes: 0, sha256: sha('') } },
    });
    expect(digestBinaries(undefined)).toBeNull();
    expect(digestBinaries(null)).toBeNull();
  });

  it('turns dates into ISO strings and maps and sets into arrays', () => {
    expect(digestBinaries({ d: new Date(Date.UTC(2026, 0, 15, 9)), bad: new Date(NaN) })).toEqual({
      d: '2026-01-15T09:00:00.000Z',
      bad: 'Invalid Date',
    });
    expect(digestBinaries(new Map([['k', new Uint8Array(0)]]))).toEqual([
      ['k', { bytes: 0, sha256: sha('') }],
    ]);
    expect(digestBinaries(new Set([1, 'a']))).toEqual([1, 'a']);
  });

  it('writes non-finite numbers as text so JSON keeps them', () => {
    expect(digestBinaries([Infinity, -Infinity, NaN, 1.5])).toEqual([
      '#number:Infinity',
      '#number:-Infinity',
      '#number:NaN',
      1.5,
    ]);
  });
});

const data = (over: Record<string, unknown> = {}) => ({
  id: 'wb',
  name: 'fixture',
  activeSheetId: 's2',
  defaultRowHeight: 20,
  defaultColWidth: 100,
  stylePool: { sty_b: { bold: true }, sty_a: { italic: true } },
  formatPool: { fmt_1: { type: 'percentage', decimalPlaces: 1 } },
  sheets: [
    {
      id: 's1',
      name: 'One',
      rowCount: 1000,
      colCount: 100,
      rowOrder: [
        [0, 'r0'],
        [1, 'r1'],
        [4, 'r4'],
      ],
      colOrder: [
        [0, 'c0'],
        [1, 'c1'],
        [27, 'c27'],
      ],
      cells: [
        { key: 'r4:c27', cell: { value: 'z', styleId: 'sty_a' } },
        {
          key: 'r0:c0',
          cell: {
            value: 1,
            formula: '=1+0',
            formulaAst: { type: 'big', ids: ['x'] },
            styleId: 'sty_b',
            formatId: 'fmt_1',
          },
        },
        { key: 'r0:c1', cell: { value: null, styleId: 'sty_b', hyperlink: 'https://example.com' } },
        { key: 'r1:c0', cell: { value: Infinity } },
        { key: 'gone:c0', cell: { value: 'no row' } },
        { key: 'r0:gone', cell: { value: 'no col' } },
        { key: 'r1:c1', cell: { value: 'x', styleId: 'missing', formatId: 'missing' } },
      ],
      config: {
        defaultRowHeight: 18,
        defaultColWidth: 90,
        rowHeights: [
          ['r4', 30],
          ['r0', 25],
          ['unknown', 1],
        ],
        colWidths: [
          ['c27', 120],
          ['c0', 80],
          ['unknown', 1],
        ],
        hiddenRows: ['r4', 'r1', 'unknown'],
        hiddenCols: ['c27', 'c1', 'unknown'],
        frozenRows: 1,
        frozenCols: 2,
        showGridLines: false,
        sortOrder: [{ column: 1, direction: 'desc' }],
        filters: [
          ['c1', { column: 1, type: 'text', criteria: { type: 'equals', value: 'x' } }],
          ['unknown', {}],
        ],
        mergedRegions: [
          { startRowId: 'r1', startColId: 'c0', endRowId: 'r1', endColId: 'c1' },
          { startRowId: 'r0', startColId: 'c0', endRowId: 'r1', endColId: 'c1' },
          { startRowId: 'r0', startColId: 'c0', endRowId: 'r0', endColId: 'c27' },
          { startRowId: 'r0', startColId: 'c1', endRowId: 'r0', endColId: 'c1' },
          { startRowId: 'r0', startColId: 'c0', endRowId: 'r0', endColId: 'c1' },
          { startRowId: 'unknown', startColId: 'c0', endRowId: 'r0', endColId: 'c0' },
        ],
      },
      threads: [
        {
          id: 't1',
          anchor: { rowId: 'r0', colId: 'c0' },
          status: 'open',
          createdAt: '2026-01-01T00:00:00Z',
          comments: [{ id: 'c1', authorId: 'u1', text: 'hello', createdAt: 'x' }],
        },
        { id: 't2', anchor: { rowId: 'gone', colId: 'c0' }, status: 'resolved', comments: [] },
      ],
    },
    { id: 's2', name: 'Two', rowCount: 5, colCount: 5, cells: [], config: {} },
  ],
  ...over,
});

describe('normalizeWorkbook', () => {
  const model = normalizeWorkbook(data()) as Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
  const one = model.sheets[0];

  it('addresses cells by position, in row-major order, dropping the formula AST', () => {
    expect(Object.keys(one.cells)).toEqual(['A1', 'B1', 'A2', 'B2', 'AB5']);
    expect(one.cells.A1).toEqual({ value: 1, formula: '=1+0', style: 's1', format: 'f1' });
    expect(one.cells.B1).toEqual({ value: null, style: 's1', hyperlink: 'https://example.com' });
    expect(one.cells.A2.value).toBe('#number:Infinity');
  });

  it('numbers pooled styles and formats by first use and marks unresolved ones', () => {
    expect(model.styles).toEqual({
      s1: { bold: true },
      s2: { unresolved: true },
      s3: { italic: true },
    });
    expect(one.cells.B2).toMatchObject({ style: 's2', format: 'f2' });
    expect(one.cells.AB5.style).toBe('s3');
    expect(model.formats).toEqual({
      f1: { type: 'percentage', decimalPlaces: 1 },
      f2: { unresolved: true },
    });
  });

  it('turns row and column ids into 1-based rows and letters, sorted, skipping unknown ids', () => {
    expect(one.config).toMatchObject({
      frozenRows: 1,
      frozenCols: 2,
      showGridLines: false,
      rowHeights: { '1': 25, '5': 30 },
      colWidths: { A: 80, AB: 120 },
      hiddenRows: [2, 5],
      hiddenCols: ['B', 'AB'],
      defaultRowHeight: 18,
      defaultColWidth: 90,
      sortOrder: [{ column: 1, direction: 'desc' }],
    });
    expect(one.config.filters).toEqual({
      B: { column: 1, criteria: { type: 'equals', value: 'x' }, type: 'text' },
    });
  });

  it('sorts merged regions and skips those with unknown corners', () => {
    expect(one.config.merges).toEqual(['A1:B1', 'A1:AB1', 'A1:B2', 'B1:B1', 'A2:B2']);
  });

  it('drops volatile thread fields and marks orphaned threads', () => {
    expect(one.threads).toEqual([
      { anchor: 'A1', status: 'open', comments: [{ text: 'hello' }] },
      { anchor: 'orphaned', status: 'resolved', comments: [] },
    ]);
  });

  it('resolves the active sheet to an index and keeps sheet sizes and defaults', () => {
    expect(model).toMatchObject({ activeSheet: 1, defaultRowHeight: 20, defaultColWidth: 100 });
    expect(one).toMatchObject({ name: 'One', rowCount: 1000, colCount: 100 });
    expect(model.sheets[1].config).toEqual({
      frozenRows: 0,
      frozenCols: 0,
      showGridLines: true,
      rowHeights: {},
      colWidths: {},
      hiddenRows: [],
      hiddenCols: [],
      merges: [],
      filters: {},
      sortOrder: [],
    });
  });

  it('falls back to the first sheet when the active id is unknown', () => {
    expect(normalizeWorkbook(data({ activeSheetId: 'nope' })).activeSheet).toBe(0);
  });

  it('survives malformed input instead of throwing', () => {
    for (const bad of [null, undefined, 5, 'x', [], { sheets: 'no' }, { sheets: [null, 5, {}] }]) {
      expect(() => normalizeWorkbook(bad)).not.toThrow();
    }
    expect(normalizeWorkbook(null)).toMatchObject({
      sheets: [],
      styles: {},
      formats: {},
      activeSheet: 0,
    });
    const odd = normalizeWorkbook({
      sheets: [
        {
          id: 'x',
          name: 'N',
          rowOrder: 'no',
          colOrder: [[0, 'c']],
          cells: [{ key: 'only-one-part', cell: 3 }, null],
          config: { mergedRegions: [null] },
          threads: [null, { anchor: null }],
        },
      ],
    });
    expect((odd.sheets as Json[])[0]).toMatchObject({
      name: 'N',
      cells: {},
      threads: [{ anchor: 'orphaned' }, { anchor: 'orphaned' }],
    });
  });

  it('keeps literal numbers exact and rounds only formula results to 13 digits', () => {
    const m = normalizeWorkbook({
      sheets: [
        {
          id: 's',
          name: 'S',
          rowOrder: [[0, 'r']],
          colOrder: [
            [0, 'a'],
            [1, 'b'],
            [2, 'c'],
            [3, 'd'],
            [4, 'e'],
          ],
          cells: [
            { key: 'r:a', cell: { value: 1234567890123 } },
            { key: 'r:b', cell: { value: 0.123456789012345 } },
            { key: 'r:c', cell: { value: 772.1734929184817, formula: '=PV(1)' } },
            { key: 'r:d', cell: { value: 'text', formula: '=A1&""' } },
            { key: 'r:e', cell: { value: 1234567890123456, formula: '=1' } },
          ],
          config: {},
        },
      ],
    }) as { sheets: Array<{ cells: Record<string, { value: unknown }> }> };
    const v = Object.fromEntries(Object.entries(m.sheets[0].cells).map(([k, c]) => [k, c.value]));
    expect(v).toEqual({
      A1: 1234567890123,
      B1: 0.123456789012345,
      C1: 772.1734929185,
      D1: 'text',
      E1: 1234567890123000,
    });
  });

  it('gives the same model for the same workbook built twice (ids are random)', () => {
    const build = () => {
      const wb = new WorkbookImpl('id', 'name');
      wb.setCellValue(undefined, 0, 0, 'a');
      wb.setFormula(undefined, 1, 0, '=1+2');
      wb.mergeCells({ startRow: 3, startCol: 0, endRow: 4, endCol: 1 });
      return normalizeWorkbook(wb.getData());
    };
    const a = build();
    expect(build()).toEqual(a);
    expect(countWorkbook(a)).toMatchObject({ sheets: 1, cells: 2, formulas: 1, merges: 1 });
  });
});

describe('resolvePools', () => {
  it('inlines styles and formats into cells and drops the pools', () => {
    const resolved = resolvePools(normalizeWorkbook(data())) as Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
    expect(resolved.styles).toBeUndefined();
    expect(resolved.formats).toBeUndefined();
    expect(resolved.sheets[0].cells.A1).toMatchObject({
      style: { bold: true },
      format: { type: 'percentage', decimalPlaces: 1 },
    });
    expect(resolved.sheets[0].cells.B2).toMatchObject({
      style: { unresolved: true },
      format: { unresolved: true },
    });
    expect(resolved.sheets[1].cells).toEqual({});
  });

  it('leaves malformed models alone', () => {
    expect(resolvePools(5)).toBe(5);
    expect(resolvePools(null)).toBeNull();
    expect(
      resolvePools({
        sheets: [
          null,
          { cells: 5 },
          { cells: { A1: 3, B1: { value: 1, style: 7 }, C1: { style: 's9' } } },
        ],
      })
    ).toEqual({
      sheets: [
        null,
        { cells: 5 },
        { cells: { A1: 3, B1: { value: 1, style: 7 }, C1: { style: { unresolved: true } } } },
      ],
    });
    expect(resolvePools({ styles: 5, formats: 'x' })).toEqual({ sheets: [] });
  });
});

describe('countWorkbook', () => {
  it('counts what a model holds and survives malformed models', () => {
    expect(countWorkbook(normalizeWorkbook(data()))).toEqual({
      sheets: 2,
      cells: 5,
      formulas: 1,
      merges: 5,
      styles: 3,
      formats: 2,
    });
    expect(countWorkbook({ sheets: 'x' })).toMatchObject({ sheets: 0 });
    expect(countWorkbook(null)).toEqual({
      sheets: 0,
      cells: 0,
      formulas: 0,
      merges: 0,
      styles: 0,
      formats: 0,
    });
    expect(
      countWorkbook({
        sheets: [5, {}, { cells: { A1: 1, B1: { formula: '=1' } }, config: { merges: 'x' } }],
        styles: 'x',
        formats: [],
      })
    ).toEqual({
      sheets: 3,
      cells: 2,
      formulas: 1,
      merges: 0,
      styles: 0,
      formats: 0,
    });
  });
});

describe('diffField', () => {
  it('folds indexes, cell addresses and row or column keys', () => {
    expect(diffField('sheets[0].cells.AB12.value')).toBe('sheets[].cells.*.value');
    expect(diffField('sheets[2].config.rowHeights.15')).toBe('sheets[].config.rowHeights.*');
    expect(diffField('sheets[2].config.colWidths.AA')).toBe('sheets[].config.colWidths.*');
    expect(diffField('sheets[2].config.filters.B')).toBe('sheets[].config.filters.*');
    expect(diffField('sheets[1].name')).toBe('sheets[].name');
  });
});

describe('scrubMessage', () => {
  it('replaces repository paths, longest first, and ignores empty roots', () => {
    expect(
      scrubMessage('failed in /work/repo/scripts/x.ts and /work/repo', ['/work', '/work/repo', ''])
    ).toBe('failed in <repo>/scripts/x.ts and <repo>');
    expect(scrubMessage('nothing here', [])).toBe('nothing here');
  });
});
