import {
  evenFractions, makeGrid, insertRow, insertColumn, removeRow, removeColumn, mergeCells, splitCell, setCellBorders,
  resolveCellEdges, spanRects, spanAt, expandRange, normalizeSpans, applySpanRects,
} from './table';
import { createTableElement } from './elements';
import { richTextFromPlainText, richTextToPlainText } from '../text/model';
import type { Stroke, TableElement } from './types';

function table(rows: number, cols: number): TableElement {
  return {
    id: 't', containerId: 's', index: 'a', type: 'table',
    x: 0, y: 0, w: 600, h: 300, rotation: 0,
    rows, cols, colFractions: evenFractions(cols), rowFractions: evenFractions(rows), cells: makeGrid(rows, cols),
  };
}
const apply = (t: TableElement, p: Partial<TableElement>): TableElement => ({ ...t, ...p }) as TableElement;
const withText = (t: TableElement, r: number, c: number, text: string): TableElement => {
  const cells = t.cells.map((row) => row.slice());
  cells[r][c] = { ...cells[r][c], richText: richTextFromPlainText(text) };
  return { ...t, cells };
};
const text = (t: TableElement, r: number, c: number) => richTextToPlainText(t.cells[r][c].richText);
const RED: Stroke = { color: { kind: 'rgb', hex: '#ff0000' }, width: 3 };
const GRID: Stroke = { color: { kind: 'rgb', hex: '#000000' }, width: 1 };
const flags = (t: TableElement) => t.cells.map((row) => row.map((c) => `${c.gridSpan ?? ''}${c.rowSpan ? 'x' + c.rowSpan : ''}${c.hMerge ? 'h' : ''}${c.vMerge ? 'v' : ''}.`).join(''));

describe('mergeCells / splitCell', () => {
  it('a single cell range is a no-op', () => {
    expect(mergeCells(table(2, 2), { r0: 1, c0: 1, r1: 1, c1: 1 })).toEqual({});
  });
  it('merges a 2x2 block with OOXML flags and joins text', () => {
    let t = withText(withText(table(3, 3), 0, 0, 'a'), 0, 1, 'b');
    t = withText(t, 1, 1, 'c');
    const m = apply(t, mergeCells(t, { r0: 0, c0: 0, r1: 1, c1: 1 }));
    expect(m.cells[0][0].gridSpan).toBe(2);
    expect(m.cells[0][0].rowSpan).toBe(2);
    expect(m.cells[0][1].hMerge).toBe(true);
    expect(m.cells[1][0].vMerge).toBe(true);
    expect(m.cells[1][1]).toMatchObject({ hMerge: true, vMerge: true });
    expect(text(m, 0, 0)).toBe('a\nb\nc');
    expect(text(m, 0, 1)).toBe('');
    expect(spanRects(m)).toEqual([{ r: 0, c: 0, rs: 2, cs: 2 }]);
  });
  it('an empty origin takes the first non-empty text; reversed ranges work', () => {
    const t = withText(table(1, 3), 0, 2, 'z');
    const m = apply(t, mergeCells(t, { r0: 0, c0: 2, r1: 0, c1: 0 }));
    expect(text(m, 0, 0)).toBe('z');
    expect(m.cells[0][0].gridSpan).toBe(3);
  });
  it('keeps covered cells fill and borders so a split restores them', () => {
    let t = table(1, 2);
    t = { ...t, cells: [[t.cells[0][0], { ...t.cells[0][1], fill: { kind: 'solid', color: { kind: 'rgb', hex: '#00ff00' } }, borders: { r: RED } }]] };
    const m = apply(t, mergeCells(t, { r0: 0, c0: 0, r1: 0, c1: 1 }));
    const s = apply(m, splitCell(m, 0, 1));
    expect(s.cells[0][1].fill).toBeDefined();
    expect(s.cells[0][1].borders).toEqual({ r: RED });
    expect(s.cells[0][0].gridSpan).toBeUndefined();
    expect(s.cells[0][1].hMerge).toBeUndefined();
  });
  it('a range cutting a merged region expands to take all of it', () => {
    let t = table(3, 3);
    t = apply(t, mergeCells(t, { r0: 0, c0: 0, r1: 0, c1: 1 }));
    const m = apply(t, mergeCells(t, { r0: 0, c0: 1, r1: 1, c1: 1 }));
    expect(m.cells[0][0]).toMatchObject({ gridSpan: 2, rowSpan: 2 });
    expect(spanRects(m)).toHaveLength(1);
  });
  it('merging over an existing region replaces it', () => {
    let t = table(2, 3);
    t = apply(t, mergeCells(t, { r0: 0, c0: 0, r1: 0, c1: 1 }));
    t = apply(t, mergeCells(t, { r0: 0, c0: 0, r1: 1, c1: 2 }));
    expect(spanRects(t)).toEqual([{ r: 0, c: 0, rs: 2, cs: 3 }]);
  });
  it('split of an unmerged cell is a no-op; split removes only that region', () => {
    let t = table(2, 4);
    expect(splitCell(t, 0, 0)).toEqual({});
    t = apply(t, mergeCells(t, { r0: 0, c0: 0, r1: 0, c1: 1 }));
    t = apply(t, mergeCells(t, { r0: 0, c0: 2, r1: 1, c1: 3 }));
    const s = apply(t, splitCell(t, 0, 1));
    expect(spanRects(s)).toEqual([{ r: 0, c: 2, rs: 2, cs: 2 }]);
    expect(spanAt(s, 1, 3)).toEqual({ r: 0, c: 2, rs: 2, cs: 2 });
    expect(spanAt(s, 1, 0)).toEqual({ r: 1, c: 0, rs: 1, cs: 1 });
  });
});

describe('spans survive structural edits', () => {
  const merged = () => { const t = table(4, 4); return apply(t, mergeCells(t, { r0: 1, c0: 1, r1: 2, c1: 2 })); };
  it('inserting inside a region grows it; at its edge shifts it', () => {
    const t = merged();
    expect(spanRects(apply(t, insertRow(t, 2)))).toEqual([{ r: 1, c: 1, rs: 3, cs: 2 }]);
    expect(spanRects(apply(t, insertRow(t, 1)))).toEqual([{ r: 2, c: 1, rs: 2, cs: 2 }]);
    expect(spanRects(apply(t, insertRow(t, 3)))).toEqual([{ r: 1, c: 1, rs: 2, cs: 2 }]);
    expect(spanRects(apply(t, insertColumn(t, 2)))).toEqual([{ r: 1, c: 1, rs: 2, cs: 3 }]);
    expect(spanRects(apply(t, insertColumn(t, 1)))).toEqual([{ r: 1, c: 2, rs: 2, cs: 2 }]);
    expect(spanRects(apply(t, insertColumn(t, 3)))).toEqual([{ r: 1, c: 1, rs: 2, cs: 2 }]);
    expect(spanRects(apply(t, insertColumn(t, 0)))).toEqual([{ r: 1, c: 2, rs: 2, cs: 2 }]);
  });
  it('removing shrinks a region, keeps its content and drops it at size one', () => {
    const t = withText(merged(), 1, 1, 'keep');
    const noFirst = apply(t, removeRow(t, 1));
    expect(spanRects(noFirst)).toEqual([{ r: 1, c: 1, rs: 1, cs: 2 }]);
    expect(text(noFirst, 1, 1)).toBe('keep');
    const noSecond = apply(t, removeRow(t, 2));
    expect(spanRects(noSecond)).toEqual([{ r: 1, c: 1, rs: 1, cs: 2 }]);
    const noCol = apply(t, removeColumn(t, 1));
    expect(spanRects(noCol)).toEqual([{ r: 1, c: 1, rs: 2, cs: 1 }]);
    expect(text(noCol, 1, 1)).toBe('keep');
    expect(spanRects(apply(t, removeColumn(t, 0)))).toEqual([{ r: 1, c: 0, rs: 2, cs: 2 }]);
    expect(spanRects(apply(t, removeColumn(t, 3)))).toEqual([{ r: 1, c: 1, rs: 2, cs: 2 }]);
    expect(spanRects(apply(t, removeRow(t, 0)))).toEqual([{ r: 0, c: 1, rs: 2, cs: 2 }]);
    expect(spanRects(apply(t, removeRow(t, 3)))).toEqual([{ r: 1, c: 1, rs: 2, cs: 2 }]);
    // shrink to one row and column: the region disappears
    let small = merged();
    small = apply(small, removeRow(small, 1));
    small = apply(small, removeColumn(small, 1));
    expect(spanRects(small)).toEqual([]);
    expect(flags(small).join('')).not.toMatch(/[xhv]/);
  });
  it('removing a row that holds a one-row region drops it', () => {
    let t = table(3, 3);
    t = apply(t, mergeCells(t, { r0: 1, c0: 0, r1: 1, c1: 2 }));
    const r = apply(t, removeRow(t, 1));
    expect(spanRects(r)).toEqual([]);
    const c = apply(t, removeColumn(t, 0));
    expect(spanRects(c)).toEqual([{ r: 1, c: 0, rs: 1, cs: 2 }]);
  });
  it('a whole-width region survives removing a column that does not hold its origin', () => {
    let t = table(2, 3);
    t = apply(t, mergeCells(t, { r0: 0, c0: 0, r1: 1, c1: 0 }));
    expect(spanRects(apply(t, removeColumn(t, 2)))).toEqual([{ r: 0, c: 0, rs: 2, cs: 1 }]);
    expect(spanRects(apply(t, removeColumn(t, 0)))).toEqual([]);
  });
});

describe('malformed spans', () => {
  it('clamps, drops overlaps and stale flags', () => {
    const g = makeGrid(2, 2);
    g[0][0] = { ...g[0][0], gridSpan: 9, rowSpan: 9 };
    g[0][1] = { ...g[0][1], gridSpan: 2 };
    g[1][1] = { ...g[1][1], hMerge: true, vMerge: true };
    const n = normalizeSpans(g, 2, 2);
    expect(flags({ cells: n } as TableElement)).toEqual(['2x2.h.', 'v.hv.']);
    const stale = makeGrid(1, 2);
    stale[0][1] = { ...stale[0][1], hMerge: true };
    expect(normalizeSpans(stale, 1, 2)[0][1].hMerge).toBeUndefined();
    const odd = makeGrid(1, 2);
    odd[0][0] = { ...odd[0][0], gridSpan: Number.NaN, rowSpan: 0.4 };
    expect(spanRects({ rows: 1, cols: 2, cells: odd })).toEqual([]);
    const clash = makeGrid(2, 2);
    clash[0][0] = { ...clash[0][0], gridSpan: 2, rowSpan: 2 };
    clash[0][1] = { ...clash[0][1], rowSpan: 2 };
    clash[1][0] = { ...clash[1][0], gridSpan: 2 };
    expect(spanRects({ rows: 2, cols: 2, cells: clash })).toEqual([{ r: 0, c: 0, rs: 2, cs: 2 }]);
    // a hole in the grid data is skipped
    expect(spanRects({ rows: 2, cols: 1, cells: [odd[0]] })).toEqual([]);
  });
  it('createTableElement normalises authored spans and keeps style flags', () => {
    const cells = makeGrid(2, 2);
    cells[0][0] = { ...cells[0][0], gridSpan: 5 };
    const el = createTableElement({ containerId: 's', index: 'a', x: 0, y: 0, w: 10, h: 10, rows: 2, cols: 2, cells, styleFlags: { firstRow: true, bandRow: true } });
    expect(el.cells[0][0].gridSpan).toBe(2);
    expect(el.cells[0][1].hMerge).toBe(true);
    expect(el.styleFlags).toEqual({ firstRow: true, bandRow: true });
    expect(createTableElement({ containerId: 's', index: 'a', x: 0, y: 0, w: 10, h: 10, rows: 1, cols: 1 }).styleFlags).toBeUndefined();
  });
  it('spanAt treats a flagged cell with no region, or a cell outside the grid, as a cell of its own', () => {
    const g = makeGrid(1, 2);
    g[0][1] = { ...g[0][1], hMerge: true };
    expect(spanAt({ rows: 1, cols: 2, cells: g }, 0, 1)).toEqual({ r: 0, c: 1, rs: 1, cs: 1 });
    expect(spanAt({ rows: 1, cols: 2, cells: g }, 4, 4)).toEqual({ r: 4, c: 4, rs: 1, cs: 1 });
  });
  it('applySpanRects leaves plain cells untouched (identity)', () => {
    const g = makeGrid(1, 1);
    expect(applySpanRects(g, [])[0][0]).toBe(g[0][0]);
  });
  it('expandRange clamps and orders its input', () => {
    expect(expandRange(table(2, 2), { r0: 5, c0: 5, r1: -3, c1: -3 })).toEqual({ r0: 0, c0: 0, r1: 1, c1: 1 });
  });
});

describe('setCellBorders / resolveCellEdges', () => {
  it('no borders resolve to the grid line on every edge', () => {
    expect(resolveCellEdges(table(2, 2), 0, 0, GRID)).toEqual({ l: GRID, t: GRID, r: GRID, b: GRID });
  });
  it('an outer border on a block is mirrored onto the neighbours and agrees from both sides', () => {
    let t = table(3, 3);
    t = apply(t, setCellBorders(t, { r0: 1, c0: 1, r1: 1, c1: 1 }, 'outer', RED));
    expect(t.cells[1][0].borders).toEqual({ r: RED });
    expect(t.cells[0][1].borders).toEqual({ b: RED });
    expect(resolveCellEdges(t, 1, 1, GRID)).toEqual({ l: RED, t: RED, r: RED, b: RED });
    expect(resolveCellEdges(t, 1, 0, GRID).r).toBe(RED);
    expect(resolveCellEdges(t, 2, 1, GRID).t).toBe(RED);
    expect(resolveCellEdges(t, 0, 0, GRID)).toEqual({ l: GRID, t: GRID, r: GRID, b: GRID });
  });
  it('inner touches only the lines inside the range; single sides and all work', () => {
    const t = table(2, 2);
    const inner = apply(t, setCellBorders(t, { r0: 0, c0: 0, r1: 1, c1: 1 }, 'inner', null));
    expect(resolveCellEdges(inner, 0, 0, GRID)).toEqual({ l: GRID, t: GRID, r: null, b: null });
    expect(resolveCellEdges(inner, 1, 1, GRID)).toEqual({ l: null, t: null, r: GRID, b: GRID });
    const top = apply(t, setCellBorders(t, { r0: 0, c0: 0, r1: 1, c1: 1 }, 't', RED));
    expect(resolveCellEdges(top, 0, 1, GRID)).toEqual({ l: GRID, t: RED, r: GRID, b: GRID });
    expect(resolveCellEdges(top, 1, 0, GRID).t).toBe(GRID);
    for (const side of ['l', 'r', 'b'] as const) {
      const one = apply(t, setCellBorders(t, { r0: 0, c0: 0, r1: 1, c1: 1 }, side, RED));
      expect(Object.values(resolveCellEdges(one, side === 'b' ? 1 : 0, side === 'r' ? 1 : 0, GRID)).filter((v) => v === RED)).toHaveLength(1);
    }
    const all = apply(t, setCellBorders(t, { r0: 0, c0: 0, r1: 1, c1: 1 }, 'all', RED));
    expect(resolveCellEdges(all, 0, 0, GRID)).toEqual({ l: RED, t: RED, r: RED, b: RED });
  });
  it('undefined resets to the grid line and drops the empty borders object', () => {
    let t = table(1, 2);
    t = apply(t, setCellBorders(t, { r0: 0, c0: 0, r1: 0, c1: 1 }, 'all', RED));
    t = apply(t, setCellBorders(t, { r0: 0, c0: 0, r1: 0, c1: 1 }, 'all', undefined));
    expect(t.cells.flat().every((c) => c.borders === undefined)).toBe(true);
    expect(resolveCellEdges(t, 0, 0, GRID).r).toBe(GRID);
  });
  it('a merged cell draws its right and bottom edges from the cells on them', () => {
    let t = table(3, 3);
    t = apply(t, mergeCells(t, { r0: 0, c0: 0, r1: 1, c1: 1 }));
    t = apply(t, setCellBorders(t, { r0: 0, c0: 0, r1: 0, c1: 0 }, 'outer', RED)); // expands to the whole region
    expect(resolveCellEdges(t, 0, 0, GRID)).toEqual({ l: RED, t: RED, r: RED, b: RED });
    expect(resolveCellEdges(t, 1, 1, GRID)).toEqual(resolveCellEdges(t, 0, 0, GRID)); // any cell of the region resolves alike
    expect(resolveCellEdges(t, 0, 2, GRID).l).toBe(RED);
    expect(resolveCellEdges(t, 2, 0, GRID).t).toBe(RED);
  });
  it('an explicit own side beats the neighbour facing it on the left/top, the neighbour wins right/bottom', () => {
    let t = table(2, 2);
    const cells = t.cells.map((r) => r.slice());
    cells[0][0] = { ...cells[0][0], borders: { r: RED, b: RED } };
    cells[0][1] = { ...cells[0][1], borders: { l: null } };
    cells[1][0] = { ...cells[1][0], borders: { t: null } };
    t = { ...t, cells };
    expect(resolveCellEdges(t, 0, 0, GRID).r).toBeNull();
    expect(resolveCellEdges(t, 0, 1, GRID).l).toBeNull();
    expect(resolveCellEdges(t, 0, 0, GRID).b).toBeNull();
    expect(resolveCellEdges(t, 1, 0, GRID).t).toBeNull();
  });
});
