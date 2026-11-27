// Pure table-grid operations. A table stores column/row sizes as fractions of
// its box (they always sum to ~1), and a row-major grid of cells. Every helper
// returns the changed subset as a Partial<TableElement>; the deck applies it via
// updateElement so undo/collab treat a structural edit as one change.

import type { RichTextDoc } from '../text/model';
import { emptyRichText, isRichTextEmpty } from '../text/model';
import type { CellBorders, CellSide, Stroke, TableCell, TableElement } from './types';

const MIN_FRACTION = 0.04;

export function evenFractions(n: number): number[] {
  return Array.from({ length: Math.max(1, n) }, () => 1 / Math.max(1, n));
}

function normalize(fracs: number[]): number[] {
  const sum = fracs.reduce((a, b) => a + b, 0) || 1;
  return fracs.map((f) => f / sum);
}

export function emptyCell(): TableCell {
  return { richText: emptyRichText() as RichTextDoc };
}

/** A fresh rows×cols grid of empty cells. */
export function makeGrid(rows: number, cols: number): TableCell[][] {
  return Array.from({ length: rows }, () => Array.from({ length: cols }, emptyCell));
}

/** Insert a row at `at` (0..rows); existing rows shrink proportionally. */
export function insertRow(t: TableElement, at: number): Partial<TableElement> {
  const rows = t.rows + 1;
  const idx = clamp(at, 0, t.rows);
  const scaled = t.rowFractions.map((f) => (f * t.rows) / rows);
  scaled.splice(idx, 0, 1 / rows);
  const rects = spanRects(t).map((s) => (s.r >= idx ? { ...s, r: s.r + 1 } : s.r < idx && idx < s.r + s.rs ? { ...s, rs: s.rs + 1 } : s));
  const cells = t.cells.map((r) => r.slice());
  cells.splice(idx, 0, Array.from({ length: t.cols }, emptyCell));
  const patch: Partial<TableElement> = { rows, rowFractions: normalize(scaled), cells: applySpanRects(cells, rects) };
  if (t.rowHeights) { const rh = t.rowHeights.slice(); rh.splice(idx, 0, 0); patch.rowHeights = rh; } // new row auto-sizes
  return patch;
}

/** Insert a column at `at` (0..cols); existing columns shrink proportionally. */
export function insertColumn(t: TableElement, at: number): Partial<TableElement> {
  const cols = t.cols + 1;
  const scaled = t.colFractions.map((f) => (f * t.cols) / cols);
  const idx = clamp(at, 0, t.cols);
  scaled.splice(idx, 0, 1 / cols);
  const rects = spanRects(t).map((s) => (s.c >= idx ? { ...s, c: s.c + 1 } : s.c < idx && idx < s.c + s.cs ? { ...s, cs: s.cs + 1 } : s));
  const cells = t.cells.map((r) => {
    const nr = r.slice();
    nr.splice(idx, 0, emptyCell());
    return nr;
  });
  return { cols, colFractions: normalize(scaled), cells: applySpanRects(cells, rects) };
}

/** Remove row `at` (no-op if only one row remains). */
export function removeRow(t: TableElement, at: number): Partial<TableElement> {
  if (t.rows <= 1) return {};
  const idx = clamp(at, 0, t.rows - 1);
  const fr = t.rowFractions.slice();
  fr.splice(idx, 1);
  // A merged cell that loses its first row keeps its content on the cell that becomes its origin.
  const carried = spanRects(t).filter((s) => s.r === idx && s.rs > 1);
  const rects: SpanRect[] = [];
  for (const s of spanRects(t)) {
    if (s.r > idx) rects.push({ ...s, r: s.r - 1 });
    else if (s.r + s.rs - 1 < idx) rects.push(s);
    else if (s.rs > 1) rects.push({ ...s, rs: s.rs - 1 });
  }
  const cells = t.cells.slice();
  cells.splice(idx, 1);
  for (const s of carried) cells[idx] = cells[idx].map((cell, c) => (c === s.c ? carryOrigin(t.cells[idx][c], cell) : cell));
  const patch: Partial<TableElement> = { rows: t.rows - 1, rowFractions: normalize(fr), cells: applySpanRects(cells, rects) };
  if (t.rowHeights) { const rh = t.rowHeights.slice(); rh.splice(idx, 1); patch.rowHeights = rh; }
  return patch;
}

/** Remove column `at` (no-op if only one column remains). */
export function removeColumn(t: TableElement, at: number): Partial<TableElement> {
  if (t.cols <= 1) return {};
  const idx = clamp(at, 0, t.cols - 1);
  const fr = t.colFractions.slice();
  fr.splice(idx, 1);
  const carried = spanRects(t).filter((s) => s.c === idx && s.cs > 1);
  const rects: SpanRect[] = [];
  for (const s of spanRects(t)) {
    if (s.c > idx) rects.push({ ...s, c: s.c - 1 });
    else if (s.c + s.cs - 1 < idx) rects.push(s);
    else if (s.cs > 1) rects.push({ ...s, cs: s.cs - 1 });
  }
  const cells = t.cells.map((r, ri) => {
    const nr = r.filter((_, c) => c !== idx);
    for (const s of carried) if (s.r === ri) nr[idx] = carryOrigin(r[idx], nr[idx]);
    return nr;
  });
  return { cols: t.cols - 1, colFractions: normalize(fr), cells: applySpanRects(cells, rects) };
}

/** Resize the boundary between column `index` and `index+1` by `delta` (a
 *  fraction of table width), borrowing from the right neighbor. */
export function resizeColumn(t: TableElement, index: number, delta: number): Partial<TableElement> {
  if (index < 0 || index >= t.cols - 1) return {};
  const fr = t.colFractions.slice();
  const d = clampDelta(fr[index], fr[index + 1], delta);
  fr[index] += d;
  fr[index + 1] -= d;
  return { colFractions: fr };
}

/** Set row `index`'s minimum height in px (content can still push it taller).
 *  0 restores auto (content-driven) sizing. */
export function setRowHeight(t: TableElement, index: number, height: number): Partial<TableElement> {
  if (index < 0 || index >= t.rows) return {};
  const rh = (t.rowHeights ?? Array.from({ length: t.rows }, () => 0)).slice();
  rh.length = t.rows; // guard against a stale-length array
  for (let i = 0; i < t.rows; i++) rh[i] = rh[i] ?? 0;
  rh[index] = Math.max(0, Math.round(height));
  return { rowHeights: rh };
}

function clamp(n: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, n));
}

/** Keep both sides ≥ MIN_FRACTION when shifting a boundary. */
function clampDelta(a: number, b: number, delta: number): number {
  return clamp(delta, MIN_FRACTION - a, b - MIN_FRACTION);
}

// ── Merged cells ─────────────────────────────────────────────────────────────

/** A merged region: its origin cell (top-left) and the rows/columns it covers. */
export interface SpanRect {
  r: number;
  c: number;
  rs: number;
  cs: number;
}

/** An inclusive rectangle of grid cells. */
export interface CellRange {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

/**
 * The merged regions of a grid, read from the origin cells' spans. Spans are clamped to the grid, and a region that
 * overlaps an earlier one is dropped, so malformed authored or imported data still yields disjoint rectangles.
 */
export function spanRects(t: Pick<TableElement, 'rows' | 'cols' | 'cells'>): SpanRect[] {
  const taken = new Set<number>();
  const out: SpanRect[] = [];
  for (let r = 0; r < t.rows; r++) {
    for (let c = 0; c < t.cols; c++) {
      const cell = t.cells[r]?.[c];
      if (!cell || cell.hMerge || cell.vMerge) continue;
      const rs = clamp(Math.floor(cell.rowSpan ?? 1) || 1, 1, t.rows - r);
      const cs = clamp(Math.floor(cell.gridSpan ?? 1) || 1, 1, t.cols - c);
      if (rs === 1 && cs === 1) continue;
      let free = true;
      for (let i = r; i < r + rs && free; i++) for (let j = c; j < c + cs; j++) if (taken.has(i * t.cols + j)) { free = false; break; }
      if (!free) continue;
      for (let i = r; i < r + rs; i++) for (let j = c; j < c + cs; j++) taken.add(i * t.cols + j);
      out.push({ r, c, rs, cs });
    }
  }
  return out;
}

/** Rewrite every cell's span flags from `rects` (the origin carries the spans; covered cells carry hMerge/vMerge). */
export function applySpanRects(cells: TableCell[][], rects: readonly SpanRect[]): TableCell[][] {
  const grid = cells.map((row) =>
    row.map((cell) => {
      if (cell.gridSpan === undefined && cell.rowSpan === undefined && !cell.hMerge && !cell.vMerge) return cell;
      const { gridSpan: _g, rowSpan: _r, hMerge: _h, vMerge: _v, ...rest } = cell;
      return rest as TableCell;
    }),
  );
  for (const { r, c, rs, cs } of rects) {
    for (let i = r; i < r + rs; i++) {
      for (let j = c; j < c + cs; j++) {
        const cell = grid[i][j];
        if (i === r && j === c) grid[i][j] = { ...cell, ...(cs > 1 ? { gridSpan: cs } : {}), ...(rs > 1 ? { rowSpan: rs } : {}) };
        else grid[i][j] = { ...cell, ...(j > c ? { hMerge: true } : {}), ...(i > r ? { vMerge: true } : {}) };
      }
    }
  }
  return grid;
}

/** The cell that becomes a merged region's origin after its old origin row/column was removed keeps the old content. */
function carryOrigin(old: TableCell, next: TableCell): TableCell {
  return { ...next, richText: old.richText, fill: old.fill, bodyStyle: old.bodyStyle, borders: old.borders ?? next.borders };
}

/** The region covering (row, col): its merged rectangle, or the cell alone. */
export function spanAt(t: Pick<TableElement, 'rows' | 'cols' | 'cells'>, row: number, col: number): SpanRect {
  const cell = t.cells[row]?.[col];
  // A cell with no span data is a region of its own (the common case; skips scanning every region).
  if (cell && cell.gridSpan === undefined && cell.rowSpan === undefined && !cell.hMerge && !cell.vMerge) return { r: row, c: col, rs: 1, cs: 1 };
  for (const s of spanRects(t)) if (row >= s.r && row < s.r + s.rs && col >= s.c && col < s.c + s.cs) return s;
  return { r: row, c: col, rs: 1, cs: 1 };
}

/** Grow a range until it cuts through no merged region (a selection that touches part of one takes all of it). */
export function expandRange(t: Pick<TableElement, 'rows' | 'cols' | 'cells'>, range: CellRange): CellRange {
  const rects = spanRects(t);
  let { r0, c0, r1, c1 } = {
    r0: clamp(Math.min(range.r0, range.r1), 0, t.rows - 1),
    c0: clamp(Math.min(range.c0, range.c1), 0, t.cols - 1),
    r1: clamp(Math.max(range.r0, range.r1), 0, t.rows - 1),
    c1: clamp(Math.max(range.c0, range.c1), 0, t.cols - 1),
  };
  for (let grew = true; grew; ) {
    grew = false;
    for (const s of rects) {
      const hits = s.r <= r1 && s.r + s.rs - 1 >= r0 && s.c <= c1 && s.c + s.cs - 1 >= c0;
      if (!hits) continue;
      const n = { r0: Math.min(r0, s.r), c0: Math.min(c0, s.c), r1: Math.max(r1, s.r + s.rs - 1), c1: Math.max(c1, s.c + s.cs - 1) };
      if (n.r0 !== r0 || n.c0 !== c0 || n.r1 !== r1 || n.c1 !== c1) { ({ r0, c0, r1, c1 } = n); grew = true; }
    }
  }
  return { r0, c0, r1, c1 };
}

/**
 * Merge the cells of `range` (expanded to whole merged regions) into one. The top-left cell is the origin and takes
 * the text of the others (non-empty bodies appended as paragraphs); the covered cells are emptied but keep their own
 * fill and borders, which come back when the cell is split. A range of one cell is a no-op.
 */
export function mergeCells(t: TableElement, range: CellRange): Partial<TableElement> {
  const { r0, c0, r1, c1 } = expandRange(t, range);
  if (r0 === r1 && c0 === c1) return {};
  const paragraphs: RichTextDoc['content'] = [];
  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      const cell = t.cells[r][c];
      if (!cell.hMerge && !cell.vMerge && !isRichTextEmpty(cell.richText)) paragraphs.push(...cell.richText.content);
    }
  }
  const grid = t.cells.map((row) => row.slice());
  for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) grid[r][c] = { ...grid[r][c], richText: emptyRichText() };
  if (paragraphs.length) grid[r0][c0] = { ...grid[r0][c0], richText: { type: 'doc', content: paragraphs } };
  const rects = spanRects(t).filter((s) => !(s.r >= r0 && s.r <= r1 && s.c >= c0 && s.c <= c1));
  rects.push({ r: r0, c: c0, rs: r1 - r0 + 1, cs: c1 - c0 + 1 });
  return { cells: applySpanRects(grid, rects) };
}

/** Split the merged region containing (row, col) back into single cells (a no-op for an unmerged cell). */
export function splitCell(t: TableElement, row: number, col: number): Partial<TableElement> {
  const s = spanAt(t, row, col);
  if (s.rs === 1 && s.cs === 1) return {};
  return { cells: applySpanRects(t.cells, spanRects(t).filter((x) => x !== s && !(x.r === s.r && x.c === s.c))) };
}

// ── Cell borders ─────────────────────────────────────────────────────────────

/** Which edges of a range `setCellBorders` writes: all, the outline, the inside lines, or one outer side. */
export type BorderScope = 'all' | 'outer' | 'inner' | CellSide;

const OPPOSITE: Record<CellSide, CellSide> = { l: 'r', r: 'l', t: 'b', b: 't' };

function withSide(cell: TableCell, side: CellSide, value: Stroke | null | undefined): TableCell {
  const borders: CellBorders = { ...cell.borders };
  if (value === undefined) delete borders[side];
  else borders[side] = value;
  const { borders: _old, ...rest } = cell;
  return Object.keys(borders).length ? { ...rest, borders } : rest;
}

/**
 * Set (a stroke), clear (`null`: no line) or reset to the grid line (`undefined`) the edges of `range` named by
 * `scope`. A shared edge is written on both neighbours so the two cells agree. The range is expanded to whole merged
 * regions first.
 */
export function setCellBorders(t: TableElement, range: CellRange, scope: BorderScope, value: Stroke | null | undefined): Partial<TableElement> {
  const { r0, c0, r1, c1 } = expandRange(t, range);
  const grid = t.cells.map((row) => row.slice());
  const put = (r: number, c: number, side: CellSide) => {
    grid[r][c] = withSide(grid[r][c], side, value);
    const nr = r + (side === 't' ? -1 : side === 'b' ? 1 : 0);
    const nc = c + (side === 'l' ? -1 : side === 'r' ? 1 : 0);
    if (grid[nr]?.[nc]) grid[nr][nc] = withSide(grid[nr][nc], OPPOSITE[side], value);
  };
  const want = (side: CellSide, outer: boolean) => scope === 'all' || (scope === 'outer' ? outer : scope === 'inner' ? !outer : outer && scope === side);
  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      if (want('l', c === c0)) put(r, c, 'l');
      if (c === c1 && want('r', true)) put(r, c, 'r');
      if (want('t', r === r0)) put(r, c, 't');
      if (r === r1 && want('b', true)) put(r, c, 'b');
    }
  }
  return { cells: grid };
}

/**
 * The four edges the cell at (row, col) is drawn with: its own side if set, else the neighbour's facing side, else
 * `fallback` (the table's grid line). `null` is no line. The same rule runs for both neighbours of a shared edge
 * (the right/lower cell's value wins), so they never disagree. A merged cell reads its right and bottom edges from the
 * grid cells that sit on them.
 */
export function resolveCellEdges(t: TableElement, row: number, col: number, fallback: Stroke): Record<CellSide, Stroke | null> {
  const s = spanAt(t, row, col);
  const at = (r: number, c: number, side: CellSide): Stroke | null | undefined => t.cells[r]?.[c]?.borders?.[side];
  // `null` (no line) is a value here, so the first side that is not `undefined` wins (not `??`).
  const pick = (first: Stroke | null | undefined, second: Stroke | null | undefined): Stroke | null =>
    first !== undefined ? first : second !== undefined ? second : fallback;
  return {
    l: pick(at(s.r, s.c, 'l'), at(s.r, s.c - 1, 'r')),
    t: pick(at(s.r, s.c, 't'), at(s.r - 1, s.c, 'b')),
    r: pick(at(s.r, s.c + s.cs, 'l'), at(s.r, s.c + s.cs - 1, 'r')),
    b: pick(at(s.r + s.rs, s.c, 't'), at(s.r + s.rs - 1, s.c, 'b')),
  };
}

/** Normalise a grid's span flags to disjoint in-bounds regions (malformed authored or imported data). */
export function normalizeSpans(cells: TableCell[][], rows: number, cols: number): TableCell[][] {
  return applySpanRects(cells, spanRects({ rows, cols, cells }));
}
