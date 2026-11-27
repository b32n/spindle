// Sheet model with sparse cell storage keyed by stable row/column IDs.
//
// Storage invariants:
//   - `cells` is keyed by `${rowId}:${colId}` (stable IDs from utils/id).
//   - `rowOrder`/`colOrder` are sparse: only touched indices have an entry.
//   - Untouched indices are "virtual"; getCell returns undefined for them.
//   - insertRows/deleteRows/insertCols/deleteCols only mutate the order maps.
//     Cells stay attached to their stable IDs across inserts and sorts; the
//     numeric index a cell appears at changes, but its identity does not.

import type { Sheet, SheetConfig, Cell, Range, SortOrder, MergedRegion } from './types';
import { CommentStore } from './comments';
import { getStableCellKey, parseStableCellKey } from './utils/cell-key';
import { generateId } from './utils/id';
import { getRangeCells } from './utils/range';
import { coveredBeyondAnchor, lastWins, normalized } from './merge-regions';

/** Whether two normalized numeric ranges (inclusive bounds) overlap. */
function rangesIntersect(a: Range, b: Range): boolean {
  return (
    a.startRow <= b.endRow &&
    a.endRow >= b.startRow &&
    a.startCol <= b.endCol &&
    a.endCol >= b.startCol
  );
}

/**
 * Most merges `mergeCellsBulk` decides pair by pair when some overlap (a sound file has none that do; the cost is a comparison for each
 * pair, so this many take a fraction of a second).
 */
export const MAX_OVERLAPPING_MERGES = 20_000;

export class SheetImpl implements Sheet {
  id: string;
  name: string;
  cells: Map<string, Cell> = new Map(); // key: "rowId:colId"
  config: SheetConfig;
  rowCount: number;
  colCount: number;
  /** Comment threads anchored to cells in this sheet. */
  comments: CommentStore = new CommentStore();

  // Sparse index → stable ID maps. Only filled for indices that have been
  // explicitly touched (cell set, row/col operation, etc.).
  protected rowOrder: Map<number, string> = new Map();
  protected colOrder: Map<number, string> = new Map();
  // Reverse maps for O(1) stable→index lookup. Kept in sync with the
  // forward maps.
  protected rowIdToIndex: Map<string, number> = new Map();
  protected colIdToIndex: Map<string, number> = new Map();

  // Fired after any structural mutation (insert/delete row/col, height/
  // width, hide/show, freeze, filter, sort). The collab binding wires
  // this on attachCollab to mirror the sheet's metadata + order maps
  // into the Y.Doc. No-op when collab is detached.
  protected structureChangeListener: ((affectsFormulas: boolean) => void) | undefined;

  /** @internal Used by the workbook (recalc + collab mirror). */
  __setStructureChangeListener(listener: ((affectsFormulas: boolean) => void) | undefined): void {
    this.structureChangeListener = listener;
  }

  /**
   * Notify a structural change. `affectsFormulas` is true only for edits that
   * shift cell references (insert/delete row/col) — cosmetic changes like column
   * width, row height, hide, and freeze pass false so they don't trigger a
   * full-workbook recalc (which would fire on every mousemove of a resize drag).
   */
  protected notifyStructureChange(affectsFormulas = false): void {
    this.structureChangeListener?.(affectsFormulas);
  }

  constructor(id: string, name: string, config: Partial<SheetConfig> = {}) {
    this.id = id;
    this.name = name;
    this.config = {
      defaultRowHeight: 20,
      defaultColWidth: 100,
      showGridLines: true,
      ...config,
    };
    this.rowCount = 1000;
    this.colCount = 100;
  }

  // ============================================
  // Stable ID helpers
  // ============================================

  getRowId(row: number): string | undefined {
    return this.rowOrder.get(row);
  }

  getColId(col: number): string | undefined {
    return this.colOrder.get(col);
  }

  ensureRowId(row: number): string {
    let id = this.rowOrder.get(row);
    if (!id) {
      id = generateId();
      this.rowOrder.set(row, id);
      this.rowIdToIndex.set(id, row);
    }
    return id;
  }

  ensureColId(col: number): string {
    let id = this.colOrder.get(col);
    if (!id) {
      id = generateId();
      this.colOrder.set(col, id);
      this.colIdToIndex.set(id, col);
    }
    return id;
  }

  getRowIndex(rowId: string): number | undefined {
    return this.rowIdToIndex.get(rowId);
  }

  getColIndex(colId: string): number | undefined {
    return this.colIdToIndex.get(colId);
  }

  /** Translate a stable cell key into (row, col) indices, if both halves are known. */
  stableKeyToIndices(key: string): { row: number; col: number } | undefined {
    const { rowId, colId } = parseStableCellKey(key);
    const row = this.rowIdToIndex.get(rowId);
    const col = this.colIdToIndex.get(colId);
    if (row === undefined || col === undefined) return undefined;
    return { row, col };
  }

  // ============================================
  // Cell access (public API stays index-based)
  // ============================================

  getCell(row: number, col: number): Cell | undefined {
    const rowId = this.rowOrder.get(row);
    const colId = this.colOrder.get(col);
    if (!rowId || !colId) return undefined;
    return this.cells.get(getStableCellKey(rowId, colId));
  }

  setCell(row: number, col: number, cell: Partial<Cell>): void {
    const rowId = this.ensureRowId(row);
    const colId = this.ensureColId(col);
    const key = getStableCellKey(rowId, colId);
    const existing = this.cells.get(key);

    if (existing) {
      this.cells.set(key, { ...existing, ...cell });
    } else {
      this.cells.set(key, { value: cell.value ?? null, ...cell });
    }
  }

  deleteCell(row: number, col: number): void {
    const rowId = this.rowOrder.get(row);
    const colId = this.colOrder.get(col);
    if (!rowId || !colId) return;
    this.cells.delete(getStableCellKey(rowId, colId));
  }

  getCellValue(row: number, col: number): unknown {
    const cell = this.getCell(row, col);
    return cell?.value ?? null;
  }

  setCellValue(row: number, col: number, value: unknown): void {
    this.setCell(row, col, { value: value as string | number | boolean | null });
  }

  getRange(range: Range): Map<string, Cell> {
    const result = new Map<string, Cell>();
    const cells = getRangeCells(range);

    for (const { row, col } of cells) {
      const cell = this.getCell(row, col);
      if (cell) {
        const rowId = this.rowOrder.get(row);
        const colId = this.colOrder.get(col);
        if (rowId && colId) {
          result.set(getStableCellKey(rowId, colId), cell);
        }
      }
    }

    return result;
  }

  setRange(range: Range, cells: Map<string, Cell> | Cell[][]): void {
    if (cells instanceof Map) {
      // Keys here are stable; translate back to indices via this sheet's reverse maps.
      for (const [key, cell] of cells) {
        const indices = this.stableKeyToIndices(key);
        if (indices) {
          this.setCell(indices.row, indices.col, cell);
        }
      }
    } else {
      const rangeCells = getRangeCells(range);
      const flat = cells.flat();
      for (let i = 0; i < rangeCells.length && i < flat.length; i++) {
        const { row, col } = rangeCells[i];
        const cell = flat[i];
        if (cell) {
          this.setCell(row, col, cell);
        }
      }
    }
  }

  clearRange(range: Range): void {
    const cells = getRangeCells(range);
    for (const { row, col } of cells) {
      this.deleteCell(row, col);
    }
  }

  /**
   * Iterate non-empty cells as (row, col, cell) triples, using current
   * row/column indices. Cells whose stable IDs are no longer in the order
   * maps (e.g. after deleteRows) are skipped.
   */
  *entries(): IterableIterator<[number, number, Cell]> {
    for (const [key, cell] of this.cells) {
      const indices = this.stableKeyToIndices(key);
      if (!indices) continue;
      yield [indices.row, indices.col, cell];
    }
  }

  getRowHeight(row: number): number {
    const rowId = this.rowOrder.get(row);
    const height = rowId ? this.config.rowHeights?.get(rowId) : undefined;
    return height ?? this.config.defaultRowHeight ?? 20;
  }

  setRowHeight(row: number, height: number): void {
    if (!this.config.rowHeights) {
      this.config.rowHeights = new Map();
    }
    this.config.rowHeights.set(this.ensureRowId(row), height);
    this.notifyStructureChange();
  }

  getColWidth(col: number): number {
    const colId = this.colOrder.get(col);
    const width = colId ? this.config.colWidths?.get(colId) : undefined;
    return width ?? this.config.defaultColWidth ?? 100;
  }

  setColWidth(col: number, width: number): void {
    if (!this.config.colWidths) {
      this.config.colWidths = new Map();
    }
    this.config.colWidths.set(this.ensureColId(col), width);
    this.notifyStructureChange();
  }

  // ============================================
  // Insert / delete (mutate order maps only)
  // ============================================

  insertRows(startRow: number, count: number): void {
    this.shiftOrder(this.rowOrder, this.rowIdToIndex, startRow, count);
    this.rowCount += count;
    this.notifyStructureChange(true);
  }

  deleteRows(startRow: number, count: number): void {
    this.removeOrderRange(this.rowOrder, this.rowIdToIndex, startRow, count, /*isRow*/ true);
    this.shiftOrder(this.rowOrder, this.rowIdToIndex, startRow + count, -count, startRow);
    this.rowCount = Math.max(0, this.rowCount - count);
    this.notifyStructureChange(true);
  }

  insertCols(startCol: number, count: number): void {
    this.shiftOrder(this.colOrder, this.colIdToIndex, startCol, count);
    this.colCount += count;
    this.notifyStructureChange(true);
  }

  deleteCols(startCol: number, count: number): void {
    this.removeOrderRange(this.colOrder, this.colIdToIndex, startCol, count, /*isRow*/ false);
    this.shiftOrder(this.colOrder, this.colIdToIndex, startCol + count, -count, startCol);
    this.colCount = Math.max(0, this.colCount - count);
    this.notifyStructureChange(true);
  }

  /** Shift every order entry with index >= threshold by `delta`. */
  private shiftOrder(
    order: Map<number, string>,
    reverse: Map<string, number>,
    threshold: number,
    delta: number,
    /** When shifting down (delta<0), only entries strictly >= floor move. */
    floor: number = threshold,
  ): void {
    if (delta === 0) return;
    const toMove: Array<[number, string]> = [];
    for (const [idx, id] of order) {
      if (idx >= threshold) toMove.push([idx, id]);
    }
    // Sort to avoid clobbering: shifting up → descending; shifting down → ascending.
    toMove.sort(([a], [b]) => (delta > 0 ? b - a : a - b));
    for (const [idx, id] of toMove) {
      order.delete(idx);
      reverse.delete(id);
    }
    for (const [idx, id] of toMove) {
      const newIdx = idx + delta;
      if (newIdx < floor) continue; // shouldn't happen, but guard
      order.set(newIdx, id);
      reverse.set(id, newIdx);
    }
  }

  /**
   * Drop order entries in [start, start+count) and any cells attached to
   * those IDs. Also drop matching config entries (height/width/hidden).
   */
  private removeOrderRange(
    order: Map<number, string>,
    reverse: Map<string, number>,
    start: number,
    count: number,
    isRow: boolean,
  ): void {
    const removedIds = new Set<string>();
    for (const [idx, id] of order) {
      if (idx >= start && idx < start + count) {
        removedIds.add(id);
      }
    }
    if (removedIds.size > 0) {
      // Delete cells whose row half (or col half) is in removedIds.
      for (const key of [...this.cells.keys()]) {
        const { rowId, colId } = parseStableCellKey(key);
        if (isRow ? removedIds.has(rowId) : removedIds.has(colId)) {
          this.cells.delete(key);
        }
      }

      for (const id of removedIds) {
        const idx = reverse.get(id);
        if (idx !== undefined) order.delete(idx);
        reverse.delete(id);
      }
    }

    this.dropConfigForIds(removedIds, isRow);
  }

  /** Drop stable-ID-keyed config entries for the given row or column IDs. */
  private dropConfigForIds(ids: Set<string>, isRow: boolean): void {
    if (ids.size === 0) return;
    if (isRow) {
      if (this.config.rowHeights) for (const id of ids) this.config.rowHeights.delete(id);
      if (this.config.hiddenRows) for (const id of ids) this.config.hiddenRows.delete(id);
    } else {
      if (this.config.colWidths) for (const id of ids) this.config.colWidths.delete(id);
      if (this.config.hiddenCols) for (const id of ids) this.config.hiddenCols.delete(id);
      if (this.config.filters) for (const id of ids) this.config.filters.delete(id);
    }
    // Drop merged regions whose removed corner row/col no longer exists.
    if (this.config.mergedRegions) {
      this.config.mergedRegions = this.config.mergedRegions.filter((region) =>
        isRow
          ? !ids.has(region.startRowId) && !ids.has(region.endRowId)
          : !ids.has(region.startColId) && !ids.has(region.endColId)
      );
    }
  }

  isRowHidden(row: number): boolean {
    const rowId = this.rowOrder.get(row);
    return rowId ? (this.config.hiddenRows?.has(rowId) ?? false) : false;
  }

  hideRow(row: number): void {
    if (!this.config.hiddenRows) {
      this.config.hiddenRows = new Set();
    }
    this.config.hiddenRows.add(this.ensureRowId(row));
    this.notifyStructureChange();
  }

  showRow(row: number): void {
    const rowId = this.rowOrder.get(row);
    if (rowId) {
      this.config.hiddenRows?.delete(rowId);
      this.notifyStructureChange();
    }
  }

  isColHidden(col: number): boolean {
    const colId = this.colOrder.get(col);
    return colId ? (this.config.hiddenCols?.has(colId) ?? false) : false;
  }

  hideCol(col: number): void {
    if (!this.config.hiddenCols) {
      this.config.hiddenCols = new Set();
    }
    this.config.hiddenCols.add(this.ensureColId(col));
    this.notifyStructureChange();
  }

  showCol(col: number): void {
    const colId = this.colOrder.get(col);
    if (colId) {
      this.config.hiddenCols?.delete(colId);
      this.notifyStructureChange();
    }
  }

  // Find hidden columns adjacent to a given column (before and after)
  getHiddenColsAdjacent(col: number): { before: number[]; after: number[] } {
    if (!this.config.hiddenCols || this.config.hiddenCols.size === 0) {
      return { before: [], after: [] };
    }

    const before: number[] = [];
    const after: number[] = [];

    for (let c = col - 1; c >= 0; c--) {
      if (this.isColHidden(c)) before.push(c);
      else break;
    }

    for (let c = col + 1; c < this.colCount; c++) {
      if (this.isColHidden(c)) after.push(c);
      else break;
    }

    return { before, after };
  }

  getHiddenRowsAdjacent(row: number): { above: number[]; below: number[] } {
    if (!this.config.hiddenRows || this.config.hiddenRows.size === 0) {
      return { above: [], below: [] };
    }

    const above: number[] = [];
    const below: number[] = [];

    for (let r = row - 1; r >= 0; r--) {
      if (this.isRowHidden(r)) above.push(r);
      else break;
    }

    for (let r = row + 1; r < this.rowCount; r++) {
      if (this.isRowHidden(r)) below.push(r);
      else break;
    }

    return { above, below };
  }

  showColsInRange(startCol: number, endCol: number): void {
    if (!this.config.hiddenCols) return;
    const minCol = Math.min(startCol, endCol);
    const maxCol = Math.max(startCol, endCol);
    for (let c = minCol; c <= maxCol; c++) {
      const colId = this.colOrder.get(c);
      if (colId) this.config.hiddenCols.delete(colId);
    }
    this.notifyStructureChange();
  }

  showRowsInRange(startRow: number, endRow: number): void {
    if (!this.config.hiddenRows) return;
    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);
    for (let r = minRow; r <= maxRow; r++) {
      const rowId = this.rowOrder.get(r);
      if (rowId) this.config.hiddenRows.delete(rowId);
    }
    this.notifyStructureChange();
  }

  // ============================================
  // Freeze Panes Methods
  // ============================================

  getFrozenRows(): number {
    return this.config.frozenRows ?? 0;
  }

  setFrozenRows(count: number): void {
    if (count < 0) {
      throw new Error('Frozen row count cannot be negative');
    }
    if (count >= this.rowCount) {
      throw new Error('Cannot freeze all rows');
    }
    this.config.frozenRows = count > 0 ? count : undefined;
    this.notifyStructureChange();
  }

  getFrozenCols(): number {
    return this.config.frozenCols ?? 0;
  }

  setFrozenCols(count: number): void {
    if (count < 0) {
      throw new Error('Frozen column count cannot be negative');
    }
    if (count >= this.colCount) {
      throw new Error('Cannot freeze all columns');
    }
    this.config.frozenCols = count > 0 ? count : undefined;
    this.notifyStructureChange();
  }

  setFreeze(rows: number, cols: number): void {
    this.setFrozenRows(rows);
    this.setFrozenCols(cols);
  }

  clearFreeze(): void {
    this.config.frozenRows = undefined;
    this.config.frozenCols = undefined;
    this.notifyStructureChange();
  }

  hasFrozenPanes(): boolean {
    return (this.config.frozenRows ?? 0) > 0 || (this.config.frozenCols ?? 0) > 0;
  }

  // ============================================
  // Sorting Methods
  // ============================================

  setSortOrder(sortOrder: SortOrder[]): void {
    this.config.sortOrder = [...sortOrder];
    this.notifyStructureChange();
  }

  getSortOrder(): SortOrder[] {
    return this.config.sortOrder ?? [];
  }

  clearSort(): void {
    this.config.sortOrder = undefined;
    this.notifyStructureChange();
  }

  hasSort(): boolean {
    return (this.config.sortOrder?.length ?? 0) > 0;
  }

  // ============================================
  // Filtering Methods
  // ============================================

  setFilter(column: number, filter: import('./types').ColumnFilter): void {
    if (!this.config.filters) {
      this.config.filters = new Map();
    }
    this.config.filters.set(this.ensureColId(column), filter);
    this.notifyStructureChange();
  }

  clearFilter(column: number): void {
    const colId = this.colOrder.get(column);
    if (colId) {
      this.config.filters?.delete(colId);
      this.notifyStructureChange();
    }
  }

  /**
   * Return filters as a Map keyed by current numeric column index, translating
   * stable colIds back through colOrder. The numeric `column` field inside each
   * ColumnFilter is overwritten to match its current index so downstream
   * consumers don't have to re-translate.
   */
  getFilters(): Map<number, import('./types').ColumnFilter> {
    if (!this.config.filters) return new Map();
    const out = new Map<number, import('./types').ColumnFilter>();
    for (const [colId, filter] of this.config.filters) {
      const col = this.colIdToIndex.get(colId);
      if (col === undefined) continue;
      out.set(col, { ...filter, column: col });
    }
    return out;
  }

  hasFilter(column: number): boolean {
    const colId = this.colOrder.get(column);
    return colId ? (this.config.filters?.has(colId) ?? false) : false;
  }

  clearAllFilters(): void {
    this.config.filters = undefined;
    this.notifyStructureChange();
  }

  // ============================================
  // Merged cells
  // ============================================

  /** Resolve a stored region to a normalized numeric range, or null if orphaned. */
  private regionToRange(region: MergedRegion): Range | null {
    const r0 = this.rowIdToIndex.get(region.startRowId);
    const c0 = this.colIdToIndex.get(region.startColId);
    const r1 = this.rowIdToIndex.get(region.endRowId);
    const c1 = this.colIdToIndex.get(region.endColId);
    if (r0 === undefined || c0 === undefined || r1 === undefined || c1 === undefined) {
      return null;
    }
    return {
      startRow: Math.min(r0, r1),
      endRow: Math.max(r0, r1),
      startCol: Math.min(c0, c1),
      endCol: Math.max(c0, c1),
    };
  }

  getMergedRegions(): Range[] {
    const regions = this.config.mergedRegions;
    if (!regions) return [];
    const out: Range[] = [];
    for (const region of regions) {
      const span = this.regionToRange(region);
      if (span) out.push(span);
    }
    return out;
  }

  getMergeAt(row: number, col: number): Range | undefined {
    for (const span of this.getMergedRegions()) {
      if (
        row >= span.startRow && row <= span.endRow &&
        col >= span.startCol && col <= span.endCol
      ) {
        return span;
      }
    }
    return undefined;
  }

  mergeCells(range: Range): void {
    const target: Range = {
      startRow: Math.min(range.startRow, range.endRow),
      endRow: Math.max(range.startRow, range.endRow),
      startCol: Math.min(range.startCol, range.endCol),
      endCol: Math.max(range.startCol, range.endCol),
    };
    // A single cell cannot be merged.
    if (target.startRow === target.endRow && target.startCol === target.endCol) return;

    // Drop any existing region that overlaps the new one (and any orphans).
    const kept = (this.config.mergedRegions ?? []).filter((region) => {
      const span = this.regionToRange(region);
      return span !== null && !rangesIntersect(span, target);
    });
    // Clear the covered (non-anchor) cells — only the anchor keeps its content.
    for (let r = target.startRow; r <= target.endRow; r++) {
      for (let c = target.startCol; c <= target.endCol; c++) {
        if (r === target.startRow && c === target.startCol) continue;
        this.deleteCell(r, c);
      }
    }
    kept.push({
      startRowId: this.ensureRowId(target.startRow),
      startColId: this.ensureColId(target.startCol),
      endRowId: this.ensureRowId(target.endRow),
      endColId: this.ensureColId(target.endCol),
    });
    this.config.mergedRegions = kept;
    this.notifyStructureChange();
  }

  /**
   * Merge many ranges as if by `mergeCells`, one after the other in this order, with one notification at the end and in near-linear
   * time (`mergeCells` repeated is quadratic in the merges, for each looks at all the others). The result is the same: the same regions
   * in the same order, the same cells cleared. The exception is a file that is not sound, whose merges overlap: those are decided pair by
   * pair, so past `MAX_OVERLAPPING_MERGES` of them (in all, merges already there included) the rest of the ranges are not applied.
   * Returns how many ranges were left out for that.
   */
  mergeCellsBulk(ranges: readonly Range[]): number {
    // a single cell cannot be merged
    const targets = ranges.map(normalized).filter((t) => t.startRow !== t.endRow || t.startCol !== t.endCol);
    if (targets.length === 0) return 0;
    // what is there already (without the orphans, which a merge drops)
    const existing: Array<{ region: MergedRegion; span: Range }> = [];
    for (const region of this.config.mergedRegions ?? []) {
      const span = this.regionToRange(region);
      if (span !== null) existing.push({ region, span });
    }
    const { alive, considered } = lastWins([...existing.map((e) => e.span), ...targets], existing.length, MAX_OVERLAPPING_MERGES);
    const applied = targets.slice(0, considered - existing.length);
    if (applied.length === 0) return targets.length;

    this.clearCovered(applied);
    const kept: MergedRegion[] = [];
    existing.forEach((e, k) => {
      if (alive[k]) kept.push(e.region);
    });
    applied.forEach((t, k) => {
      // the ids of every merge are made, those of a merge a later one drops too, as `mergeCells` makes them
      const region = {
        startRowId: this.ensureRowId(t.startRow),
        startColId: this.ensureColId(t.startCol),
        endRowId: this.ensureRowId(t.endRow),
        endColId: this.ensureColId(t.endCol),
      };
      if (alive[existing.length + k]) kept.push(region);
    });
    this.config.mergedRegions = kept;
    this.notifyStructureChange();
    return targets.length - applied.length;
  }

  /** Clear the cells the merges cover but their first cells (what `mergeCells` does for each), by the cells of the area where that is fewer, or by one sweep over the cells. */
  private clearCovered(merges: readonly Range[]): void {
    if (this.cells.size === 0) return;
    let area = 0;
    for (const t of merges) area += (t.endRow - t.startRow + 1) * (t.endCol - t.startCol + 1);
    if (area <= this.cells.size) {
      for (const t of merges) {
        for (let r = t.startRow; r <= t.endRow; r++) {
          for (let c = t.startCol; c <= t.endCol; c++) {
            if (r !== t.startRow || c !== t.startCol) this.deleteCell(r, c);
          }
        }
      }
      return;
    }
    const keys: string[] = [];
    const at: Array<[number, number]> = [];
    for (const key of this.cells.keys()) {
      const pos = this.stableKeyToIndices(key);
      if (pos === undefined) continue;
      keys.push(key);
      at.push([pos.row, pos.col]);
    }
    for (const i of coveredBeyondAnchor(at, merges)) this.cells.delete(keys[i]);
  }

  unmergeCells(range: Range): void {
    const regions = this.config.mergedRegions;
    if (!regions || regions.length === 0) return;
    const target: Range = {
      startRow: Math.min(range.startRow, range.endRow),
      endRow: Math.max(range.startRow, range.endRow),
      startCol: Math.min(range.startCol, range.endCol),
      endCol: Math.max(range.startCol, range.endCol),
    };
    const kept = regions.filter((region) => {
      const span = this.regionToRange(region);
      return span !== null && !rangesIntersect(span, target);
    });
    if (kept.length === regions.length) return; // nothing changed
    this.config.mergedRegions = kept;
    this.notifyStructureChange();
  }

  // ============================================
  // Permute / replace order (used by sort and history restore)
  // ============================================

  /** Replace rowOrder/colOrder + reverse maps wholesale. Cells are untouched. */
  replaceOrderMaps(
    rowOrder: Map<number, string>,
    colOrder: Map<number, string>,
  ): void {
    this.rowOrder = new Map(rowOrder);
    this.colOrder = new Map(colOrder);
    this.rowIdToIndex = new Map();
    this.colIdToIndex = new Map();
    for (const [idx, id] of this.rowOrder) this.rowIdToIndex.set(id, idx);
    for (const [idx, id] of this.colOrder) this.colIdToIndex.set(id, idx);
    this.notifyStructureChange();
  }

  /** Snapshot the current order maps (for history). */
  snapshotOrderMaps(): { rowOrder: Map<number, string>; colOrder: Map<number, string> } {
    return {
      rowOrder: new Map(this.rowOrder),
      colOrder: new Map(this.colOrder),
    };
  }
}
