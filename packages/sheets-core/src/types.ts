// Core types for Spindle Sheets

import type { CommentStore, SheetCommentThread } from './comments';

export type CellValue = string | number | boolean | null;

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  fontFamily?: string;
  fontSize?: number;
  fontColor?: string;
  backgroundColor?: string;
  textAlign?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  borderTop?: string;
  borderRight?: string;
  borderBottom?: string;
  borderLeft?: string;
  textWrap?: boolean;
  textDecoration?: 'none' | 'underline' | 'line-through';
  textRotation?: number; // degrees, -90..90; 0/undefined = horizontal
}

export type FormatType =
  | 'text'
  | 'number'
  | 'currency'
  | 'accounting'
  | 'percentage'
  | 'scientific'
  | 'fraction'
  | 'date'
  | 'time'
  | 'datetime'
  | 'duration'
  | 'custom'
  | undefined;

export interface CellFormat {
  type?: FormatType;

  // Number format options
  decimalPlaces?: number; // 0-30
  useThousandsSeparator?: boolean;

  // Currency format options
  currencyCode?: string; // USD, EUR, GBP, JPY, etc.
  currencySymbolPosition?: 'prefix' | 'suffix';

  // Negative number display options
  negativeFormat?: 'minus' | 'parentheses' | 'red';

  // Date/Time format options
  dateFormat?: string; // e.g., 'MM/DD/YYYY', 'DD-MM-YYYY'
  timeFormat?: string; // e.g., 'HH:mm:ss', 'h:mm AM/PM'

  // Fraction format options
  fractionType?: 'upToOne' | 'upToTwo' | 'upToThree' | 'asHalves' | 'asQuarters' | 'asEighths' | 'asSixteenths' | 'asTenths' | 'asHundredths';

  // Duration format options
  durationFormat?: 'hours' | 'minutes' | 'seconds' | 'milliseconds';

  // Custom format pattern
  pattern?: string; // Custom number format pattern (e.g., "#,##0.00", "0.0%")
}

export interface Cell {
  value: CellValue;
  /**
   * Display string for the formula (A1 notation). When `formulaAst` is also
   * present, the AST is the source of truth and `formula` is a cache that
   * gets regenerated on every evaluation. A cell with only `formula` and no
   * `formulaAst` (e.g. set via legacy paths) is upgraded to a stable AST on
   * first evaluation.
   */
  formula?: string;
  /**
   * Stable-ID formula AST. Source of truth for the formula's structure; refs
   * point at rowId/colId so they survive insert/delete/sort. Imported lazily
   * to avoid pulling formula-parser types into every consumer of Cell.
   */
  formulaAst?: import('./formula-parser/stable-ast').StableFormulaNode;
  styleId?: string; // Reference to shared style in StylePool
  formatId?: string; // Reference to shared format in FormatPool
  hyperlink?: string;
}

export interface Range {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface Selection {
  ranges: Range[];
  activeCell: { row: number; col: number };
}

/**
 * A merged cell region, anchored by the stable IDs of its two corner cells so
 * it survives insert/delete/sort (mirrors how comments anchor).
 */
export interface MergedRegion {
  startRowId: string;
  startColId: string;
  endRowId: string;
  endColId: string;
}

export interface SortOrder {
  column: number;
  direction: 'asc' | 'desc';
}

export interface ColumnFilter {
  column: number;
  type: 'text' | 'number' | 'date';
  criteria: FilterCriteria;
}

export type FilterCriteria =
  | { type: 'equals'; value: string | number }
  | { type: 'notEquals'; value: string | number }
  | { type: 'contains'; value: string }
  | { type: 'notContains'; value: string }
  | { type: 'startsWith'; value: string }
  | { type: 'endsWith'; value: string }
  | { type: 'greaterThan'; value: number }
  | { type: 'lessThan'; value: number }
  | { type: 'greaterThanOrEqual'; value: number }
  | { type: 'lessThanOrEqual'; value: number }
  | { type: 'between'; min: number; max: number }
  | { type: 'custom'; values: Set<string | number> }; // For multi-select

export interface SheetConfig {
  defaultRowHeight?: number;
  defaultColWidth?: number;
  // Row/column config keyed by stable rowId/colId (not numeric index).
  // Translate at the public API surface via SheetImpl.getRowId / getColId.
  rowHeights?: Map<string, number>; // rowId -> height
  colWidths?: Map<string, number>;  // colId -> width
  hiddenRows?: Set<string>;          // set of rowIds
  hiddenCols?: Set<string>;          // set of colIds
  frozenRows?: number;               // count; index-based by definition
  frozenCols?: number;               // count; index-based by definition
  showGridLines?: boolean;
  sortOrder?: SortOrder[];           // criteria still reference numeric columns (revisit in v2)
  filters?: Map<string, ColumnFilter>; // colId -> filter
  mergedRegions?: MergedRegion[];    // merged cell regions, keyed by stable IDs
}

export interface Sheet {
  id: string;
  name: string;
  cells: Map<string, Cell>; // key: "rowId:colId" (stable IDs — translate via SheetImpl helpers)
  config: SheetConfig;
  rowCount: number;
  colCount: number;
  /** Comment threads anchored to cells in this sheet. */
  comments: CommentStore;

  // Stable-ID helpers (translate between numeric indices and stable rowId/colId)
  getRowId(row: number): string | undefined;
  getColId(col: number): string | undefined;
  ensureRowId(row: number): string;
  ensureColId(col: number): string;
  getRowIndex(rowId: string): number | undefined;
  getColIndex(colId: string): number | undefined;
  stableKeyToIndices(key: string): { row: number; col: number } | undefined;

  getCell(row: number, col: number): Cell | undefined;
  setCell(row: number, col: number, cell: Partial<Cell>): void;
  deleteCell(row: number, col: number): void;
  getCellValue(row: number, col: number): unknown;
  setCellValue(row: number, col: number, value: unknown): void;
  getRange(range: Range): Map<string, Cell>;
  setRange(range: Range, cells: Map<string, Cell> | Cell[][]): void;
  clearRange(range: Range): void;
  entries(): IterableIterator<[number, number, Cell]>;
  getRowHeight(row: number): number;
  setRowHeight(row: number, height: number): void;
  getColWidth(col: number): number;
  setColWidth(col: number, width: number): void;
  insertRows(startRow: number, count: number): void;
  deleteRows(startRow: number, count: number): void;
  insertCols(startCol: number, count: number): void;
  deleteCols(startCol: number, count: number): void;
  isRowHidden(row: number): boolean;
  hideRow(row: number): void;
  showRow(row: number): void;
  isColHidden(col: number): boolean;
  hideCol(col: number): void;
  showCol(col: number): void;
  getHiddenColsAdjacent(col: number): { before: number[]; after: number[] };
  getHiddenRowsAdjacent(row: number): { above: number[]; below: number[] };
  showColsInRange(startCol: number, endCol: number): void;
  showRowsInRange(startRow: number, endRow: number): void;
  getFrozenRows(): number;
  setFrozenRows(count: number): void;
  getFrozenCols(): number;
  setFrozenCols(count: number): void;
  setFreeze(rows: number, cols: number): void;
  clearFreeze(): void;
  hasFrozenPanes(): boolean;
  setSortOrder(sortOrder: SortOrder[]): void;
  getSortOrder(): SortOrder[];
  clearSort(): void;
  hasSort(): boolean;
  setFilter(column: number, filter: ColumnFilter): void;
  clearFilter(column: number): void;
  getFilters(): Map<number, ColumnFilter>;
  hasFilter(column: number): boolean;
  clearAllFilters(): void;

  // Merged cells. Ranges are returned/accepted as numeric indices; storage is
  // by stable ID so regions survive insert/delete/sort.
  getMergedRegions(): Range[];
  mergeCells(range: Range): void;
  unmergeCells(range: Range): void;
  getMergeAt(row: number, col: number): Range | undefined;

  // Order-map plumbing (used by sort, history snapshot/restore, and the
  // future CRDT binding). Cells are not touched by these calls.
  replaceOrderMaps(rowOrder: Map<number, string>, colOrder: Map<number, string>): void;
  snapshotOrderMaps(): { rowOrder: Map<number, string>; colOrder: Map<number, string> };
}

export interface Workbook {
  id: string;
  name: string;
  sheets: Map<string, Sheet>;
  activeSheetId: string;
  defaultRowHeight: number;
  defaultColWidth: number;

  // Undo/Redo
  recordHistory(): void;
  undo(): boolean;
  redo(): boolean;
  canUndo(): boolean;
  canRedo(): boolean;

  // Sorting
  setSortOrder(sortOrder: SortOrder[], sheetId?: string): void;
  getSortOrder(sheetId?: string): SortOrder[];
  clearSort(sheetId?: string): void;
  sortSheet(sheetId?: string): void;

  // Filtering
  setFilter(column: number, filter: ColumnFilter, sheetId?: string): void;
  clearFilter(column: number, sheetId?: string): void;
  getFilters(sheetId?: string): Map<number, ColumnFilter>;
  clearAllFilters(sheetId?: string): void;

  // Merged cells
  mergeCells(range: Range, sheetId?: string): void;
  unmergeCells(range: Range, sheetId?: string): void;

  // Data serialization/deserialization
  getData(): WorkbookData;
  setData(data: WorkbookData): void;
}

export interface StylePool {
  styles: Map<string, CellStyle>;
  getOrCreate(style: CellStyle): string;
  get(styleId: string): CellStyle | undefined;
  getAllStyles(): Map<string, CellStyle>;
}

export interface FormatPool {
  getOrCreate(format: CellFormat): string;
  get(formatId: string): CellFormat | undefined;
  getAllFormats(): Map<string, CellFormat>;
}

// Workbook data serialization interfaces
export interface WorkbookData {
  id: string;
  name: string;
  activeSheetId: string;
  defaultRowHeight: number;
  defaultColWidth: number;
  stylePool: Record<string, CellStyle>; // styleId -> style object
  formatPool?: Record<string, CellFormat>; // formatId -> format object (optional for backward compatibility)
  sheets: SheetData[];
  selection?: Selection; // Optional UI state
}

export interface SheetData {
  id: string;
  name: string;
  /**
   * Cell entries. The key format depends on whether `rowOrder`/`colOrder`
   * are present:
   *   - With rowOrder/colOrder: key is `"{rowId}:{colId}"` (stable). This is
   *     what `getData()` emits.
   *   - Without: key is `"row:col"` (numeric). This legacy form is accepted
   *     by `setData()` and is convenient for hand- or AI-authored JSON.
   */
  cells: Array<{ key: string; cell: Cell }>;
  /**
   * Stable row IDs at each numeric index, sparse. Optional on input — when
   * absent, cell keys / config keys are interpreted as numeric and IDs are
   * materialized during load. Always emitted by `getData()`.
   */
  rowOrder?: Array<[number, string]>;
  colOrder?: Array<[number, string]>;
  config: {
    defaultRowHeight?: number;
    defaultColWidth?: number;
    /** When rowOrder/colOrder is present, keys are rowId/colId strings; otherwise numeric indices. */
    rowHeights?: Array<[number | string, number]>;
    colWidths?: Array<[number | string, number]>;
    hiddenRows?: Array<number | string>;
    hiddenCols?: Array<number | string>;
    frozenRows?: number;
    frozenCols?: number;
    showGridLines?: boolean;
    sortOrder?: SortOrder[];
    filters?: Array<[number | string, ColumnFilter]>;
    /** Merged regions, always keyed by stable IDs (no numeric legacy form). */
    mergedRegions?: MergedRegion[];
  };
  rowCount: number;
  colCount: number;
  /** Comment threads anchored to cells in this sheet. */
  threads?: SheetCommentThread[];
}

/** A range dependency, stored as its two corner cell keys (stable ids). */
export interface RangeDependency {
  startKey: string; // "rowId:colId"
  endKey: string; // "rowId:colId"
}

/** The dependencies a formula reads: single cells, plus ranges as rectangles. */
export interface FormulaDependencies {
  cells: Set<string>;
  ranges: RangeDependency[];
}

/**
 * What a formula on a dependency cycle evaluates to: `'zero'` (Excel: 0, plus a
 * 'circularReference' event), `'ref'` (#REF!) or `'error'` (the engine's
 * diagnostic #CIRCULAR! token, which is never written to XLSX).
 */
export type CircularReferenceMode = 'zero' | 'ref' | 'error';

export interface FormulaNode {
  cellKey: string;
  formula: string;
  dependencies: Set<string>; // single-cell deps "rowId:colId"
  rangeRects: RangeDependency[]; // range deps as stable corner-key rectangles
  cachedValue?: CellValue;
  isDirty: boolean;
  volatile?: boolean; // recomputed on every recalc pass (RAND, NOW, OFFSET, …)
  /** Cells a blocked (#SPILL!) array formula would fill: re-evaluate it when any of them changes. */
  spillRects?: RangeDependency[];
}

/**
 * Resolves a stable cell key to its current position, including which sheet
 * owns it. Cell keys are globally unique, so this identifies exactly one cell
 * across the whole workbook — the sheetId lets range containment reject a
 * same-index cell on a different sheet.
 */
export type CellResolver = (cellKey: string) => { sheetId: string; row: number; col: number } | undefined;

export interface FormulaGraph {
  nodes: Map<string, FormulaNode>;
  addFormula(
    cellKey: string,
    formula: string,
    dependencies: FormulaDependencies,
    volatile?: boolean
  ): void;
  removeFormula(cellKey: string): void;
  markClean(cellKey: string, value: CellValue): void;
  /** Make a formula depend on the cells its array result would spill into (empty list clears it). */
  setSpillFootprint(cellKey: string, rects: RangeDependency[]): void;
  /** Invalidate every formula's cached value (used before a full re-evaluation). */
  markAllDirty(): void;
  /**
   * Collect every cell to recompute after the `seeds` changed — their transitive
   * dependents (single-cell and range-based) plus all volatile cells — along
   * with the dependency edges among them.
   */
  collectDirty(
    seeds: string[],
    resolveCell: CellResolver
  ): { dirty: Set<string>; edges: Map<string, Set<string>> };
  /** Order a dirty set so dependencies precede dependents; cyclic nodes can't be ordered. */
  topologicalOrder(
    dirty: Set<string>,
    edges: Map<string, Set<string>>
  ): { ordered: string[]; cyclic: string[] };
}

export type EventType =
  | 'cellChange'
  | 'cellSelection'
  | 'sheetChange'
  | 'sheetAdd'
  | 'sheetDelete'
  | 'sheetRename'
  | 'commentChange'
  | 'commentEvent'
  | 'workbookChange'
  | 'circularReference';

export interface EventData {
  type: EventType;
  payload: unknown;
}

export type EventHandler = (data: EventData) => void;

