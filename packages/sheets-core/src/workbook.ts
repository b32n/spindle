// Workbook model

import type { Workbook, Sheet, Cell, Range, MergedRegion, Selection, CellValue, SortOrder, CellFormat, CellStyle, EventType, EventHandler } from './types';
import { SheetImpl } from './sheet';
import { EventEmitter } from './event-emitter';
import { FormulaGraphImpl } from './formula-graph';
import { SpillIndex } from './spill';
import type { SpillRegion } from './spill';
import { StylePool } from './style-pool';
import { FormatPool } from './format-pool';
import { SortManager } from './features/sort';
import { getCellKey, getStableCellKey, parseCellKey } from './utils/cell-key';
import { FormulaParser } from './formula-parser';
import type { RangeReference, StableFormulaNode, SheetResolver } from './formula-parser';
import * as Y from 'yjs';
import { attachCollabToWorkbook } from './collab/binding';
import type { CollabHandle, AttachCollabOptions } from './collab/binding';
import {
  getWorkbookYTypes,
  getSheetYTypes,
  ensureSheetYMap,
  threadKey,
  mergedRegionKey,
} from './collab/y-schema';
import type { CommentStore, CommentMutationEvent, SheetCommentThread } from './comments';
import {
  toStableAst,
  fromStableAst,
  renderStableAst,
  collectStableDependencies,
  formulaIsVolatile,
  StableRefDeletedError,
} from './formula-parser/stable-ast';

// Snapshot for undo/redo
interface WorkbookSnapshot {
  sheets: Map<string, SheetSnapshot>;
  activeSheetId: string;
  selection: Selection;
}

interface SheetSnapshot {
  id: string;
  name: string;
  cells: Map<string, Cell>; // key: "rowId:colId"
  rowOrder: Map<number, string>;
  colOrder: Map<number, string>;
  config: Sheet['config'];
  rowCount: number;
  colCount: number;
}

export class WorkbookImpl implements Workbook {
  id: string;
  name: string;
  sheets: Map<string, Sheet> = new Map();
  activeSheetId: string;
  defaultRowHeight: number;
  defaultColWidth: number;

  private events: EventEmitter = new EventEmitter();
  private formulaGraph: FormulaGraphImpl = new FormulaGraphImpl();
  private spillIndexes: Map<string, SpillIndex> = new Map(); // per-sheet spill overlay
  private spillSeeds: string[] = []; // footprint cells to recheck after a spill changes
  private stylePool: StylePool = new StylePool();
  private formatPool: FormatPool = new FormatPool();
  private formulaParser: FormulaParser = new FormulaParser();
  private selection: Selection = {
    ranges: [],
    activeCell: { row: 0, col: 0 },
  };
  private evaluatingCells: Set<string> = new Set(); // Track cells being evaluated to detect circular references
  private undoStack: WorkbookSnapshot[] = [];
  private redoStack: WorkbookSnapshot[] = [];
  private maxHistorySize = 50; // Limit history size to prevent memory issues
  private isUndoing = false; // Flag to prevent recording history during undo/redo
  private isRedoing = false;
  private isBatching = false; // Flag to track if we're in a batch operation
  // Set by attachCollab; consulted in places that need to know whether the
  // workbook is currently sourcing state from a Y.Doc.
  private collabHandle: CollabHandle | null = null;
  // True while _reloadFromCollab is rebuilding internal state from a remote
  // Y.Doc update. Mirror calls in mutation methods short-circuit when set,
  // otherwise we'd echo every remote write back over the wire.
  private isApplyingRemoteChange = false;

  constructor(id: string, name: string) {
    this.id = id;
    this.name = name;
    this.defaultRowHeight = 20;
    this.defaultColWidth = 100;

    // Create default sheet
    const defaultSheet = this.addSheet('Sheet1');
    this.activeSheetId = defaultSheet.id;
  }

  addSheet(name: string): Sheet {
    const id = `sheet_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    const sheet = new SheetImpl(id, name);
    this.sheets.set(id, sheet);
    this.wireCommentListener(sheet);
    this.attachStructureListenerIfNeeded(sheet);
    this.mirrorSheetAdd(id, name, sheet);
    // A formula that referenced this sheet by name before it existed holds a
    // deferred (#REF!) ref with only a sentinel dependency. Re-parse those
    // formulas now that the sheet exists so the ref binds to real ids and a real
    // dependency — recovering the value AND reactive tracking of the target cell.
    // Guarded to >1 sheet so the constructor's first addSheet is a no-op.
    if (this.sheets.size > 1 && !this.isApplyingRemoteChange) {
      this.rebindFormulasReferencingSheet(name);
    }
    this.events.emit('sheetAdd', { sheetId: id, name });
    return sheet;
  }

  /**
   * Re-parse every formula whose text mentions `sheetName`, re-registering its
   * dependencies and recomputing. Used when a sheet is (re)created so a formula
   * that referenced it while it was absent stops being a permanent #REF!.
   */
  private rebindFormulasReferencingSheet(sheetName: string): void {
    const needle = sheetName.toLowerCase();
    // Collect first — setFormula mutates cells, so don't re-parse mid-iteration.
    const targets: Array<{ sheetId: string; row: number; col: number; formula: string }> = [];
    for (const [sheetId, sheet] of this.sheets.entries()) {
      for (const [row, col, cell] of sheet.entries()) {
        if (cell.formula && cell.formula.toLowerCase().includes(needle)) {
          targets.push({ sheetId, row, col, formula: cell.formula });
        }
      }
    }
    if (targets.length === 0) return;
    const wasUndoing = this.isUndoing;
    this.isUndoing = true; // re-binding is not a user edit — don't record history
    try {
      for (const t of targets) this.setFormula(t.sheetId, t.row, t.col, t.formula);
    } finally {
      this.isUndoing = wasUndoing;
    }
  }

  /**
   * Return the comment store for a sheet (defaults to the active sheet).
   * Re-fetch this after any reload — a remote update replaces the SheetImpl.
   */
  getCommentStore(sheetId?: string): CommentStore {
    return this.getSheet(sheetId).comments;
  }

  deleteSheet(sheetId: string): void {
    if (this.sheets.size <= 1) {
      throw new Error('Cannot delete the last sheet');
    }
    this.sheets.delete(sheetId);
    if (this.activeSheetId === sheetId) {
      // Switch to first available sheet (local view state).
      this.activeSheetId = Array.from(this.sheets.keys())[0];
    }
    this.mirrorSheetDelete(sheetId);
    // Formulas on other sheets that referenced the deleted one must now
    // recompute to #REF!.
    this.recalculateAllFormulas();
    this.events.emit('sheetDelete', { sheetId });
  }

  getSheet(sheetId?: string): SheetImpl {
    const id = sheetId ?? this.activeSheetId;
    const sheet = this.sheets.get(id);
    if (!sheet) {
      throw new Error(`Sheet not found: ${id}`);
    }
    return sheet as SheetImpl;
  }

  setActiveSheet(sheetId: string): void {
    if (!this.sheets.has(sheetId)) {
      throw new Error(`Sheet not found: ${sheetId}`);
    }
    this.activeSheetId = sheetId;
    // No mirror: which sheet a user is viewing is per-user view state.
    this.events.emit('sheetChange', { sheetId });
  }

  renameSheet(sheetId: string, newName: string): void {
    const sheet = this.getSheet(sheetId);
    const oldName = sheet.name;
    if (oldName === newName) return;
    sheet.name = newName;
    // Rewrite cross-sheet references so they follow the rename instead of
    // breaking to #REF!, then recompute the affected formulas.
    this.rewriteSheetNameInFormulas(oldName, newName);
    this.mirrorSheetRename(sheetId, newName);
    this.recalculateAllFormulas();
    this.events.emit('sheetRename', { sheetId, oldName, newName });
  }

  getSheetIdByName(sheetName: string): string | undefined {
    for (const [id, sheet] of this.sheets.entries()) {
      if (sheet.name === sheetName) {
        return id;
      }
    }
    return undefined;
  }

  getCell(sheetId: string | undefined, row: number, col: number): Cell | undefined {
    const sheet = this.getSheet(sheetId);
    return sheet.getCell(row, col);
  }

  /**
   * Build the FormulaGraph key for a (row, col) on the given sheet,
   * lazily materializing stable IDs if needed.
   */
  private graphKey(sheetId: string | undefined, row: number, col: number): string {
    const sheet = this.getSheet(sheetId);
    return getStableCellKey(sheet.ensureRowId(row), sheet.ensureColId(col));
  }

  /**
   * Resolver passed to the stable-AST converters so cross-sheet refs can
   * look up the target sheet's order maps.
   */
  private buildSheetResolver(): SheetResolver {
    return {
      getSheet: (name?: string): Sheet | undefined => {
        if (!name) return undefined;
        const id = this.getSheetIdByName(name);
        return id ? this.sheets.get(id) : undefined;
      },
    };
  }

  setCell(
    sheetId: string | undefined,
    row: number,
    col: number,
    cell: Partial<Cell>
  ): void {
    // Handle style pooling before storing
    if (cell.styleId) {
      // Style already pooled
    } else if (cell.styleId === undefined && 'style' in cell) {
      // Need to pool the style
      const styleId = this.stylePool.getOrCreate(cell.style as CellStyle);
      cell.styleId = styleId;
      delete (cell as Partial<Cell> & { style?: CellStyle }).style;
    }

    // Handle format pooling before storing
    if (cell.formatId) {
      // Format already pooled
    } else if (cell.formatId === undefined && 'format' in cell) {
      // Clean the format before pooling to ensure consistent keys
      const cleanedFormat = this.cleanFormat(cell.format as CellFormat);
      const formatId = this.formatPool.getOrCreate(cleanedFormat);
      cell.formatId = formatId;
      delete (cell as Partial<Cell> & { format?: CellFormat }).format;
    }

    const sheet = this.getSheet(sheetId);
    sheet.setCell(row, col, cell);

    this.mirrorCellWrite(sheet.id, row, col);

    this.events.emit('cellChange', {
      sheetId: sheet.id,
      row,
      col,
      cellKey: getCellKey(row, col),
    });
  }

  setCellValue(sheetId: string | undefined, row: number, col: number, value: unknown): void {
    // Record history before making changes (unless we're undoing/redoing or in a batch)
    if (!this.isUndoing && !this.isRedoing && !this.isBatching) {
      const currentCell = this.getCell(sheetId, row, col);
      const currentValue = currentCell?.value;
      if (currentValue !== value) {
        this.recordHistory();
      }
    }
    
    const cellKey = this.graphKey(sheetId, row, col);
    const hadFormula = this.formulaGraph.nodes.has(cellKey);

    // Remove formula if setting a direct value
    if (hadFormula) {
      this.formulaGraph.removeFormula(cellKey);
    }
    // Release any spill this cell anchored — its overlay must go.
    this.releaseSpill(sheetId, row, col);

    // Clear formula/formulaAst so a plain value never coexists with a stale
    // formula on the same cell.
    this.setCell(sheetId, row, col, {
      value: value as string | number | boolean | null,
      formula: undefined,
      formulaAst: undefined,
    });

    // Recalculate dependents of this cell
    if (hadFormula || this.getCell(sheetId, row, col)?.value !== undefined) {
      this.recalculateDependents(cellKey, sheetId);
    }
  }

  getCellValue(sheetId: string | undefined, row: number, col: number): unknown {
    const cell = this.getCell(sheetId, row, col);
    return cell?.value ?? null;
  }

  /** The value shown at a cell filled by a dynamic-array spill, if any. */
  getSpilledValue(sheetId: string | undefined, row: number, col: number): unknown | undefined {
    return this.getSpillIndex(sheetId).spilledValueAt(row, col);
  }

  /** Whether a cell is filled by a spill it does not itself anchor (read-only). */
  isSpilledCell(sheetId: string | undefined, row: number, col: number): boolean {
    return this.getSpillIndex(sheetId).isCovered(row, col);
  }

  /** The anchor position of the spill covering a cell, if any. */
  getSpillAnchor(
    sheetId: string | undefined,
    row: number,
    col: number
  ): { row: number; col: number } | undefined {
    return this.getSpillIndex(sheetId).anchorOf(row, col);
  }

  setFormula(sheetId: string | undefined, row: number, col: number, formula: string): void {
    // Record history before making changes (unless we're undoing/redoing or in a batch)
    if (!this.isUndoing && !this.isRedoing && !this.isBatching) {
      const currentCell = this.getCell(sheetId, row, col);
      const currentFormula = currentCell?.formula;
      if (currentFormula !== formula) {
        this.recordHistory();
      }
    }
    
    const cellKey = this.graphKey(sheetId, row, col);
    const sheet = this.getSheet(sheetId);

    // Parse to numeric AST, then convert to stable AST (the source of truth).
    const parseResult = this.formulaParser.parse(formula, row, col);

    if (parseResult.error) {
      this.setCell(sheetId, row, col, { formula, value: parseResult.error as CellValue });
      return;
    }

    const resolver = this.buildSheetResolver();
    const formulaAst = toStableAst(parseResult.ast, resolver, sheet, row, col);
    const stableDeps = collectStableDependencies(formulaAst);

    this.formulaGraph.addFormula(cellKey, formula, stableDeps, formulaIsVolatile(formulaAst));

    // Store both AST (truth) and input string (display cache; gets refreshed
    // from the AST on every evaluation).
    this.setCell(sheetId, row, col, { formula, formulaAst });

    // Evaluate and store result
    this.evaluateFormula(sheetId, row, col);

    // Recalculate dependents
    this.recalculateDependents(cellKey, sheetId);
  }

  getSelection(): Selection {
    return { ...this.selection };
  }

  setSelection(selection: Selection): void {
    this.selection = selection;
    this.events.emit('cellSelection', { selection });
  }

  batch(operations: () => void): void {
    // Record history before batch operations (unless we're undoing/redoing)
    if (!this.isUndoing && !this.isRedoing) {
      this.isBatching = true;
      this.recordHistory();
    }
    this.events.batch(operations);
    this.isBatching = false;
  }

  on(event: EventType, handler: EventHandler): () => void {
    return this.events.on(event, handler);
  }

  off(event: EventType, handler: EventHandler): void {
    this.events.off(event, handler);
  }

  getFormulaGraph(): FormulaGraphImpl {
    return this.formulaGraph;
  }

  getStylePool(): StylePool {
    return this.stylePool;
  }

  getFormatPool(): FormatPool {
    return this.formatPool;
  }

  /**
   * Get the calculated value of a cell (evaluates formulas)
   */
  getCellCalculatedValue(sheetId: string | undefined, row: number, col: number): unknown {
    const cell = this.getCell(sheetId, row, col);
    const cellKey = this.graphKey(sheetId, row, col);
    
    // If cell has a formula, evaluate it
    if (cell?.formula) {
      const formulaNode = this.formulaGraph.nodes.get(cellKey);
      
      // Return cached value if available and not dirty
      if (formulaNode && !formulaNode.isDirty && formulaNode.cachedValue !== undefined) {
        return formulaNode.cachedValue;
      }
      
      // Check for circular reference BEFORE evaluating
      if (this.evaluatingCells.has(cellKey)) {
        // This cell is already being evaluated - circular reference!
        return '#CIRCULAR!';
      }
      
      // Evaluate formula
      return this.evaluateFormula(sheetId, row, col);
    }

    // A cell with no formula of its own may still be covered by a spill.
    const spilled = this.getSpillIndex(sheetId).spilledValueAt(row, col);
    if (spilled !== undefined) return spilled;

    // Return direct value
    // For empty cells, return 0 (standard spreadsheet behavior for numeric formulas)
    // For cells with values, return the actual value (string, number, boolean, etc.)
    if (!cell) return 0;
    
    // Return the actual value - don't convert null/undefined to 0 here
    // The formula parser will handle type conversion as needed
    if (cell.value === null || cell.value === undefined) {
      return 0; // Empty cells return 0 for numeric calculations
    }
    
    return cell.value;
  }

  /**
   * Evaluate a formula and store the result
   */
  private evaluateFormula(sheetId: string | undefined, row: number, col: number): unknown {
    const cellKey = this.graphKey(sheetId, row, col);
    const cell = this.getCell(sheetId, row, col);
    
    if (!cell?.formula) {
      return cell?.value ?? null;
    }

    // Check for circular reference BEFORE adding to evaluatingCells
    if (this.evaluatingCells.has(cellKey)) {
      const error = '#CIRCULAR!';
      this.formulaGraph.markClean(cellKey, error);
      this.setCell(sheetId, row, col, { value: error });
      return error;
    }

    // Add to evaluating cells to detect circular references
    this.evaluatingCells.add(cellKey);

    try {
      const sheet = this.getSheet(sheetId);
      const resolver = this.buildSheetResolver();

      // Source of truth is cell.formulaAst (stable). If only the legacy
      // string is present, parse it and upgrade in place.
      let stableAst: StableFormulaNode | undefined = cell.formulaAst;
      if (!stableAst) {
        const parsed = this.formulaParser.parse(cell.formula, row, col);
        if (parsed.error) {
          this.evaluatingCells.delete(cellKey);
          this.formulaGraph.markClean(cellKey, parsed.error as CellValue);
          this.setCell(sheetId, row, col, { value: parsed.error as CellValue });
          return parsed.error;
        }
        stableAst = toStableAst(parsed.ast, resolver, sheet, row, col);
        // Persist the upgraded AST so future evaluations skip the re-parse.
        this.setCell(sheetId, row, col, { formulaAst: stableAst });
      }

      // Convert stable refs back to numeric for the evaluator. Any ref whose
      // target row/col ID has been deleted bubbles up as a #REF! error.
      let numericAst;
      try {
        numericAst = fromStableAst(stableAst, resolver, sheet, row, col);
      } catch (e) {
        if (e instanceof StableRefDeletedError) {
          this.evaluatingCells.delete(cellKey);
          this.formulaGraph.markClean(cellKey, '#REF!' as CellValue);
          this.setCell(sheetId, row, col, {
            value: '#REF!' as CellValue,
            formula: renderStableAst(stableAst, resolver, sheet),
          });
          return '#REF!';
        }
        throw e;
      }

      // Refresh the display string from the AST so the user sees A1
      // notation that tracks structural edits (insert/delete/sort).
      const refreshedFormula = renderStableAst(stableAst, resolver, sheet);

      const evaluationContext = {
        getCellValue: (r: number, c: number, sId?: string, sheetName?: string) => {
          let targetSheetId = sId ?? sheetId;
          if (sheetName) {
            const resolvedSheetId = this.getSheetIdByName(sheetName);
            if (resolvedSheetId) {
              targetSheetId = resolvedSheetId;
            } else {
              return '#REF!';
            }
          }
          return this.getCellCalculatedValue(targetSheetId, r, c);
        },
        getRangeValues: (range: RangeReference, sId?: string, sheetName?: string) => {
          let targetSheetId = sId ?? sheetId;
          if (sheetName) {
            const resolvedSheetId = this.getSheetIdByName(sheetName);
            if (resolvedSheetId) {
              targetSheetId = resolvedSheetId;
            } else {
              return [];
            }
          }
          return this.getRangeValues(targetSheetId, range, row, col);
        },
        getSheetIdByName: (sheetName: string) => {
          return this.getSheetIdByName(sheetName);
        },
        isCellFormula: (r: number, c: number, sheetName?: string) => {
          let targetSheetId: string | undefined = sheetId;
          if (sheetName) {
            const resolvedSheetId = this.getSheetIdByName(sheetName);
            if (!resolvedSheetId) return false;
            targetSheetId = resolvedSheetId;
          }
          const target = this.getCell(targetSheetId, r, c);
          return !!(target && (target.formula || target.formulaAst));
        },
      };

      const result = this.formulaParser.evaluate(numericAst, evaluationContext, row, col);
      this.evaluatingCells.delete(cellKey);

      // An array result spills into a block of cells.
      if (Array.isArray(result)) {
        return this.applySpill(sheetId, row, col, cellKey, result as unknown[][], refreshedFormula);
      }
      // No longer an array result — release any spill region this cell anchored.
      this.releaseSpill(sheetId, row, col);

      // Handle division by zero / other non-finite results.
      if (typeof result === 'number' && !isFinite(result)) {
        const error = result === Infinity || result === -Infinity ? '#DIV/0!' : '#NUM!';
        this.formulaGraph.markClean(cellKey, error);
        this.setCell(sheetId, row, col, { value: error });
        return error;
      }

      // Store the scalar result with the refreshed display string.
      return this.storeFormulaResult(
        sheetId,
        row,
        col,
        cellKey,
        result as CellValue,
        refreshedFormula
      );
    } catch (error) {
      this.evaluatingCells.delete(cellKey);
      this.releaseSpill(sheetId, row, col);
      const raw = error instanceof Error ? error.message : '';
      // Only surface a spreadsheet error token (starts with '#'). A raw JS
      // exception message (e.g. "Invalid string length" from REPT) would
      // otherwise be stored as the cell value; normalize it to #VALUE!.
      const errorMsg = raw.startsWith('#') ? raw : '#VALUE!';
      this.formulaGraph.markClean(cellKey, errorMsg as CellValue);
      this.setCell(sheetId, row, col, { value: errorMsg as CellValue });
      return errorMsg;
    }
  }

  /**
   * Get values from a range for formula evaluation
   * Note: Range references are already resolved to absolute positions during parsing,
   * but we keep the currentRow/currentCol parameters for consistency with the API
   */
  private getRangeValues(
    sheetId: string | undefined,
    range: RangeReference,
    currentRow: number,
    currentCol: number
  ): unknown[][] {
    const { start, end } = range;
    
    // Resolve relative references to absolute positions
    const startRow = start.rowAbsolute ? start.row : currentRow + start.row;
    const endRow = end.rowAbsolute ? end.row : currentRow + end.row;
    const startCol = start.colAbsolute ? start.col : currentCol + start.col;
    const endCol = end.colAbsolute ? end.col : currentCol + end.col;
    
    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);
    const minCol = Math.min(startCol, endCol);
    const maxCol = Math.max(startCol, endCol);

    // Guard against an enormous range (e.g. =SUM(A1:A1048576)) freezing the main
    // thread. Cells past the sheet's populated extent are empty and contribute
    // nothing, so clamp the iteration to the actual data. Only pay the O(cells)
    // extent scan for an oversized range — normal ranges iterate directly.
    let effMaxRow = maxRow;
    let effMaxCol = maxCol;
    const MAX_RANGE_EVAL_CELLS = 100_000;
    if ((maxRow - minRow + 1) * (maxCol - minCol + 1) > MAX_RANGE_EVAL_CELLS) {
      const sheet = this.getSheet(sheetId);
      let dataMaxRow = 0;
      let dataMaxCol = 0;
      for (const [r, c] of sheet.entries()) {
        if (r > dataMaxRow) dataMaxRow = r;
        if (c > dataMaxCol) dataMaxCol = c;
      }
      effMaxRow = Math.min(maxRow, dataMaxRow);
      effMaxCol = Math.min(maxCol, dataMaxCol);
    }

    const values: unknown[][] = [];
    for (let row = minRow; row <= effMaxRow; row++) {
      const rowValues: unknown[] = [];
      for (let col = minCol; col <= effMaxCol; col++) {
        rowValues.push(this.getRangeCellValue(sheetId, row, col));
      }
      values.push(rowValues);
    }

    return values;
  }

  /**
   * Value of a cell as seen inside a range. Unlike getCellCalculatedValue this
   * yields `null` (not `0`) for a truly-empty cell, so blank-aware aggregates
   * (AVERAGE/COUNT/COUNTA/MIN/MEDIAN/COUNTBLANK/STDEV…) can distinguish an empty
   * cell from a real 0. Numeric coercion still treats null as 0 (toNum(null)===0),
   * so SUM and arithmetic are unaffected. Single-cell refs (=A1+B1) keep going
   * through getCellValue, which still returns 0 for a blank operand.
   */
  private getRangeCellValue(sheetId: string | undefined, row: number, col: number): unknown {
    const cell = this.getCell(sheetId, row, col);
    if (cell?.formula) return this.getCellCalculatedValue(sheetId, row, col);
    const spilled = this.getSpillIndex(sheetId).spilledValueAt(row, col);
    if (spilled !== undefined) return spilled;
    if (!cell || cell.value === null || cell.value === undefined) return null;
    return cell.value;
  }

  /**
   * Recalculate every cell that transitively depends on the given cell.
   * Each dependent is evaluated exactly once per pass, in topological order, so
   * a cell reachable by several paths is not recomputed several times. Cells in
   * (or downstream of) a dependency cycle are reported as #CIRCULAR!.
   *
   * When a recomputed formula spills (or stops spilling), the cells in its
   * footprint are fed back as seeds for another pass, so formulas that read a
   * spilled cell pick up the change. The loop runs to a fixed point.
   */
  private recalculateDependents(cellKey: string, _sheetId: string | undefined): void {
    // Cell keys are globally unique, so dependents (and range corners) are
    // resolved against the whole workbook — a formula on another sheet that
    // references this cell is found and re-evaluated on its own sheet.
    //
    // resolveGlobalCell is O(sheets); it's called per dirty key and per range
    // corner in the containment scan, so memoize it for the duration of this
    // recalc. Cell positions don't shift mid-recalc, so the cache stays valid.
    const locCache = new Map<string, { sheetId: string; row: number; col: number } | undefined>();
    const resolveCell = (key: string) => {
      let loc = locCache.get(key);
      if (loc === undefined && !locCache.has(key)) {
        loc = this.resolveGlobalCell(key);
        locCache.set(key, loc);
      }
      return loc;
    };

    let seeds = [cellKey, ...this.spillSeeds];
    this.spillSeeds = [];

    for (let pass = 0; pass < 64 && seeds.length > 0; pass++) {
      const { dirty, edges } = this.formulaGraph.collectDirty(seeds, resolveCell);
      if (dirty.size === 0) break;

      const { ordered, cyclic } = this.formulaGraph.topologicalOrder(dirty, edges);
      for (const key of ordered) {
        const loc = resolveCell(key);
        if (loc) {
          this.evaluateFormula(loc.sheetId, loc.row, loc.col);
        }
      }
      for (const key of cyclic) {
        const loc = resolveCell(key);
        if (!loc) continue;
        this.formulaGraph.markClean(key, '#CIRCULAR!' as CellValue);
        this.setCell(loc.sheetId, loc.row, loc.col, { value: '#CIRCULAR!' as CellValue });
      }

      // Spill footprints touched this pass become the seeds for the next.
      seeds = this.spillSeeds;
      this.spillSeeds = [];
    }
  }

  /**
   * Resolve a global cell key to its owning sheet and position. Keys are
   * globally unique, so at most one sheet matches. O(number of sheets), which
   * is fine for the handful of sheets a real workbook has.
   */
  private resolveGlobalCell(
    key: string,
  ): { sheetId: string; row: number; col: number } | undefined {
    for (const [id, sheet] of this.sheets) {
      const idx = sheet.stableKeyToIndices(key);
      if (idx) return { sheetId: id, row: idx.row, col: idx.col };
    }
    return undefined;
  }

  /**
   * Re-evaluate every formula in the workbook. Used after a structural sheet
   * change (delete/rename) that can invalidate cross-sheet references anywhere.
   */
  private recalculateAllFormulas(): void {
    // Clear the spill overlay first — it's keyed by absolute row/col, so a
    // structural edit (insert/delete row/col, sort) that shifts or removes a
    // spill anchor would otherwise leave orphaned coverage (ghost `isSpilled`
    // cells) and make a re-spill collide with its own stale footprint (`#SPILL!`).
    // Re-evaluating every formula below rebuilds it fresh, as setData does.
    this.spillIndexes.clear();
    // Invalidate every cached formula value first. evaluateFormula pulls its
    // dependencies through getCellCalculatedValue, which returns a node's cached
    // value while it is clean — so without this, a formula evaluated before a
    // dependency it forward-references (in arbitrary sheet.entries() order) would
    // read the pre-change cached value (e.g. a stale number instead of #REF!).
    this.formulaGraph.markAllDirty();
    for (const [sheetId, sheet] of this.sheets.entries()) {
      for (const [row, col, cell] of sheet.entries()) {
        if (cell.formula) this.evaluateFormula(sheetId, row, col);
      }
    }
  }

  /**
   * Rewrite every formula's references to a renamed sheet so they follow the
   * new name (Excel behavior) instead of breaking to #REF!. The display string
   * is refreshed from the AST on the next evaluation.
   */
  private rewriteSheetNameInFormulas(oldName: string, newName: string): void {
    const rewrite = (node: StableFormulaNode): void => {
      if (node.cellRef?.sheetName === oldName) node.cellRef.sheetName = newName;
      if (node.rangeRef?.start.sheetName === oldName) node.rangeRef.start.sheetName = newName;
      if (node.rangeRef?.end.sheetName === oldName) node.rangeRef.end.sheetName = newName;
      node.args?.forEach(rewrite);
      if (node.left) rewrite(node.left);
      if (node.right) rewrite(node.right);
    };
    for (const [, sheet] of this.sheets.entries()) {
      for (const [, , cell] of sheet.entries()) {
        if (cell.formulaAst) rewrite(cell.formulaAst);
      }
    }
  }

  /** The spill overlay for a sheet, created lazily. */
  private getSpillIndex(sheetId: string | undefined): SpillIndex {
    const id = this.getSheet(sheetId).id;
    let index = this.spillIndexes.get(id);
    if (!index) {
      index = new SpillIndex();
      this.spillIndexes.set(id, index);
    }
    return index;
  }

  /** Queue every cell of a spill region for the next recalculation pass. */
  private pushSpillSeeds(sheetId: string | undefined, region: SpillRegion): void {
    for (let r = 0; r < region.rows; r++) {
      for (let c = 0; c < region.cols; c++) {
        this.spillSeeds.push(this.graphKey(sheetId, region.anchorRow + r, region.anchorCol + c));
      }
    }
  }

  /** Release any spill region anchored at a cell (it no longer returns an array). */
  private releaseSpill(sheetId: string | undefined, row: number, col: number): void {
    const previous = this.getSpillIndex(sheetId).unregister(row, col);
    if (previous) this.pushSpillSeeds(sheetId, previous);
  }

  /** Store a scalar formula result on its cell and mark the graph node clean. */
  private storeFormulaResult(
    sheetId: string | undefined,
    row: number,
    col: number,
    cellKey: string,
    value: CellValue,
    refreshedFormula: string
  ): unknown {
    this.formulaGraph.markClean(cellKey, value);
    const currentCell = this.getCell(sheetId, row, col);
    if (currentCell) {
      this.setCell(sheetId, row, col, { ...currentCell, value, formula: refreshedFormula });
    }
    return value;
  }

  /**
   * Spill a formula's 2D array result into a block of cells. The anchor cell
   * keeps the formula and shows the array's top-left value; the rest of the
   * array is a derived overlay. A blocked target cell yields #SPILL!.
   */
  private applySpill(
    sheetId: string | undefined,
    row: number,
    col: number,
    cellKey: string,
    array: unknown[][],
    refreshedFormula: string
  ): unknown {
    const spill = this.getSpillIndex(sheetId);
    const matrix = array.length > 0 && Array.isArray(array[0]) ? array : [array];
    const rows = matrix.length;
    const cols = matrix[0]?.length ?? 0;

    const previous = spill.unregister(row, col);
    if (previous) this.pushSpillSeeds(sheetId, previous);

    // A 1x1 array is just a scalar.
    if (rows <= 1 && cols <= 1) {
      return this.storeFormulaResult(
        sheetId,
        row,
        col,
        cellKey,
        (matrix[0]?.[0] ?? null) as CellValue,
        refreshedFormula
      );
    }

    // A target cell holding real content, or already covered by another
    // spill, blocks this one.
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        if (r === 0 && c === 0) continue; // the anchor itself
        const tr = row + r;
        const tc = col + c;
        const target = this.getCell(sheetId, tr, tc);
        const occupied = !!(target && (target.value != null || target.formula || target.formulaAst));
        if (occupied || spill.isCovered(tr, tc)) {
          return this.storeFormulaResult(sheetId, row, col, cellKey, '#SPILL!', refreshedFormula);
        }
      }
    }

    const region: SpillRegion = { anchorRow: row, anchorCol: col, rows, cols, values: matrix };
    spill.register(region);
    this.pushSpillSeeds(sheetId, region);
    return this.storeFormulaResult(
      sheetId,
      row,
      col,
      cellKey,
      matrix[0][0] as CellValue,
      refreshedFormula
    );
  }

  /**
   * Create a snapshot of the current workbook state
   */
  private createSnapshot(): WorkbookSnapshot {
    const sheets = new Map<string, SheetSnapshot>();

    for (const [sheetId, sheet] of this.sheets.entries()) {
      // Deep clone cells (stable keys are immutable strings; cells get shallow-cloned)
      const cells = new Map<string, Cell>();
      for (const [key, cell] of sheet.cells.entries()) {
        cells.set(key, { ...cell });
      }

      // Snapshot order maps (cloned so later mutations don't leak in)
      const { rowOrder, colOrder } = sheet.snapshotOrderMaps();

      // Deep clone config
      const config = {
        ...sheet.config,
        rowHeights: sheet.config.rowHeights ? new Map(sheet.config.rowHeights) : undefined,
        colWidths: sheet.config.colWidths ? new Map(sheet.config.colWidths) : undefined,
        hiddenRows: sheet.config.hiddenRows ? new Set(sheet.config.hiddenRows) : undefined,
        hiddenCols: sheet.config.hiddenCols ? new Set(sheet.config.hiddenCols) : undefined,
        mergedRegions: sheet.config.mergedRegions
          ? sheet.config.mergedRegions.map((r) => ({ ...r }))
          : undefined,
      };

      sheets.set(sheetId, {
        id: sheet.id,
        name: sheet.name,
        cells,
        rowOrder,
        colOrder,
        config,
        rowCount: sheet.rowCount,
        colCount: sheet.colCount,
      });
    }

    return {
      sheets,
      activeSheetId: this.activeSheetId,
      selection: { ...this.selection },
    };
  }

  /**
   * Restore workbook state from a snapshot
   */
  private restoreSnapshot(snapshot: WorkbookSnapshot): void {
    // Comments are not part of the undo timeline — capture them before the
    // rebuild so an unrelated undo/redo never drops a comment.
    const preservedComments = new Map<string, SheetCommentThread[]>();
    for (const [id, sheet] of this.sheets) {
      preservedComments.set(id, sheet.comments.toJSON());
    }

    // Clear existing sheets
    this.sheets.clear();

    // Restore sheets
    for (const [sheetId, sheetSnapshot] of snapshot.sheets.entries()) {
      const sheet = new SheetImpl(sheetSnapshot.id, sheetSnapshot.name, sheetSnapshot.config);
      sheet.rowCount = sheetSnapshot.rowCount;
      sheet.colCount = sheetSnapshot.colCount;

      // Restore order maps first so cell keys resolve to indices correctly.
      sheet.replaceOrderMaps(sheetSnapshot.rowOrder, sheetSnapshot.colOrder);

      // Restore cells (keys are already stable from the snapshot)
      for (const [key, cell] of sheetSnapshot.cells.entries()) {
        sheet.cells.set(key, { ...cell });
      }

      // Carry comments across the rebuild (see note above).
      const restoredComments = preservedComments.get(sheetId);
      if (restoredComments) sheet.comments.loadJSON(restoredComments);
      this.wireCommentListener(sheet);
      this.attachStructureListenerIfNeeded(sheet);

      this.sheets.set(sheetId, sheet);
    }

    // Restore active sheet
    this.activeSheetId = snapshot.activeSheetId;

    // Restore selection
    this.selection = { ...snapshot.selection };

    // Rebuild formula graph. Prefer existing stable AST; parse + upgrade if absent.
    this.formulaGraph = new FormulaGraphImpl();
    this.spillIndexes.clear(); // spill overlay is derived; re-evaluation rebuilds it
    const resolver = this.buildSheetResolver();
    for (const [sheetId, sheet] of this.sheets.entries()) {
      for (const [row, col, cell] of sheet.entries()) {
        if (!cell.formula && !cell.formulaAst) continue;
        let ast = cell.formulaAst;
        if (!ast && cell.formula) {
          const parsed = this.formulaParser.parse(cell.formula, row, col);
          if (parsed.error) continue;
          ast = toStableAst(parsed.ast, resolver, sheet, row, col);
          cell.formulaAst = ast;
        }
        if (!ast) continue;
        const cellKey = this.graphKey(sheetId, row, col);
        const deps = collectStableDependencies(ast);
        this.formulaGraph.addFormula(cellKey, cell.formula ?? '', deps, formulaIsVolatile(ast));
      }
    }

    // Recalculate all formulas
    for (const [sheetId, sheet] of this.sheets.entries()) {
      for (const [row, col, cell] of sheet.entries()) {
        if (cell.formula) {
          this.evaluateFormula(sheetId, row, col);
        }
      }
    }
    
    // Emit change events
    this.events.emit('workbookChange', {});
    this.events.emit('sheetChange', { sheetId: this.activeSheetId });
    this.events.emit('cellSelection', { selection: this.selection });
  }

  /**
   * Record current state for undo
   */
  recordHistory(): void {
    if (this.isUndoing || this.isRedoing) {
      return; // Don't record history during undo/redo operations
    }
    // Collab mode uses Y.UndoManager (in the collab handle) to track local
    // mutations through Y operations — those broadcast naturally. Skip the
    // snapshot stack entirely so it doesn't double-undo or diverge.
    if (this.collabHandle) return;

    const snapshot = this.createSnapshot();
    this.undoStack.push(snapshot);

    // Limit history size
    if (this.undoStack.length > this.maxHistorySize) {
      this.undoStack.shift();
    }

    // Clear redo stack when new action is performed
    this.redoStack = [];
  }

  /**
   * Undo the last operation
   */
  undo(): boolean {
    if (this.collabHandle) {
      // Route through Y.UndoManager — its inverse op flows through ydoc
      // .on('update') with origin === undoManager, which the binding's
      // dispatcher broadcasts AND reloads the local workbook from Y.
      const stackItem = this.collabHandle.undoManager.undo();
      return stackItem !== null;
    }
    if (this.undoStack.length === 0) {
      return false;
    }

    // Save current state to redo stack
    const currentSnapshot = this.createSnapshot();
    this.redoStack.push(currentSnapshot);

    // Restore previous state
    const previousSnapshot = this.undoStack.pop()!;
    this.isUndoing = true;
    try {
      this.restoreSnapshot(previousSnapshot);
    } finally {
      this.isUndoing = false;
    }

    return true;
  }

  /**
   * Redo the last undone operation
   */
  redo(): boolean {
    if (this.collabHandle) {
      const stackItem = this.collabHandle.undoManager.redo();
      return stackItem !== null;
    }
    if (this.redoStack.length === 0) {
      return false;
    }

    // Save current state to undo stack
    const currentSnapshot = this.createSnapshot();
    this.undoStack.push(currentSnapshot);

    // Restore next state
    const nextSnapshot = this.redoStack.pop()!;
    this.isRedoing = true;
    try {
      this.restoreSnapshot(nextSnapshot);
    } finally {
      this.isRedoing = false;
    }

    return true;
  }

  /**
   * Check if undo is available
   */
  canUndo(): boolean {
    if (this.collabHandle) return this.collabHandle.undoManager.canUndo();
    return this.undoStack.length > 0;
  }

  /**
   * Check if redo is available
   */
  canRedo(): boolean {
    if (this.collabHandle) return this.collabHandle.undoManager.canRedo();
    return this.redoStack.length > 0;
  }

  // ============================================
  // Data Serialization Methods
  // ============================================

  /**
   * Get complete workbook data for serialization
   */
  getData(): import('./types').WorkbookData {
    // Serialize style pool
    const stylePool: Record<string, import('./types').CellStyle> = {};
    for (const [styleId, style] of this.stylePool.getAllStyles()) {
      stylePool[styleId] = style;
    }

    // Serialize format pool
    const formatPool: Record<string, import('./types').CellFormat> = {};
    for (const [formatId, format] of this.formatPool.getAllFormats()) {
      formatPool[formatId] = format;
    }

    // Serialize sheets in the stable-ID wire format (phase 2a.6).
    const sheets: import('./types').SheetData[] = [];
    for (const sheet of this.sheets.values()) {
      // Cells: emit stable keys directly. The accompanying rowOrder/colOrder
      // entries below let consumers reconstruct numeric layout.
      const cells: Array<{ key: string; cell: import('./types').Cell }> = [];
      for (const [row, col, cell] of sheet.entries()) {
        const rowId = sheet.getRowId(row);
        const colId = sheet.getColId(col);
        if (!rowId || !colId) continue; // unreachable: entries() only yields cells with IDs
        cells.push({ key: `${rowId}:${colId}`, cell });
      }

      const orderSnapshot = sheet.snapshotOrderMaps();
      const rowOrder = Array.from(orderSnapshot.rowOrder.entries());
      const colOrder = Array.from(orderSnapshot.colOrder.entries());

      const config = {
        defaultRowHeight: sheet.config.defaultRowHeight,
        defaultColWidth: sheet.config.defaultColWidth,
        // Row/col config entries already use stable IDs internally — pass
        // them through directly.
        rowHeights: sheet.config.rowHeights
          ? Array.from(sheet.config.rowHeights.entries()) as Array<[string, number]>
          : undefined,
        colWidths: sheet.config.colWidths
          ? Array.from(sheet.config.colWidths.entries()) as Array<[string, number]>
          : undefined,
        hiddenRows: sheet.config.hiddenRows
          ? Array.from(sheet.config.hiddenRows) as string[]
          : undefined,
        hiddenCols: sheet.config.hiddenCols
          ? Array.from(sheet.config.hiddenCols) as string[]
          : undefined,
        frozenRows: sheet.config.frozenRows,
        frozenCols: sheet.config.frozenCols,
        showGridLines: sheet.config.showGridLines,
        sortOrder: sheet.config.sortOrder,
        filters: sheet.config.filters
          ? Array.from(sheet.config.filters.entries()) as Array<[string, import('./types').ColumnFilter]>
          : undefined,
        mergedRegions: sheet.config.mergedRegions
          ? sheet.config.mergedRegions.map((r) => ({ ...r }))
          : undefined,
      };

      const threads = sheet.comments.toJSON();

      sheets.push({
        id: sheet.id,
        name: sheet.name,
        cells,
        rowOrder,
        colOrder,
        config,
        rowCount: sheet.rowCount,
        colCount: sheet.colCount,
        threads: threads.length > 0 ? threads : undefined,
      });
    }

    return {
      id: this.id,
      name: this.name,
      activeSheetId: this.activeSheetId,
      defaultRowHeight: this.defaultRowHeight,
      defaultColWidth: this.defaultColWidth,
      stylePool,
      formatPool,
      sheets,
      selection: { ...this.selection },
    };
  }

  /**
   * Set complete workbook data from serialized data
   */
  setData(data: import('./types').WorkbookData): void {
    // Clear existing state
    this.sheets.clear();
    this.formulaGraph = new FormulaGraphImpl();
    this.spillIndexes.clear();
    this.stylePool.clear();
    this.undoStack.length = 0;
    this.redoStack.length = 0;

    // Restore workbook metadata
    this.id = data.id;
    this.name = data.name;
    this.activeSheetId = data.activeSheetId;
    this.defaultRowHeight = data.defaultRowHeight;
    this.defaultColWidth = data.defaultColWidth;

    // Restore selection if provided
    if (data.selection) {
      this.selection = { ...data.selection };
    }

    // Restore the whole style pool in one pass. (Setting entries one at a time
    // via setStyles replaces the entire map each call, collapsing the pool to a
    // single style and leaving every other cell's styleId unresolved.)
    (this.stylePool as StylePool).load(data.stylePool);

    // Restore format pool (if present - for backward compatibility)
    if (data.formatPool) {
      (this.formatPool as FormatPool).load(data.formatPool);
    }

    // Restore sheets. Accept two wire formats:
    //   1. Stable form  — rowOrder/colOrder present; cell keys are stable.
    //      Produced by getData(). Pass-through with no translation.
    //   2. Legacy form  — no rowOrder; cell keys are numeric "r:c". Convenient
    //      for hand- or AI-authored JSON. We materialize stable IDs as we
    //      load.
    for (const sheetData of data.sheets) {
      const useStable = !!(sheetData.rowOrder || sheetData.colOrder);

      // 1. Create the sheet with only the config fields that don't need ID
      //    translation. The rest are filled in below.
      const sheet = new SheetImpl(sheetData.id, sheetData.name, {
        defaultRowHeight: sheetData.config.defaultRowHeight,
        defaultColWidth: sheetData.config.defaultColWidth,
        showGridLines: sheetData.config.showGridLines,
        frozenRows: sheetData.config.frozenRows,
        frozenCols: sheetData.config.frozenCols,
        sortOrder: sheetData.config.sortOrder,
      });
      sheet.rowCount = sheetData.rowCount;
      sheet.colCount = sheetData.colCount;

      // 2. Populate the order maps. In stable form they're explicit; in
      //    legacy form they get materialized below as cells load.
      if (useStable) {
        const rowOrder = new Map(sheetData.rowOrder ?? []);
        const colOrder = new Map(sheetData.colOrder ?? []);
        sheet.replaceOrderMaps(rowOrder, colOrder);
      }

      // 3. Restore cells.
      for (const { key, cell } of sheetData.cells) {
        const cellToStore = { ...cell };
        if ('format' in cellToStore && cellToStore.format && !cellToStore.formatId) {
          const cleanedFormat = this.cleanFormat(cellToStore.format as CellFormat);
          const formatId = this.formatPool.getOrCreate(cleanedFormat);
          cellToStore.formatId = formatId;
          delete (cellToStore as Partial<Cell> & { format?: CellFormat }).format;
        }

        let stableKey: string;
        if (useStable) {
          stableKey = key; // already "rowId:colId"
        } else {
          const { row, col } = parseCellKey(key);
          if (!Number.isFinite(row) || !Number.isFinite(col)) continue;
          stableKey = getStableCellKey(sheet.ensureRowId(row), sheet.ensureColId(col));
        }
        sheet.cells.set(stableKey, cellToStore);
      }

      // 4. Restore ID-keyed config — stable input pass-through; legacy
      //    input gets ensureRowId/ensureColId.
      if (sheetData.config.rowHeights) {
        sheet.config.rowHeights = new Map();
        for (const [key, h] of sheetData.config.rowHeights) {
          const rowId = useStable ? (key as string) : sheet.ensureRowId(key as number);
          sheet.config.rowHeights.set(rowId, h);
        }
      }
      if (sheetData.config.colWidths) {
        sheet.config.colWidths = new Map();
        for (const [key, w] of sheetData.config.colWidths) {
          const colId = useStable ? (key as string) : sheet.ensureColId(key as number);
          sheet.config.colWidths.set(colId, w);
        }
      }
      if (sheetData.config.hiddenRows) {
        sheet.config.hiddenRows = new Set();
        for (const key of sheetData.config.hiddenRows) {
          const rowId = useStable ? (key as string) : sheet.ensureRowId(key as number);
          sheet.config.hiddenRows.add(rowId);
        }
      }
      if (sheetData.config.hiddenCols) {
        sheet.config.hiddenCols = new Set();
        for (const key of sheetData.config.hiddenCols) {
          const colId = useStable ? (key as string) : sheet.ensureColId(key as number);
          sheet.config.hiddenCols.add(colId);
        }
      }
      if (sheetData.config.filters) {
        sheet.config.filters = new Map();
        for (const [key, filter] of sheetData.config.filters) {
          const colId = useStable ? (key as string) : sheet.ensureColId(key as number);
          sheet.config.filters.set(colId, filter);
        }
      }
      // Merged regions are always stored by stable ID — pass through directly.
      if (sheetData.config.mergedRegions) {
        sheet.config.mergedRegions = sheetData.config.mergedRegions.map((r) => ({ ...r }));
      }

      sheet.comments.loadJSON(sheetData.threads);

      this.sheets.set(sheetData.id, sheet);
    }

    // Defensive: activeSheetId must resolve to a real sheet. A collab race
    // (e.g. a sheet-delete update arriving before the meta update) could
    // otherwise leave it dangling and crash getSheet().
    if (!this.sheets.has(this.activeSheetId)) {
      const first = this.sheets.keys().next().value;
      if (first) this.activeSheetId = first;
    }

    // Re-attach structure listeners. setData replaces every SheetImpl instance,
    // so the listeners are gone — without this, a structural edit (insert/delete
    // row/col) after a load stops recalculating dependent formulas, and (in
    // collab) a peer's structural mutations stop mirroring.
    for (const sheet of this.sheets.values()) {
      this.attachStructureListenerIfNeeded(sheet as SheetImpl);
    }

    // Wire comment-change listeners — setData replaces every SheetImpl, so
    // any listeners installed earlier are gone. Always wired (not collab-
    // gated): they emit `commentChange` for the UI and mirror to the Y.Doc
    // only when collab is attached.
    for (const sheet of this.sheets.values()) {
      this.wireCommentListener(sheet as SheetImpl);
    }

    // Rebuild formula graph and evaluate formulas
    this.rebuildFormulaGraph();

    // Emit change events
    this.events.emit('workbookChange', {});
    this.events.emit('sheetChange', { sheetId: this.activeSheetId });
    this.events.emit('cellSelection', { selection: this.selection });
  }

  /**
   * Clean up format object to only include properties relevant to the format type
   */
  private cleanFormat(format: CellFormat): CellFormat {
    if (!format.type) {
      return format;
    }

    const cleaned: CellFormat = { type: format.type };

    switch (format.type) {
      case 'number':
        cleaned.decimalPlaces = format.decimalPlaces;
        cleaned.useThousandsSeparator = format.useThousandsSeparator;
        cleaned.negativeFormat = format.negativeFormat;
        break;

      case 'currency':
        cleaned.decimalPlaces = format.decimalPlaces;
        cleaned.currencyCode = format.currencyCode;
        cleaned.currencySymbolPosition = format.currencySymbolPosition;
        cleaned.negativeFormat = format.negativeFormat;
        break;

      case 'accounting':
        cleaned.decimalPlaces = format.decimalPlaces;
        cleaned.currencyCode = format.currencyCode;
        cleaned.negativeFormat = format.negativeFormat;
        break;

      case 'percentage':
        cleaned.decimalPlaces = format.decimalPlaces;
        break;

      case 'scientific':
        cleaned.decimalPlaces = format.decimalPlaces;
        break;

      case 'fraction':
        cleaned.fractionType = format.fractionType;
        break;

      case 'date':
        cleaned.dateFormat = format.dateFormat;
        break;

      case 'time':
        cleaned.timeFormat = format.timeFormat;
        break;

      case 'datetime':
        cleaned.dateFormat = format.dateFormat;
        cleaned.timeFormat = format.timeFormat;
        break;

      case 'duration':
        cleaned.durationFormat = format.durationFormat;
        break;

      case 'custom':
        cleaned.pattern = format.pattern;
        break;

      case 'text':
      default:
        // Text format doesn't need additional properties
        break;
    }

    return cleaned;
  }

  /**
   * Rebuild formula graph from current cells and evaluate all formulas
   */
  private rebuildFormulaGraph(): void {
    const resolver = this.buildSheetResolver();
    for (const [sheetId, sheet] of this.sheets.entries()) {
      for (const [row, col, cell] of sheet.entries()) {
        if (!cell.formula && !cell.formulaAst) continue;
        let ast = cell.formulaAst;
        if (!ast && cell.formula) {
          const parsed = this.formulaParser.parse(cell.formula, row, col);
          if (parsed.error) continue;
          ast = toStableAst(parsed.ast, resolver, sheet, row, col);
          cell.formulaAst = ast;
        }
        if (!ast) continue;
        const cellKey = this.graphKey(sheetId, row, col);
        const deps = collectStableDependencies(ast);
        this.formulaGraph.addFormula(cellKey, cell.formula ?? '', deps, formulaIsVolatile(ast));
        this.evaluateFormula(sheetId, row, col);
      }
    }
  }

  // ============================================
  // Sorting Methods
  // ============================================

  /**
   * Set the sort order for a sheet
   * @param sortOrder Sort order to apply
   * @param sheetId Target sheet ID (defaults to active sheet)
   */
  setSortOrder(sortOrder: SortOrder[], sheetId?: string): void {
    const targetSheetId = sheetId ?? this.activeSheetId;
    const sheet = this.sheets.get(targetSheetId);
    if (sheet) {
      // Record history before making changes
      if (!this.isUndoing && !this.isRedoing && !this.isBatching) {
        this.recordHistory();
      }
      sheet.setSortOrder(sortOrder);
    }
  }

  /**
   * Get the sort order for a sheet
   * @param sheetId Target sheet ID (defaults to active sheet)
   * @returns Current sort order
   */
  getSortOrder(sheetId?: string): SortOrder[] {
    const targetSheetId = sheetId ?? this.activeSheetId;
    const sheet = this.sheets.get(targetSheetId);
    return sheet ? sheet.getSortOrder() : [];
  }

  /**
   * Clear sorting for a sheet
   * @param sheetId Target sheet ID (defaults to active sheet)
   */
  clearSort(sheetId?: string): void {
    const targetSheetId = sheetId ?? this.activeSheetId;
    const sheet = this.sheets.get(targetSheetId);
    if (sheet) {
      // Record history before making changes
      if (!this.isUndoing && !this.isRedoing && !this.isBatching) {
        this.recordHistory();
      }
      sheet.clearSort();
    }
  }

  /**
   * Sort the sheet data according to current sort order
   * @param sheetId Target sheet ID (defaults to active sheet)
   */
  sortSheet(sheetId?: string): void {
    const targetSheetId = sheetId ?? this.activeSheetId;
    const sheet = this.sheets.get(targetSheetId);
    if (sheet) {
      // Sorting permutes rows; a merged region's corners would scatter.
      // Refuse the sort while any merged region exists (matches Excel/Sheets).
      if (sheet.getMergedRegions().length > 0) return;
      const sortOrder = sheet.getSortOrder();
      if (sortOrder.length > 0) {
        // Record history before sorting
        if (!this.isUndoing && !this.isRedoing && !this.isBatching) {
          this.recordHistory();
        }

        SortManager.sortRows(sheet, sortOrder);
        // Sorting reorders the data rows, so any formula that references into the
        // sorted range must recompute (SortManager fires notifyStructureChange
        // with affectsFormulas=false — cosmetic-vs-reindex is ambiguous for a
        // permute, so recalc here explicitly, matching delete/rename-sheet).
        if (!this.isApplyingRemoteChange) this.recalculateAllFormulas();
      }
    }
  }

  // ============================================
  // Filtering Methods
  // ============================================

  /**
   * Set a filter for a column
   * @param column Column index
   * @param filter Filter configuration
   * @param sheetId Target sheet ID (defaults to active sheet)
   */
  setFilter(column: number, filter: import('./types').ColumnFilter, sheetId?: string): void {
    const targetSheetId = sheetId ?? this.activeSheetId;
    const sheet = this.sheets.get(targetSheetId);
    if (sheet) {
      // Record history before making changes
      if (!this.isUndoing && !this.isRedoing && !this.isBatching) {
        this.recordHistory();
      }
      sheet.setFilter(column, filter);
    }
  }

  /**
   * Clear filter for a specific column
   * @param column Column index
   * @param sheetId Target sheet ID (defaults to active sheet)
   */
  clearFilter(column: number, sheetId?: string): void {
    const targetSheetId = sheetId ?? this.activeSheetId;
    const sheet = this.sheets.get(targetSheetId);
    if (sheet) {
      // Record history before making changes
      if (!this.isUndoing && !this.isRedoing && !this.isBatching) {
        this.recordHistory();
      }
      sheet.clearFilter(column);
    }
  }

  /**
   * Get all active filters for a sheet
   * @param sheetId Target sheet ID (defaults to active sheet)
   * @returns Map of column -> filter
   */
  getFilters(sheetId?: string): Map<number, import('./types').ColumnFilter> {
    const targetSheetId = sheetId ?? this.activeSheetId;
    const sheet = this.sheets.get(targetSheetId);
    return sheet ? sheet.getFilters() : new Map();
  }

  /**
   * Clear all filters for a sheet
   * @param sheetId Target sheet ID (defaults to active sheet)
   */
  clearAllFilters(sheetId?: string): void {
    const targetSheetId = sheetId ?? this.activeSheetId;
    const sheet = this.sheets.get(targetSheetId);
    if (sheet) {
      // Record history before making changes
      if (!this.isUndoing && !this.isRedoing && !this.isBatching) {
        this.recordHistory();
      }
      sheet.clearAllFilters();
    }
  }

  // ============================================
  // Merged cells
  // ============================================

  /** Merge a range of cells into a single region (records history). */
  mergeCells(range: Range, sheetId?: string): void {
    const sheet = this.sheets.get(sheetId ?? this.activeSheetId);
    if (sheet) {
      if (!this.isUndoing && !this.isRedoing && !this.isBatching) {
        this.recordHistory();
      }
      sheet.mergeCells(range);
    }
  }

  /** Remove any merged regions intersecting the range (records history). */
  unmergeCells(range: Range, sheetId?: string): void {
    const sheet = this.sheets.get(sheetId ?? this.activeSheetId);
    if (sheet) {
      if (!this.isUndoing && !this.isRedoing && !this.isBatching) {
        this.recordHistory();
      }
      sheet.unmergeCells(range);
    }
  }

  // ============================================
  // Collaboration
  // ============================================

  /**
   * Attach a CollabProvider so this workbook's cells, structure, and styles
   * sync with peers via Yjs. The returned handle exposes the Y.Doc and
   * Awareness; the React layer reads it via getCollabHandle() to overlay
   * remote cell selections.
   *
   * Idempotent guard: throws if already attached.
   */
  async attachCollab(
    provider: import('@b32nio/spindle-shared').CollabProvider,
    identity: import('@b32nio/spindle-shared').CollabIdentity,
    options?: AttachCollabOptions,
  ): Promise<CollabHandle> {
    if (this.collabHandle) {
      throw new Error('Collaboration is already attached to this workbook.');
    }
    const handle = await attachCollabToWorkbook(
      this.getData(),
      this,
      provider,
      identity,
      options,
    );
    this.collabHandle = handle;
    // Wire structure-change listeners on every existing sheet so structural
    // mutations (insert rows/cols, height/width, hide, freeze, filter, sort)
    // flow into the Y.Doc. Sheets added later go through addSheet which calls
    // attachStructureListenerIfNeeded.
    for (const sheet of this.sheets.values()) {
      this.attachStructureListenerIfNeeded(sheet as SheetImpl);
    }
    this.events.emit('workbookChange', { action: 'attachCollab' });
    return handle;
  }

  /** Return the live collab handle, or null if not in collab mode. */
  getCollabHandle(): CollabHandle | null {
    return this.collabHandle;
  }

  /** Detach from the current collaboration session. */
  detachCollab(): void {
    if (!this.collabHandle) return;
    for (const sheet of this.sheets.values()) {
      (sheet as SheetImpl).__setStructureChangeListener(undefined);
    }
    this.collabHandle.detach();
    this.collabHandle = null;
    this.events.emit('workbookChange', { action: 'detachCollab' });
  }

  /**
   * Install our mirror handler as `sheet`'s structure-change listener iff
   * collab is currently attached. Called from attachCollab (for existing
   * sheets) and addSheet (for new ones).
   */
  private attachStructureListenerIfNeeded(sheet: SheetImpl): void {
    sheet.__setStructureChangeListener((affectsFormulas) =>
      this.onStructureChange(sheet.id, affectsFormulas),
    );
  }

  /**
   * React to a structural edit on a sheet. Insert/delete row/col shift what
   * formulas reference (`affectsFormulas`), so every formula must recompute — the
   * stable-AST refs rebase correctly but nothing recomputes the cached values
   * otherwise, leaving stale totals. Cosmetic changes (column width, row height,
   * hide, freeze) pass `affectsFormulas: false` and must NOT recalc — otherwise a
   * resize drag recomputes the whole workbook on every mousemove. Always mirror
   * the structure to collab when attached. All skipped during a remote apply (the
   * remote already carries recomputed values, and mirroring would echo it back).
   */
  private onStructureChange(sheetId: string, affectsFormulas: boolean): void {
    if (this.isApplyingRemoteChange) return;
    if (affectsFormulas) this.recalculateAllFormulas();
    if (this.collabHandle) {
      this.mirrorSheetStructure(sheetId);
    }
  }

  /**
   * @internal Used by the collab binding to apply a remote-driven
   * WorkbookData reload without tripping any guard rails that would
   * otherwise reject setData while collab is attached. Behaves like
   * setData but signals collab origin so consumers can disambiguate.
   */
  _reloadFromCollab(data: import('./types').WorkbookData): void {
    // Suspend history recording AND collab mirroring during the reload —
    // remote edits should not collide with local undo/redo stacks, and
    // mirror calls would echo every remote write back over the wire.
    const wasUndoing = this.isUndoing;
    this.isUndoing = true;
    this.isApplyingRemoteChange = true;
    try {
      this.setData(data);
    } finally {
      this.isUndoing = wasUndoing;
      this.isApplyingRemoteChange = false;
    }
  }

  // ============================================
  // Collab mirroring helpers (write local mutations into the Y.Doc)
  // ============================================

  /**
   * Echo a cell write at (row, col) into the workbook's Y.Doc so peers
   * receive the change. No-op when collab is detached or we're currently
   * applying a remote update.
   *
   * Mirrors the cell value + the rowId/colId positions, so a brand-new
   * row or column (not present at hydration time) lands on the peer side
   * with its display index intact.
   */
  private mirrorCellWrite(sheetId: string, row: number, col: number): void {
    if (!this.collabHandle || this.isApplyingRemoteChange) return;
    const sheet = this.sheets.get(sheetId) as SheetImpl | undefined;
    if (!sheet) return;
    const cell = sheet.getCell(row, col);
    if (!cell) return;

    const rowId = sheet.ensureRowId(row);
    const colId = sheet.ensureColId(col);
    const stableKey = getStableCellKey(rowId, colId);

    const ydoc = this.collabHandle.ydoc;
    const yTypes = getWorkbookYTypes(ydoc);
    const ySheetMap = yTypes.sheets.get(sheetId);
    if (!ySheetMap) return;
    const t = getSheetYTypes(ySheetMap);
    if (!t.cells || !t.rowOrder || !t.colOrder) return;

    ydoc.transact(() => {
      // Ensure rowOrder/colOrder entries exist for these IDs at the right index.
      if (t.rowOrder.get(rowId) !== row) t.rowOrder.set(rowId, row);
      if (t.colOrder.get(colId) !== col) t.colOrder.set(colId, col);

      t.cells.set(stableKey, { ...cell });

      // Pool sync — if the cell references a style/format we haven't
      // mirrored yet, push it now so peers don't render with stale ids.
      if (cell.styleId) {
        const style = this.stylePool.get(cell.styleId);
        if (style && !yTypes.stylePool.has(cell.styleId)) {
          yTypes.stylePool.set(cell.styleId, { ...style });
        }
      }
      if (cell.formatId) {
        const format = this.formatPool.get(cell.formatId);
        if (format && !yTypes.formatPool.has(cell.formatId)) {
          yTypes.formatPool.set(cell.formatId, { ...format });
        }
      }
    });
  }

  /** Echo a new sheet into the Y.Doc. */
  private mirrorSheetAdd(sheetId: string, name: string, sheet: SheetImpl): void {
    if (!this.collabHandle || this.isApplyingRemoteChange) return;
    const ydoc = this.collabHandle.ydoc;
    const yTypes = getWorkbookYTypes(ydoc);
    ydoc.transact(() => {
      const ySheetMap = ensureSheetYMap(yTypes.sheets, sheetId);
      // Seed minimal meta + counts (now safe; sheetMap is integrated).
      const meta = ySheetMap.get('meta') as Y.Map<unknown>;
      meta.set('id', sheetId);
      meta.set('name', name);
      meta.set('rowCount', sheet.rowCount);
      meta.set('colCount', sheet.colCount);
      // Append to sheetIds order if not already present
      const ids = yTypes.sheetIds.toArray();
      if (!ids.includes(sheetId)) {
        yTypes.sheetIds.push([sheetId]);
      }
    });
  }

  /**
   * Echo a sheet deletion. activeSheetId is NOT touched — it's per-user
   * view state; a peer that was viewing the deleted sheet snaps to the
   * first sheet via the defensive fallback in setData.
   */
  private mirrorSheetDelete(sheetId: string): void {
    if (!this.collabHandle || this.isApplyingRemoteChange) return;
    const ydoc = this.collabHandle.ydoc;
    const yTypes = getWorkbookYTypes(ydoc);
    ydoc.transact(() => {
      yTypes.sheets.delete(sheetId);
      // Remove from sheetIds order
      const ids = yTypes.sheetIds.toArray();
      const idx = ids.indexOf(sheetId);
      if (idx !== -1) yTypes.sheetIds.delete(idx, 1);
      // Purge the sheet's comment threads.
      const prefix = `${sheetId}/`;
      for (const key of Array.from(yTypes.threads.keys())) {
        if (key.startsWith(prefix)) yTypes.threads.delete(key);
      }
    });
  }

  /** Echo a sheet rename. */
  private mirrorSheetRename(sheetId: string, newName: string): void {
    if (!this.collabHandle || this.isApplyingRemoteChange) return;
    const ydoc = this.collabHandle.ydoc;
    const yTypes = getWorkbookYTypes(ydoc);
    const ySheetMap = yTypes.sheets.get(sheetId);
    if (!ySheetMap) return;
    const meta = ySheetMap.get('meta') as Y.Map<unknown> | undefined;
    if (!meta) return;
    ydoc.transact(() => {
      meta.set('name', newName);
    });
  }


  /**
   * Re-sync a sheet's full structure to the Y.Doc: order maps, dimension
   * maps (heights/widths), hidden sets, filters, sort order, frozen counts,
   * row/col counts. Cells are NOT touched here — they flow via
   * mirrorCellWrite per setCell. Called by the SheetImpl structure-change
   * listener after any structural mutation.
   */
  private mirrorSheetStructure(sheetId: string): void {
    if (!this.collabHandle || this.isApplyingRemoteChange) return;
    const sheet = this.sheets.get(sheetId) as SheetImpl | undefined;
    if (!sheet) return;
    const ydoc = this.collabHandle.ydoc;
    const yTypes = getWorkbookYTypes(ydoc);
    const ySheetMap = yTypes.sheets.get(sheetId);
    if (!ySheetMap) return;
    const t = getSheetYTypes(ySheetMap);

    const orderSnapshot = sheet.snapshotOrderMaps();

    ydoc.transact(() => {
      // rowOrder: Y.Map<rowId, idx>. Rebuild wholesale (clear + repopulate).
      for (const k of Array.from(t.rowOrder.keys())) t.rowOrder.delete(k);
      for (const [idx, id] of orderSnapshot.rowOrder) t.rowOrder.set(id, idx);

      for (const k of Array.from(t.colOrder.keys())) t.colOrder.delete(k);
      for (const [idx, id] of orderSnapshot.colOrder) t.colOrder.set(id, idx);

      // Drop cells whose row/col id was deleted from the order maps.
      const validRowIds = new Set(orderSnapshot.rowOrder.values());
      const validColIds = new Set(orderSnapshot.colOrder.values());
      const orphans: string[] = [];
      for (const key of t.cells.keys()) {
        const sep = key.indexOf(':');
        const rowId = key.slice(0, sep);
        const colId = key.slice(sep + 1);
        if (!validRowIds.has(rowId) || !validColIds.has(colId)) orphans.push(key);
      }
      for (const k of orphans) t.cells.delete(k);

      // Dimension maps + hidden sets + filters: rebuild from sheet config.
      syncMapToY(t.rowHeights, sheet.config.rowHeights);
      syncMapToY(t.colWidths, sheet.config.colWidths);
      syncSetToYMap(t.hiddenRows, sheet.config.hiddenRows);
      syncSetToYMap(t.hiddenCols, sheet.config.hiddenCols);
      syncMapToY(t.filters, sheet.config.filters);

      // Meta updates.
      t.meta.set('rowCount', sheet.rowCount);
      t.meta.set('colCount', sheet.colCount);
      setOrDelete(t.meta, 'frozenRows', sheet.config.frozenRows);
      setOrDelete(t.meta, 'frozenCols', sheet.config.frozenCols);
      setOrDelete(t.meta, 'showGridLines', sheet.config.showGridLines);
      setOrDelete(t.meta, 'sortOrder', sheet.config.sortOrder);
      // Merged regions: keyed by composite stable-ID key so concurrent
      // merges of different ranges by different peers both survive.
      const mergedByKey = new Map<string, MergedRegion>();
      if (sheet.config.mergedRegions) {
        for (const r of sheet.config.mergedRegions) {
          mergedByKey.set(mergedRegionKey(r), r);
        }
      }
      syncMapToY(t.mergedRegions, mergedByKey);
    });
  }

  // ============================================
  // Comment threads
  // ============================================

  /**
   * Wire a sheet's comment store so every mutation emits a `commentChange`
   * event for the UI and mirrors the threads into the Y.Doc.
   */
  private wireCommentListener(sheet: SheetImpl): void {
    sheet.comments.__setChangeListener((event) => this.onCommentChange(sheet.id, event));
  }

  /**
   * Comment-mutation handler: mirror to collab (if attached), trigger a UI
   * re-render (`commentChange`), and surface a semantic `commentEvent` the
   * host app can hook for notifications.
   */
  private onCommentChange(sheetId: string, event: CommentMutationEvent): void {
    this.mirrorSheetThreads(sheetId);
    this.events.emit('commentChange', { sheetId });
    this.events.emit('commentEvent', { ...event, sheetId });
  }

  /**
   * Re-sync a sheet's comment threads into the Y.Doc. Threads live in a flat
   * top-level `threads` map keyed "sheetId/threadId"; drop the sheet's
   * existing entries and repopulate from the store. No-op when collab is
   * detached or a remote update is being applied.
   */
  private mirrorSheetThreads(sheetId: string): void {
    if (!this.collabHandle || this.isApplyingRemoteChange) return;
    const sheet = this.sheets.get(sheetId) as SheetImpl | undefined;
    if (!sheet) return;
    const ydoc = this.collabHandle.ydoc;
    const yTypes = getWorkbookYTypes(ydoc);
    const ySheetMap = yTypes.sheets.get(sheetId);
    const prefix = `${sheetId}/`;
    ydoc.transact(() => {
      for (const key of Array.from(yTypes.threads.keys())) {
        if (key.startsWith(prefix)) yTypes.threads.delete(key);
      }
      const sheetT = ySheetMap ? getSheetYTypes(ySheetMap) : undefined;
      for (const thread of sheet.comments.toJSON()) {
        yTypes.threads.set(threadKey(sheetId, thread.id), thread);
        // Ensure the anchored cell's row/col positions reach peers, so the
        // anchor still resolves when the commented cell is otherwise empty.
        if (sheetT) {
          const r = sheet.getRowIndex(thread.anchor.rowId);
          const c = sheet.getColIndex(thread.anchor.colId);
          if (r !== undefined) sheetT.rowOrder.set(thread.anchor.rowId, r);
          if (c !== undefined) sheetT.colOrder.set(thread.anchor.colId, c);
        }
      }
    });
  }
}

function syncMapToY<V>(yMap: Y.Map<V>, sheetMap: Map<string, V> | undefined): void {
  for (const k of Array.from(yMap.keys())) yMap.delete(k);
  if (!sheetMap) return;
  for (const [k, v] of sheetMap) yMap.set(k, v);
}

function syncSetToYMap(yMap: Y.Map<true>, sheetSet: Set<string> | undefined): void {
  for (const k of Array.from(yMap.keys())) yMap.delete(k);
  if (!sheetSet) return;
  for (const k of sheetSet) yMap.set(k, true);
}

function setOrDelete(yMap: Y.Map<unknown>, key: string, value: unknown): void {
  if (value === undefined) yMap.delete(key);
  else yMap.set(key, value);
}

