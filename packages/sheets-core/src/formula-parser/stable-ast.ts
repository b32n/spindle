// Stable-ID formula AST.
//
// The parser produces a ParsedFormulaNode whose CellReferences hold numeric
// (row, col) coords (absolute for $-marked refs, offset for relatives).
// That representation is fragile under structural changes: inserting a row
// shifts the cell-at-row-N but not the formulas that pointed there.
//
// A StableFormulaNode replaces each numeric ref with a (rowId, colId) pair.
// Inserts/deletes/sorts mutate only the order maps; the AST's IDs keep
// pointing at the same logical cells. Display text is regenerated from the
// AST on every evaluation so the user sees the up-to-date A1 spelling.

import type { Sheet, FormulaDependencies, RangeDependency } from '../types';
import { columnIndexToLabel } from './cell-reference';
import { volatileFunctions } from './functions';
import type { CellReference, ParsedFormulaNode } from './types';

export interface StableCellRef {
  rowId: string;
  colId: string;
  rowAbsolute: boolean;
  colAbsolute: boolean;
  sheetName?: string;
  // A cross-sheet ref whose target sheet did NOT exist when the formula was
  // entered can't capture that sheet's stable ids. Rather than binding to the
  // current sheet (which caused a false #CIRCULAR! and never recovered), it
  // stores its absolute numeric position here (rowId/colId are set to the #REF!
  // sentinel) and resolves by position once the named sheet appears.
  unresolvedRow?: number;
  unresolvedCol?: number;
}

export interface StableRangeRef {
  start: StableCellRef;
  end: StableCellRef;
}

export interface StableFormulaNode {
  type: ParsedFormulaNode['type'];
  value?: unknown;
  cellRef?: StableCellRef;
  rangeRef?: StableRangeRef;
  functionName?: string;
  args?: StableFormulaNode[];
  operator?: string;
  left?: StableFormulaNode;
  right?: StableFormulaNode;
}

/** Thrown when fromStableAst encounters a ref whose target IDs are gone. */
export class StableRefDeletedError extends Error {
  constructor() {
    super('#REF!');
    this.name = 'StableRefDeletedError';
  }
}

// A row/col id that resolves to nothing (real ids are 12-char base32, so '#'
// can never collide). Assigned to a rebased ref that fell off the top/left of
// the grid, so it surfaces as #REF! instead of a silent clamp to the edge.
const INVALID_REF_ID = '#REF!';

export interface SheetResolver {
  getSheet(name?: string): Sheet | undefined;
}

// ============================================================================
// Numeric AST → Stable AST  (called on setFormula)
// ============================================================================

export function toStableAst(
  node: ParsedFormulaNode,
  resolver: SheetResolver,
  currentSheet: Sheet,
  currentRow: number,
  currentCol: number,
): StableFormulaNode {
  const out: StableFormulaNode = { type: node.type };
  if (node.value !== undefined) out.value = node.value;
  if (node.functionName !== undefined) out.functionName = node.functionName;
  if (node.operator !== undefined) out.operator = node.operator;
  if (node.cellRef) {
    out.cellRef = cellRefToStable(node.cellRef, resolver, currentSheet, currentRow, currentCol);
  }
  if (node.rangeRef) {
    out.rangeRef = {
      start: cellRefToStable(node.rangeRef.start, resolver, currentSheet, currentRow, currentCol),
      end: cellRefToStable(node.rangeRef.end, resolver, currentSheet, currentRow, currentCol),
    };
  }
  if (node.args) {
    out.args = node.args.map((a) => toStableAst(a, resolver, currentSheet, currentRow, currentCol));
  }
  if (node.left) {
    out.left = toStableAst(node.left, resolver, currentSheet, currentRow, currentCol);
  }
  if (node.right) {
    out.right = toStableAst(node.right, resolver, currentSheet, currentRow, currentCol);
  }
  return out;
}

function cellRefToStable(
  ref: CellReference,
  resolver: SheetResolver,
  currentSheet: Sheet,
  currentRow: number,
  currentCol: number,
): StableCellRef {
  // Parser convention: absolute refs hold absolute positions; relatives hold offsets.
  const absRow = ref.rowAbsolute ? ref.row : currentRow + ref.row;
  const absCol = ref.colAbsolute ? ref.col : currentCol + ref.col;
  const target = ref.sheetName ? resolver.getSheet(ref.sheetName) : currentSheet;
  if (!target) {
    // Named sheet doesn't exist yet — defer. Don't bind to currentSheet (that
    // collides with the formula's own cell → false #CIRCULAR!, and never
    // recovers). Keep the position; resolve when the sheet appears.
    return {
      rowId: INVALID_REF_ID,
      colId: INVALID_REF_ID,
      rowAbsolute: ref.rowAbsolute,
      colAbsolute: ref.colAbsolute,
      sheetName: ref.sheetName,
      unresolvedRow: absRow,
      unresolvedCol: absCol,
    };
  }
  return {
    rowId: target.ensureRowId(absRow),
    colId: target.ensureColId(absCol),
    rowAbsolute: ref.rowAbsolute,
    colAbsolute: ref.colAbsolute,
    sheetName: ref.sheetName,
  };
}

// ============================================================================
// Stable AST → Numeric AST  (called on every evaluate)
// ============================================================================

export function fromStableAst(
  node: StableFormulaNode,
  resolver: SheetResolver,
  currentSheet: Sheet,
  currentRow: number,
  currentCol: number,
): ParsedFormulaNode {
  const out: ParsedFormulaNode = { type: node.type };
  if (node.value !== undefined) out.value = node.value;
  if (node.functionName !== undefined) out.functionName = node.functionName;
  if (node.operator !== undefined) out.operator = node.operator;
  if (node.cellRef) {
    out.cellRef = cellRefFromStable(node.cellRef, resolver, currentSheet, currentRow, currentCol);
  }
  if (node.rangeRef) {
    out.rangeRef = {
      start: cellRefFromStable(node.rangeRef.start, resolver, currentSheet, currentRow, currentCol),
      end: cellRefFromStable(node.rangeRef.end, resolver, currentSheet, currentRow, currentCol),
    };
  }
  if (node.args) {
    out.args = node.args.map((a) => fromStableAst(a, resolver, currentSheet, currentRow, currentCol));
  }
  if (node.left) out.left = fromStableAst(node.left, resolver, currentSheet, currentRow, currentCol);
  if (node.right) out.right = fromStableAst(node.right, resolver, currentSheet, currentRow, currentCol);
  return out;
}

function cellRefFromStable(
  ref: StableCellRef,
  resolver: SheetResolver,
  currentSheet: Sheet,
  currentRow: number,
  currentCol: number,
): CellReference {
  // A named sheet that no longer resolves is #REF! — never silently fall back to
  // the current sheet (that returns neighboring local data / a false circular).
  let sheet: Sheet;
  if (ref.sheetName) {
    const resolved = resolver.getSheet(ref.sheetName);
    if (!resolved) throw new StableRefDeletedError();
    sheet = resolved;
  } else {
    sheet = currentSheet;
  }

  // A deferred cross-sheet ref (target sheet absent at entry time) resolves by
  // its stored absolute position now that the sheet exists.
  if (ref.unresolvedRow !== undefined && ref.unresolvedCol !== undefined) {
    const uRow = ref.unresolvedRow;
    const uCol = ref.unresolvedCol;
    return {
      row: ref.rowAbsolute ? uRow : uRow - currentRow,
      col: ref.colAbsolute ? uCol : uCol - currentCol,
      rowAbsolute: ref.rowAbsolute,
      colAbsolute: ref.colAbsolute,
      sheetName: ref.sheetName,
    };
  }

  const absRow = sheet.getRowIndex(ref.rowId);
  const absCol = sheet.getColIndex(ref.colId);
  if (absRow === undefined || absCol === undefined) {
    throw new StableRefDeletedError();
  }
  // Re-encode in the parser's offset convention so the existing evaluator
  // consumes it unchanged.
  return {
    row: ref.rowAbsolute ? absRow : absRow - currentRow,
    col: ref.colAbsolute ? absCol : absCol - currentCol,
    rowAbsolute: ref.rowAbsolute,
    colAbsolute: ref.colAbsolute,
    sheetName: ref.sheetName,
  };
}

// ============================================================================
// Stable AST → Display string  (regenerated on every evaluate)
// ============================================================================

export function renderStableAst(
  node: StableFormulaNode,
  resolver: SheetResolver,
  currentSheet: Sheet,
): string {
  return '=' + renderNode(node, resolver, currentSheet);
}

function renderNode(
  node: StableFormulaNode,
  resolver: SheetResolver,
  currentSheet: Sheet,
): string {
  switch (node.type) {
    case 'number':
      return String(node.value);
    case 'string':
      return typeof node.value === 'string' ? `"${node.value}"` : String(node.value ?? '');
    case 'variable':
      return String(node.value ?? '');
    case 'cell':
      return node.cellRef ? renderRef(node.cellRef, resolver, currentSheet) : '#REF!';
    case 'range':
      if (!node.rangeRef) return '#REF!';
      return `${renderRef(node.rangeRef.start, resolver, currentSheet)}:${renderRef(node.rangeRef.end, resolver, currentSheet)}`;
    case 'function':
      return `${node.functionName ?? ''}(${(node.args ?? [])
        .map((a) => renderNode(a, resolver, currentSheet))
        .join(',')})`;
    case 'operator': {
      const op = node.operator ?? '';
      const render = (n: StableFormulaNode) => renderNode(n, resolver, currentSheet);
      // A leading sign is stored as `"" op x`; it binds tighter than any binary
      // operator, so a binary operand needs its parentheses back.
      if ((op === '-' || op === '+') && isEmptyLiteral(node.left) && node.right) {
        const r = render(node.right);
        return `${op}${node.right.type === 'operator' && !isUnarySign(node.right) ? `(${r})` : r}`;
      }
      const prec = OPERATOR_PRECEDENCE[op] ?? 0;
      const l = node.left ? render(node.left) : '';
      const r = node.right ? render(node.right) : '';
      // Parentheses are not kept in the AST; restore the ones precedence needs.
      // All binary operators are left-associative, so an equal-precedence right
      // operand (a-(b-c)) needs them too.
      const wrapLeft = node.left?.type === 'operator' && !isUnarySign(node.left) && (OPERATOR_PRECEDENCE[node.left.operator ?? ''] ?? 0) < prec;
      const wrapRight = node.right?.type === 'operator' && !isUnarySign(node.right) && (OPERATOR_PRECEDENCE[node.right.operator ?? ''] ?? 0) <= prec;
      return `${wrapLeft ? `(${l})` : l}${op}${wrapRight ? `(${r})` : r}`;
    }
    default:
      return '';
  }
}

// Same order as the parser: comparison < & < +,- < *,/ < ^.
const OPERATOR_PRECEDENCE: Record<string, number> = {
  '=': 0, '<>': 0, '<': 0, '>': 0, '<=': 0, '>=': 0,
  '&': 1, '+': 2, '-': 2, '*': 3, '/': 3, '^': 4,
};

function isEmptyLiteral(node: StableFormulaNode | undefined): boolean {
  return !!node && node.type === 'string' && node.value === '';
}

/** `"" - x` / `"" + x`: the parser's form of a unary sign. */
function isUnarySign(node: StableFormulaNode): boolean {
  return node.type === 'operator' && (node.operator === '-' || node.operator === '+') && isEmptyLiteral(node.left);
}

function renderRef(
  ref: StableCellRef,
  resolver: SheetResolver,
  currentSheet: Sheet,
): string {
  let row: number | undefined;
  let col: number | undefined;
  if (ref.unresolvedRow !== undefined && ref.unresolvedCol !== undefined) {
    // Deferred cross-sheet ref: render from its stored position so the formula
    // bar keeps showing e.g. Ghost!A1 rather than #REF! while the sheet is absent.
    row = ref.unresolvedRow;
    col = ref.unresolvedCol;
  } else {
    const sheet = ref.sheetName ? (resolver.getSheet(ref.sheetName) ?? currentSheet) : currentSheet;
    row = sheet.getRowIndex(ref.rowId);
    col = sheet.getColIndex(ref.colId);
  }
  if (row === undefined || col === undefined) return '#REF!';
  const colLabel = columnIndexToLabel(col);
  const cell =
    (ref.colAbsolute ? '$' : '') + colLabel + (ref.rowAbsolute ? '$' : '') + (row + 1);
  if (ref.sheetName) {
    return /\s/.test(ref.sheetName) ? `'${ref.sheetName}'!${cell}` : `${ref.sheetName}!${cell}`;
  }
  return cell;
}

// ============================================================================
// Dependency extraction (returns stable graph keys "rowId:colId")
// ============================================================================

// ============================================================================
// Copy/paste rebasing (used by fill, copy)
// ============================================================================

/**
 * Produce a new stable AST suitable for pasting a formula that originated at
 * (sourceRow, sourceCol) into a cell at (targetRow, targetCol).
 *
 * Absolute refs ($A$1) are preserved unchanged — they continue to identify
 * the same cell. Relative refs are rebased by the source→target offset: if
 * the original ref pointed at the cell at (srcRow + N), the new ref points
 * at the cell at (tgtRow + N). This is the spreadsheet drag-fill convention.
 *
 * Replaces the old regex-based adjustFormula on A1 strings; the AST walk
 * is correct under arbitrary row/column reorderings whereas the regex
 * implicitly assumed numeric coords matched current indices.
 */
export function adjustStableAstForCopy(
  node: StableFormulaNode,
  resolver: SheetResolver,
  currentSheet: Sheet,
  sourceRow: number,
  sourceCol: number,
  targetRow: number,
  targetCol: number,
): StableFormulaNode {
  const out: StableFormulaNode = { type: node.type };
  if (node.value !== undefined) out.value = node.value;
  if (node.functionName !== undefined) out.functionName = node.functionName;
  if (node.operator !== undefined) out.operator = node.operator;
  if (node.cellRef) {
    out.cellRef = rebaseRef(
      node.cellRef,
      resolver,
      currentSheet,
      sourceRow,
      sourceCol,
      targetRow,
      targetCol,
    );
  }
  if (node.rangeRef) {
    out.rangeRef = {
      start: rebaseRef(
        node.rangeRef.start,
        resolver,
        currentSheet,
        sourceRow,
        sourceCol,
        targetRow,
        targetCol,
      ),
      end: rebaseRef(
        node.rangeRef.end,
        resolver,
        currentSheet,
        sourceRow,
        sourceCol,
        targetRow,
        targetCol,
      ),
    };
  }
  if (node.args) {
    out.args = node.args.map((a) =>
      adjustStableAstForCopy(a, resolver, currentSheet, sourceRow, sourceCol, targetRow, targetCol),
    );
  }
  if (node.left) {
    out.left = adjustStableAstForCopy(
      node.left, resolver, currentSheet, sourceRow, sourceCol, targetRow, targetCol,
    );
  }
  if (node.right) {
    out.right = adjustStableAstForCopy(
      node.right, resolver, currentSheet, sourceRow, sourceCol, targetRow, targetCol,
    );
  }
  return out;
}

function rebaseRef(
  ref: StableCellRef,
  resolver: SheetResolver,
  currentSheet: Sheet,
  sourceRow: number,
  sourceCol: number,
  targetRow: number,
  targetCol: number,
): StableCellRef {
  const sheet = ref.sheetName ? (resolver.getSheet(ref.sheetName) ?? currentSheet) : currentSheet;

  let newRowId = ref.rowId;
  let newColId = ref.colId;

  if (!ref.rowAbsolute) {
    const cur = sheet.getRowIndex(ref.rowId);
    if (cur !== undefined) {
      const offset = cur - sourceRow;
      const newRow = targetRow + offset;
      // A relative ref that lands above row 0 / left of column A is #REF! in
      // Excel, not a silent clamp to the edge. A sentinel id that resolves to
      // nothing makes the ref render and evaluate as #REF!.
      newRowId = newRow < 0 ? INVALID_REF_ID : sheet.ensureRowId(newRow);
    }
  }
  if (!ref.colAbsolute) {
    const cur = sheet.getColIndex(ref.colId);
    if (cur !== undefined) {
      const offset = cur - sourceCol;
      const newCol = targetCol + offset;
      newColId = newCol < 0 ? INVALID_REF_ID : sheet.ensureColId(newCol);
    }
  }

  return {
    rowId: newRowId,
    colId: newColId,
    rowAbsolute: ref.rowAbsolute,
    colAbsolute: ref.colAbsolute,
    sheetName: ref.sheetName,
  };
}

/**
 * Walk the stable AST and collect dependency keys for the formula graph.
 * Cross-sheet refs are tracked too: their row/col ids resolve against the
 * target sheet and are globally unique, so `rowId:colId` identifies the right
 * cell across the whole workbook and editing it dirties this formula.
 */
// Functions that consume only a reference's shape/position, never its cell
// values (ROWS/COLUMNS/ROW/COLUMN). A direct ref argument to one of these is not
// a value dependency: its result changes only on a structural edit (which
// triggers a full recalc), so tracking it would add spurious edges — and, when
// the formula sits inside its own referenced range, a false #CIRCULAR!.
const REF_SHAPE_FUNCTIONS = new Set(['ROWS', 'COLUMNS', 'ROW', 'COLUMN']);

export function collectStableDependencies(node: StableFormulaNode): FormulaDependencies {
  const cells = new Set<string>();
  const ranges: RangeDependency[] = [];
  // `skipDirectRef` suppresses collecting this node's OWN cell/range ref when it
  // is the direct argument of a shape-only function; it does not propagate into
  // nested function calls (their args are collected normally).
  const visit = (n: StableFormulaNode, skipDirectRef: boolean): void => {
    if (n.cellRef && !skipDirectRef) {
      cells.add(`${n.cellRef.rowId}:${n.cellRef.colId}`);
    }
    // A range is kept as its two corner keys — a rectangle. Containment is
    // tested at recalc time, so a cell that is empty when the formula is
    // entered, or a row/column inserted into the range later, is still tracked.
    if (n.rangeRef && !skipDirectRef) {
      ranges.push({
        startKey: `${n.rangeRef.start.rowId}:${n.rangeRef.start.colId}`,
        endKey: `${n.rangeRef.end.rowId}:${n.rangeRef.end.colId}`,
      });
    }
    const argsAreShapeOnly = !!n.functionName && REF_SHAPE_FUNCTIONS.has(n.functionName.toUpperCase());
    for (const a of n.args ?? []) visit(a, argsAreShapeOnly);
    if (n.left) visit(n.left, false);
    if (n.right) visit(n.right, false);
  };
  visit(node, false);
  return { cells, ranges };
}

/** True when a formula contains a volatile function (RAND, NOW, OFFSET, …). */
export function formulaIsVolatile(node: StableFormulaNode): boolean {
  if (node.functionName && volatileFunctions.has(node.functionName.toUpperCase())) {
    return true;
  }
  for (const a of node.args ?? []) {
    if (formulaIsVolatile(a)) return true;
  }
  if (node.left && formulaIsVolatile(node.left)) return true;
  if (node.right && formulaIsVolatile(node.right)) return true;
  return false;
}
