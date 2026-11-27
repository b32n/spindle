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
import { MAX_SHEET_COLS, MAX_SHEET_ROWS } from './types';
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
  /**
   * A whole-column (`A:C`) or whole-row (`1:3`) range. Only the bounded axis
   * holds real ids: the other axis' id is WHOLE_AXIS_ID on both corners, so
   * inserting or deleting rows never touches a whole-column ref (and vice versa).
   */
  whole?: 'column' | 'row';
}

export interface StableFormulaNode {
  type: ParsedFormulaNode['type'];
  value?: unknown;
  cellRef?: StableCellRef;
  rangeRef?: StableRangeRef;
  /** Last sheet of a 3D reference (`Sheet1:Sheet3!A1`); the first is the ref's `sheetName`. */
  sheetEnd?: string;
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

/**
 * Id of the unbounded axis of a whole-column / whole-row ref (real ids are
 * 12-char base32, so '*' can never collide). It resolves to no row or column on
 * purpose: code that meets it must go through the `whole` flag, never an index.
 */
export const WHOLE_AXIS_ID = '*';

export interface SheetResolver {
  getSheet(name?: string): Sheet | undefined;
}

// ============================================================================
// Numeric AST → Stable AST  (called on setFormula)
// ============================================================================

/** The scalar fields every walker copies; children and refs are converted by each walker. */
function shell(node: { type: StableFormulaNode['type'] } & Partial<StableFormulaNode>): StableFormulaNode {
  const out: StableFormulaNode = { type: node.type };
  if (node.value !== undefined) out.value = node.value;
  if (node.functionName !== undefined) out.functionName = node.functionName;
  if (node.operator !== undefined) out.operator = node.operator;
  if (node.sheetEnd !== undefined) out.sheetEnd = node.sheetEnd;
  return out;
}

export function toStableAst(
  node: ParsedFormulaNode,
  resolver: SheetResolver,
  currentSheet: Sheet,
  currentRow: number,
  currentCol: number,
): StableFormulaNode {
  const out = shell(node as StableFormulaNode);
  // A 3D reference spans sheets whose row/column ids differ, so it is positional.
  const positional = node.type === 'ref3d';
  const convert = (ref: CellReference, axis?: 'column' | 'row') =>
    cellRefToStable(ref, resolver, currentSheet, currentRow, currentCol, positional, axis);
  if (node.cellRef) out.cellRef = convert(node.cellRef);
  if (node.rangeRef) {
    const whole = node.rangeRef.whole;
    out.rangeRef = { start: convert(node.rangeRef.start, whole), end: convert(node.rangeRef.end, whole) };
    if (whole) out.rangeRef.whole = whole;
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
  positional: boolean,
  axis?: 'column' | 'row',
): StableCellRef {
  // Parser convention: absolute refs hold absolute positions; relatives hold offsets.
  const absRow = ref.rowAbsolute ? ref.row : currentRow + ref.row;
  const absCol = ref.colAbsolute ? ref.col : currentCol + ref.col;
  const target = ref.sheetName ? resolver.getSheet(ref.sheetName) : currentSheet;
  if (!target || positional) {
    // Named sheet doesn't exist yet — defer. Don't bind to currentSheet (that
    // collides with the formula's own cell → false #CIRCULAR!, and never
    // recovers). Keep the position; resolve when the sheet appears. A 3D
    // reference is always kept this way: it addresses the same position on
    // every sheet of its span.
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
    rowId: axis === 'column' ? WHOLE_AXIS_ID : target.ensureRowId(absRow),
    colId: axis === 'row' ? WHOLE_AXIS_ID : target.ensureColId(absCol),
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
  const out = shell(node) as ParsedFormulaNode;
  if (node.cellRef) {
    out.cellRef = cellRefFromStable(node.cellRef, resolver, currentSheet, currentRow, currentCol);
  }
  if (node.rangeRef) {
    const whole = node.rangeRef.whole;
    // The unbounded axis spans the whole sheet: first row/column up to the last.
    const bound = (isEnd: boolean) =>
      whole === 'column'
        ? { axis: 'column' as const, fixed: isEnd ? MAX_SHEET_ROWS - 1 : 0 }
        : whole === 'row'
          ? { axis: 'row' as const, fixed: isEnd ? MAX_SHEET_COLS - 1 : 0 }
          : undefined;
    out.rangeRef = {
      start: cellRefFromStable(node.rangeRef.start, resolver, currentSheet, currentRow, currentCol, bound(false)),
      end: cellRefFromStable(node.rangeRef.end, resolver, currentSheet, currentRow, currentCol, bound(true)),
    };
    if (whole) out.rangeRef.whole = whole;
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
  /** For a whole-column/row corner: which axis is unbounded and the position it takes (first or last). */
  unbounded?: { axis: 'column' | 'row'; fixed: number },
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
  // its stored absolute position now that the sheet exists. A 3D reference is
  // always positional.
  if (ref.unresolvedRow !== undefined && ref.unresolvedCol !== undefined) {
    const uRow = ref.unresolvedRow;
    const uCol = ref.unresolvedCol;
    // A relative ref copied above row 1 / left of column A has no position left.
    if (uRow < 0 || uCol < 0) throw new StableRefDeletedError();
    return {
      row: ref.rowAbsolute ? uRow : uRow - currentRow,
      col: ref.colAbsolute ? uCol : uCol - currentCol,
      rowAbsolute: ref.rowAbsolute,
      colAbsolute: ref.colAbsolute,
      sheetName: ref.sheetName,
    };
  }

  const absRow = unbounded?.axis === 'column' ? unbounded.fixed : sheet.getRowIndex(ref.rowId);
  const absCol = unbounded?.axis === 'row' ? unbounded.fixed : sheet.getColIndex(ref.colId);
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

// Same order as the parser: comparison < & < +,- < *,/ < ^ < (percent) < range operators.
// A leading sign (5) and `%` (6) bind tighter than any binary operator; the
// reference operators bind tighter still: ` ` (8) and `:` (9). `,` (union) is
// always printed inside its own parentheses.
const OPERATOR_PRECEDENCE: Record<string, number> = {
  '=': 0, '<>': 0, '<': 0, '>': 0, '<=': 0, '>=': 0,
  '&': 1, '+': 2, '-': 2, '*': 3, '/': 3, '^': 4,
  ' ': 8, ':': 9,
};


const quoteString = (text: string): string => `"${text.replace(/"/g, '""')}"`;

/** An array-constant element: errors print bare, text quoted, booleans as TRUE/FALSE. */
function renderArrayElement(value: unknown): string {
  if (typeof value === 'object' && value !== null) return (value as { error: string }).error;
  if (typeof value === 'string') return quoteString(value);
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}

function renderNode(
  node: StableFormulaNode,
  resolver: SheetResolver,
  currentSheet: Sheet,
): string {
  const render = (n: StableFormulaNode) => renderNode(n, resolver, currentSheet);
  switch (node.type) {
    case 'number':
      return node.value === undefined ? '' : String(node.value); // an omitted argument
    case 'string':
      return quoteString(String(node.value));
    case 'variable':
      return String(node.value);
    case 'error':
      return String(node.value);
    case 'array':
      return `{${(node.value as unknown[][]).map((row) => row.map(renderArrayElement).join(',')).join(';')}}`;
    case 'cell':
      return node.cellRef ? renderRef(node.cellRef, resolver, currentSheet) : '#REF!';
    case 'range':
      return renderRange(node, resolver, currentSheet);
    case 'ref3d':
      return renderRef3d(node, resolver, currentSheet);
    case 'spillref':
      return `${renderRef(node.cellRef as StableCellRef, resolver, currentSheet)}#`;
    case 'percent': {
      const operand = node.left as StableFormulaNode;
      const text = render(operand);
      return `${operand.type === 'operator' && !isUnarySign(operand) && operand.operator !== ',' ? `(${text})` : text}%`;
    }
    case 'implicit': {
      const operand = node.left as StableFormulaNode;
      const text = render(operand);
      // `@` takes a range-level operand: anything looser needs its parentheses.
      const loose = operand.type === 'percent' || (operand.type === 'operator' && operand.operator !== ':' && operand.operator !== ',');
      return `@${loose ? `(${text})` : text}`;
    }
    case 'function':
      return `${node.functionName as string}(${(node.args as StableFormulaNode[]).map(render).join(',')})`;
    case 'operator': {
      const op = node.operator ?? '';
      if (op === ',') return `(${unionOperands(node).map(render).join(',')})`;
      // A leading sign is stored as `"" op x`; it binds tighter than any binary
      // operator, so a binary operand needs its parentheses back.
      if ((op === '-' || op === '+') && isEmptyLiteral(node.left) && node.right) {
        const r = render(node.right);
        const wrap = (node.right.type === 'operator' && !isUnarySign(node.right) && node.right.operator !== ',') || node.right.type === 'percent';
        return `${op}${wrap ? `(${r})` : r}`;
      }
      const prec = OPERATOR_PRECEDENCE[op] ?? 0;
      const l = node.left ? render(node.left) : '';
      const r = node.right ? render(node.right) : '';
      // Parentheses are not kept in the AST; restore the ones precedence needs.
      // All binary operators are left-associative, so an equal-precedence right
      // operand (a-(b-c)) needs them too. A union prints its own parentheses.
      const needsParens = (n: StableFormulaNode | undefined, strict: boolean): boolean => {
        if (n?.type !== 'operator' || isUnarySign(n) || n.operator === ',') return false;
        const p = OPERATOR_PRECEDENCE[n.operator as string];
        return strict ? p <= prec : p < prec;
      };
      return `${needsParens(node.left, false) ? `(${l})` : l}${op}${needsParens(node.right, true) ? `(${r})` : r}`;
    }
    default:
      return '';
  }
}

/** The operands of a union chain `((a,b),c)`, flattened to `a, b, c`. */
function unionOperands(node: StableFormulaNode): StableFormulaNode[] {
  const left = node.left as StableFormulaNode;
  const head = left.type === 'operator' && left.operator === ',' ? unionOperands(left) : [left];
  return [...head, node.right as StableFormulaNode];
}

function isEmptyLiteral(node: StableFormulaNode | undefined): boolean {
  return !!node && node.type === 'string' && node.value === '';
}

/** `"" - x` / `"" + x`: the parser's form of a unary sign. */
function isUnarySign(node: StableFormulaNode): boolean {
  return node.type === 'operator' && (node.operator === '-' || node.operator === '+') && isEmptyLiteral(node.left);
}

/** A name that Excel accepts without quotes: no leading digit, only letters, digits, `_` and `.`, and not a cell address. */
function isPlainSheetName(name: string): boolean {
  return (
    /^[\p{L}_\\][\p{L}\p{N}_.]*$/u.test(name) &&
    !/^[A-Za-z]{1,3}\d+$/.test(name) &&
    !/^(?:TRUE|FALSE)$/i.test(name)
  );
}

/** `Sheet1!`, `'My Sheet'!` or the 3D `Sheet1:Sheet3!` / `'Sheet 1:Sheet 3'!`. */
function sheetPrefix(first: string, last?: string): string {
  const names = last === undefined ? [first] : [first, last];
  if (names.every(isPlainSheetName)) return `${names.join(':')}!`;
  return `'${names.map((n) => n.replace(/'/g, "''")).join(':')}'!`;
}

function renderRef(
  ref: StableCellRef,
  resolver: SheetResolver,
  currentSheet: Sheet,
  withSheet = true,
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
  // (A negative stored position is a positional ref copied off the top/left of the sheet.)
  if (row === undefined || col === undefined || row < 0 || col < 0) return '#REF!';
  const colLabel = columnIndexToLabel(col);
  const cell =
    (ref.colAbsolute ? '$' : '') + colLabel + (ref.rowAbsolute ? '$' : '') + (row + 1);
  return withSheet && ref.sheetName ? sheetPrefix(ref.sheetName) + cell : cell;
}

/** The bare text of a range's corner (`A1`, `$A`, `$3`) without a sheet prefix. */
function renderCorner(
  ref: StableCellRef,
  whole: 'column' | 'row' | undefined,
  resolver: SheetResolver,
  currentSheet: Sheet,
): string {
  if (!whole) return renderRef(ref, resolver, currentSheet, false);
  let index: number | undefined;
  if (ref.unresolvedRow !== undefined && ref.unresolvedCol !== undefined) {
    index = whole === 'column' ? ref.unresolvedCol : ref.unresolvedRow;
  } else {
    const sheet = ref.sheetName ? (resolver.getSheet(ref.sheetName) ?? currentSheet) : currentSheet;
    index = whole === 'column' ? sheet.getColIndex(ref.colId) : sheet.getRowIndex(ref.rowId);
  }
  if (index === undefined || index < 0) return '#REF!';
  return whole === 'column'
    ? (ref.colAbsolute ? '$' : '') + columnIndexToLabel(index)
    : (ref.rowAbsolute ? '$' : '') + (index + 1);
}

function renderRange(node: StableFormulaNode, resolver: SheetResolver, currentSheet: Sheet): string {
  const range = node.rangeRef as StableRangeRef;
  const start = renderCorner(range.start, range.whole, resolver, currentSheet);
  const end = renderCorner(range.end, range.whole, resolver, currentSheet);
  // A range with a deleted corner is the single token #REF! (as in Excel), which also parses back.
  if (start === '#REF!' || end === '#REF!') return '#REF!';
  const sheet = range.start.sheetName;
  return `${sheet ? sheetPrefix(sheet) : ''}${start}:${end}`;
}

function renderRef3d(node: StableFormulaNode, resolver: SheetResolver, currentSheet: Sheet): string {
  const first = (node.cellRef ?? (node.rangeRef as StableRangeRef).start).sheetName as string;
  const prefix = sheetPrefix(first, node.sheetEnd);
  if (node.cellRef) {
    const cell = renderCorner(node.cellRef, undefined, resolver, currentSheet);
    return cell === '#REF!' ? cell : prefix + cell;
  }
  const range = node.rangeRef as StableRangeRef;
  const start = renderCorner(range.start, range.whole, resolver, currentSheet);
  const end = renderCorner(range.end, range.whole, resolver, currentSheet);
  return start === '#REF!' || end === '#REF!' ? '#REF!' : `${prefix}${start}:${end}`;
}

// ============================================================================
// Sheet-name edits (rename, delete) on 3D references
// ============================================================================

/** Call `visit` on every node of an AST, parents before children. */
export function walkStableAst(node: StableFormulaNode, visit: (n: StableFormulaNode) => void): void {
  visit(node);
  node.args?.forEach((a) => walkStableAst(a, visit));
  if (node.left) walkStableAst(node.left, visit);
  if (node.right) walkStableAst(node.right, visit);
}

/** Point every reference to sheet `oldName` (including a 3D span's endpoints) at `newName`. */
export function renameSheetInAst(root: StableFormulaNode, oldName: string, newName: string): void {
  walkStableAst(root, (n) => {
    if (n.cellRef?.sheetName === oldName) n.cellRef.sheetName = newName;
    if (n.rangeRef?.start.sheetName === oldName) n.rangeRef.start.sheetName = newName;
    if (n.rangeRef?.end.sheetName === oldName) n.rangeRef.end.sheetName = newName;
    if (n.sheetEnd === oldName) n.sheetEnd = newName;
  });
}

/**
 * Before sheet `deleted` is removed: pull a 3D span's endpoint inward when it is
 * that sheet (Excel: "if you delete Sheet2 or Sheet6, Excel adjusts the
 * calculation to include the sheets between them"). `order` is the sheet-name
 * order before the delete. A span that was only the deleted sheet keeps its
 * names, which no longer resolve, so it evaluates to #REF!.
 */
export function shrinkThreeDSpansForDelete(root: StableFormulaNode, deleted: string, order: string[]): void {
  walkStableAst(root, (n) => {
    if (n.type !== 'ref3d' || n.sheetEnd === undefined) return;
    const firstRef = n.cellRef ?? (n.rangeRef as StableRangeRef).start;
    const a = order.indexOf(firstRef.sheetName as string);
    const b = order.indexOf(n.sheetEnd);
    if (a < 0 || b < 0 || a === b) return;
    const inward = a < b ? 1 : -1; // from the first sheet towards the last
    if (order[a] === deleted) setFirstSheet(n, order[a + inward]);
    if (order[b] === deleted) n.sheetEnd = order[b - inward];
  });
}

function setFirstSheet(node: StableFormulaNode, name: string): void {
  if (node.cellRef) node.cellRef.sheetName = name;
  if (node.rangeRef) {
    node.rangeRef.start.sheetName = name;
    node.rangeRef.end.sheetName = name;
  }
}

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
 * Whole-column / whole-row refs move along their bounded axis only (`A:A` copied
 * one column right is `B:B`; copied down it is unchanged).
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
  const out = shell(node);
  const rebase = (ref: StableCellRef) =>
    rebaseRef(ref, resolver, currentSheet, sourceRow, sourceCol, targetRow, targetCol);
  if (node.cellRef) out.cellRef = rebase(node.cellRef);
  if (node.rangeRef) {
    out.rangeRef = { start: rebase(node.rangeRef.start), end: rebase(node.rangeRef.end) };
    if (node.rangeRef.whole) out.rangeRef.whole = node.rangeRef.whole;
  }
  const recurse = (n: StableFormulaNode) =>
    adjustStableAstForCopy(n, resolver, currentSheet, sourceRow, sourceCol, targetRow, targetCol);
  if (node.args) out.args = node.args.map(recurse);
  if (node.left) out.left = recurse(node.left);
  if (node.right) out.right = recurse(node.right);
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
  // A positional ref (3D, or a sheet that does not exist yet) shifts its stored position.
  if (ref.unresolvedRow !== undefined && ref.unresolvedCol !== undefined) {
    return {
      ...ref,
      unresolvedRow: ref.rowAbsolute ? ref.unresolvedRow : ref.unresolvedRow + (targetRow - sourceRow),
      unresolvedCol: ref.colAbsolute ? ref.unresolvedCol : ref.unresolvedCol + (targetCol - sourceCol),
    };
  }

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
 *
 * A whole-column / whole-row range is one rectangle bounded on one axis only
 * (`axis`), never expanded into cells. A 3D reference is not listed at all: its
 * span can change with the sheet order, so its formula is volatile
 * (formulaIsVolatile) and recalculated on every pass instead.
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
    // A spill reference's shape is the anchor's result, so ROWS(A1#) still depends on A1.
    const listed = (!skipDirectRef || n.type === 'spillref') && n.type !== 'ref3d';
    if (n.cellRef && listed) {
      cells.add(`${n.cellRef.rowId}:${n.cellRef.colId}`);
    }
    // A range is kept as its two corner keys — a rectangle. Containment is
    // tested at recalc time, so a cell that is empty when the formula is
    // entered, or a row/column inserted into the range later, is still tracked.
    if (n.rangeRef && listed) {
      const dep: RangeDependency = {
        startKey: `${n.rangeRef.start.rowId}:${n.rangeRef.start.colId}`,
        endKey: `${n.rangeRef.end.rowId}:${n.rangeRef.end.colId}`,
      };
      if (n.rangeRef.whole) dep.axis = n.rangeRef.whole;
      ranges.push(dep);
    }
    const argsAreShapeOnly = !!n.functionName && REF_SHAPE_FUNCTIONS.has(n.functionName.toUpperCase());
    for (const a of n.args ?? []) visit(a, argsAreShapeOnly);
    if (n.left) visit(n.left, false);
    if (n.right) visit(n.right, false);
  };
  visit(node, false);
  return { cells, ranges };
}

/**
 * True when a formula must be recomputed on every pass: it contains a volatile
 * function (RAND, NOW, OFFSET, …), a 3D reference (whose span follows the sheet
 * order) or a `:` between computed references (`A1:INDEX(...)`, whose extent is
 * only known when it is evaluated).
 */
export function formulaIsVolatile(node: StableFormulaNode): boolean {
  if (node.functionName && volatileFunctions.has(node.functionName.toUpperCase())) {
    return true;
  }
  if (node.type === 'ref3d' || (node.type === 'operator' && node.operator === ':')) return true;
  for (const a of node.args ?? []) {
    if (formulaIsVolatile(a)) return true;
  }
  if (node.left && formulaIsVolatile(node.left)) return true;
  if (node.right && formulaIsVolatile(node.right)) return true;
  return false;
}
