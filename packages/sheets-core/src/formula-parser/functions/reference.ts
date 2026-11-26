// Reference-aware functions.
//
// Unlike every other category, these need an argument's *reference* — its
// row/column coordinates — rather than its evaluated value. They receive the
// raw argument AST and the current cell position (see RefFn / RefFnContext).
//
// Dependency-tracking note: the parser records dependencies statically, so a
// cell that OFFSETs or INDIRECTs to another cell is not re-evaluated when that
// target changes. The computed value is correct on (re)evaluation; live
// recalculation would need volatile-function support in the engine.

import { MAX_SHEET_COLS, MAX_SHEET_ROWS } from '../types';
import type { CellReference, EvaluationContext, ParsedFormulaNode, RangeReference } from '../types';
import type { RefFn, RefFnContext } from './helpers';
import { toNum, toText, flatten } from './helpers';
import { lookupFunctions } from './lookup';
import { parseCellReference, parseRangeReference, columnIndexToLabel } from '../cell-reference';

/** Row count of an evaluated array value (fallback when the arg isn't a range). */
function arrayRowCount(v: unknown): number {
  return Array.isArray(v) ? (v.length > 0 && Array.isArray(v[0]) ? v.length : 1) : 1;
}

/** Column count of an evaluated array value. */
function arrayColCount(v: unknown): number {
  if (!Array.isArray(v)) return 1;
  return v.length > 0 && Array.isArray(v[0]) ? (v[0] as unknown[]).length : v.length;
}

/** Absolute row of a reference (relative refs are stored as offsets). */
function resolveRow(ref: CellReference, currentRow: number): number {
  return ref.rowAbsolute ? ref.row : currentRow + ref.row;
}

/** Absolute column of a reference. */
function resolveCol(ref: CellReference, currentCol: number): number {
  return ref.colAbsolute ? ref.col : currentCol + ref.col;
}

/** An absolute single-cell RangeReference bound, used to read computed ranges. */
function absBound(row: number, col: number, sheetName?: string): CellReference {
  return { row, col, rowAbsolute: true, colAbsolute: true, sheetName };
}


/** What reference resolution needs from the evaluator. */
export type RefEnv = Pick<RefFnContext, 'ctx' | 'currentRow' | 'currentCol' | 'evaluate'>;

/** A reference resolved to absolute, normalised (top <= bottom, left <= right), 0-based bounds. */
export interface ResolvedRef {
  sheetName?: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** A resolved reference that also records whether it was written as a range (`A1:A1`) or a cell. */
interface IndirectTarget extends ResolvedRef {
  isRange: boolean;
}

/** Both references must sit on the same sheet (or both on the formula's own); anything else is #VALUE!. */
const sameSheet = (a: ResolvedRef, b: ResolvedRef): string | undefined => {
  if (a.sheetName !== b.sheetName) throw new Error('#VALUE!');
  return a.sheetName;
};

/**
 * The reference an AST node denotes, as bounds, or undefined when the node is
 * not a reference (a literal, a computed array, a function that returns a value).
 * Understands cells, ranges (including whole-column/row), the `:` range operator
 * over any references, the ` ` intersection operator, OFFSET, INDEX, INDIRECT
 * and the spill operator. Throws a `#` error token when a reference-valued
 * expression fails (`#NULL!` for an empty intersection, `#REF!` out of bounds).
 */
export function resolveReference(node: ParsedFormulaNode, env: RefEnv): ResolvedRef | undefined {
  const { currentRow, currentCol } = env;
  switch (node.type) {
    case 'cell': {
      const ref = node.cellRef as CellReference;
      const row = resolveRow(ref, currentRow);
      const col = resolveCol(ref, currentCol);
      return { sheetName: ref.sheetName, r1: row, c1: col, r2: row, c2: col };
    }
    case 'range': {
      const { start, end } = node.rangeRef as RangeReference;
      const [ra, rb] = [resolveRow(start, currentRow), resolveRow(end, currentRow)];
      const [ca, cb] = [resolveCol(start, currentCol), resolveCol(end, currentCol)];
      return {
        sheetName: start.sheetName ?? end.sheetName,
        r1: Math.min(ra, rb),
        c1: Math.min(ca, cb),
        r2: Math.max(ra, rb),
        c2: Math.max(ca, cb),
      };
    }
    case 'operator': {
      if (node.operator !== ':' && node.operator !== ' ') return undefined;
      const a = resolveStrict(node.left as ParsedFormulaNode, env);
      const b = resolveStrict(node.right as ParsedFormulaNode, env);
      const sheetName = sameSheet(a, b);
      if (node.operator === ':') {
        return {
          sheetName,
          r1: Math.min(a.r1, b.r1),
          c1: Math.min(a.c1, b.c1),
          r2: Math.max(a.r2, b.r2),
          c2: Math.max(a.c2, b.c2),
        };
      }
      const r1 = Math.max(a.r1, b.r1);
      const c1 = Math.max(a.c1, b.c1);
      const r2 = Math.min(a.r2, b.r2);
      const c2 = Math.min(a.c2, b.c2);
      if (r1 > r2 || c1 > c2) throw new Error('#NULL!');
      return { sheetName, r1, c1, r2, c2 };
    }
    case 'function':
      switch (node.functionName) {
        case 'OFFSET':
          return offsetTarget(node.args as ParsedFormulaNode[], env);
        case 'INDEX':
          return indexTarget(node.args as ParsedFormulaNode[], env);
        case 'INDIRECT':
          return indirectTarget(node.args as ParsedFormulaNode[], env);
        default:
          return undefined;
      }
    case 'error':
      throw new Error(node.value as string); // an error operand (a deleted reference) is that error
    case 'spillref': {
      const ref = node.cellRef as CellReference;
      const row = resolveRow(ref, currentRow);
      const col = resolveCol(ref, currentCol);
      const region = env.ctx.getSpillRegion?.(row, col, ref.sheetName);
      if (!region) throw new Error('#REF!'); // not a spill anchor
      return {
        sheetName: ref.sheetName,
        r1: row,
        c1: col,
        r2: row + region.rows - 1,
        c2: col + region.cols - 1,
      };
    }
    default:
      return undefined;
  }
}

/** `resolveReference` for an operand that must be a reference (`:`, ` `): anything else is #VALUE!. */
function resolveStrict(node: ParsedFormulaNode, env: RefEnv): ResolvedRef {
  const ref = resolveReference(node, env);
  if (!ref) throw new Error('#VALUE!');
  return ref;
}

/** Does this AST node denote a reference (without evaluating it)? Drives ISREF. */
export function isReferenceExpression(node: ParsedFormulaNode): boolean {
  switch (node.type) {
    case 'cell':
    case 'range':
    case 'ref3d':
    case 'spillref':
      return true;
    case 'operator':
      return node.operator === ':' || node.operator === ' ' || node.operator === ',';
    case 'function': {
      const args = node.args as ParsedFormulaNode[];
      if (node.functionName === 'OFFSET' || node.functionName === 'INDIRECT') return true;
      return node.functionName === 'INDEX' && args.length > 0 && isReferenceExpression(args[0]);
    }
    default:
      return false;
  }
}

/** Read the cells a resolved reference covers: one cell is a scalar, anything larger a matrix. */
export function readReference(ref: ResolvedRef, ctx: EvaluationContext): unknown {
  if (ref.r1 === ref.r2 && ref.c1 === ref.c2) return ctx.getCellValue(ref.r1, ref.c1, undefined, ref.sheetName);
  const range: RangeReference = {
    start: absBound(ref.r1, ref.c1, ref.sheetName),
    end: absBound(ref.r2, ref.c2, ref.sheetName),
  };
  // A reference spanning a whole column / row is bounded to the used range.
  if (ref.r2 - ref.r1 + 1 >= MAX_SHEET_ROWS) range.whole = 'column';
  else if (ref.c2 - ref.c1 + 1 >= MAX_SHEET_COLS) range.whole = 'row';
  return ctx.getRangeValues(range, undefined, ref.sheetName);
}

/** The cell an OFFSET lands on, as the rectangle it covers. */
function offsetTarget(args: ParsedFormulaNode[], env: RefEnv): ResolvedRef {
  const base = args[0] && resolveReference(args[0], env);
  if (!base) throw new Error('#REF!');
  const { evaluate } = env;
  const baseHeight = base.r2 - base.r1 + 1;
  const baseWidth = base.c2 - base.c1 + 1;
  const rowsOffset = args[1] !== undefined ? Math.trunc(toNum(evaluate(args[1]))) : 0;
  const colsOffset = args[2] !== undefined ? Math.trunc(toNum(evaluate(args[2]))) : 0;
  const height = args[3] !== undefined ? Math.trunc(toNum(evaluate(args[3]))) : baseHeight;
  const width = args[4] !== undefined ? Math.trunc(toNum(evaluate(args[4]))) : baseWidth;
  if (height < 1 || width < 1) throw new Error('#REF!');

  const newRow = base.r1 + rowsOffset;
  const newCol = base.c1 + colsOffset;
  if (newRow < 0 || newCol < 0) throw new Error('#REF!');
  return {
    sheetName: base.sheetName,
    r1: newRow,
    c1: newCol,
    r2: newRow + height - 1,
    c2: newCol + width - 1,
  };
}

/**
 * The cell an INDEX picks out of a reference, or undefined when the first
 * argument is not a reference (INDEX then works on the evaluated array).
 */
function indexTarget(args: ParsedFormulaNode[], env: RefEnv): ResolvedRef | undefined {
  const base = args[0] && resolveReference(args[0], env);
  if (!base) return undefined;
  const rows = base.r2 - base.r1 + 1;
  const cols = base.c2 - base.c1 + 1;
  let r = args[1] !== undefined ? toNum(env.evaluate(args[1])) : 1;
  let c = args[2] !== undefined ? toNum(env.evaluate(args[2])) : 1;

  // Single-axis form: INDEX(row_vector, n) or INDEX(col_vector, n).
  if (args[2] === undefined) {
    if (rows === 1) {
      c = r;
      r = 1;
    } else if (cols === 1) {
      c = 1;
    } else {
      throw new Error('#REF!');
    }
  }
  if (r < 1 || c < 1 || r > rows || c > cols) throw new Error('#REF!');
  const row = base.r1 + Math.trunc(r) - 1;
  const col = base.c1 + Math.trunc(c) - 1;
  return { sheetName: base.sheetName, r1: row, c1: col, r2: row, c2: col };
}

/** The reference an INDIRECT text names (A1 style, optionally sheet-qualified, whole columns/rows included). */
function indirectTarget(args: ParsedFormulaNode[], env: RefEnv): IndirectTarget {
  const text = toText(env.evaluate(args[0]));
  const range = parseRangeReference(text);
  if (range) {
    const [ra, rb] = [range.start.row, range.end.row];
    const [ca, cb] = [range.start.col, range.end.col];
    return {
      sheetName: range.start.sheetName ?? range.end.sheetName,
      r1: Math.min(ra, rb),
      c1: Math.min(ca, cb),
      r2: Math.max(ra, rb),
      c2: Math.max(ca, cb),
      isRange: true,
    };
  }
  const cell = parseCellReference(text);
  if (cell) {
    return { sheetName: cell.sheetName, r1: cell.row, c1: cell.col, r2: cell.row, c2: cell.col, isRange: false };
  }
  throw new Error('#REF!');
}

export const referenceFunctions: Record<string, RefFn> = {
  ROW: (rc) => {
    if (rc.args.length === 0) return rc.currentRow + 1;
    return resolveStrict(rc.args[0], rc).r1 + 1;
  },

  COLUMN: (rc) => {
    if (rc.args.length === 0) return rc.currentCol + 1;
    return resolveStrict(rc.args[0], rc).c1 + 1;
  },

  // ROWS/COLUMNS are ref-aware so they read a range's true size from the
  // reference rather than the materialized values. This is both correct for huge
  // ranges (=ROWS(A1:A1048576) is 1048576, not the populated-data extent that
  // getRangeValues clamps to for DoS safety) and free of that per-eval extent
  // scan. A non-range argument (an array literal or a function that returns an
  // array) falls back to measuring the evaluated value.
  ROWS: (rc) => {
    const node = rc.args[0];
    if (!node) throw new Error('#VALUE!'); // ROWS() with no argument
    const ref = resolveReference(node, rc);
    return ref ? ref.r2 - ref.r1 + 1 : arrayRowCount(rc.evaluate(node));
  },

  COLUMNS: (rc) => {
    const node = rc.args[0];
    if (!node) throw new Error('#VALUE!'); // COLUMNS() with no argument
    const ref = resolveReference(node, rc);
    return ref ? ref.c2 - ref.c1 + 1 : arrayColCount(rc.evaluate(node));
  },

  // Ref-aware so a huge range's blank count is the nominal size minus the
  // populated cells, rather than iterating (and clamping) millions of values.
  // getRangeValues now yields null for empty cells, so non-blank counting is
  // accurate within the returned (possibly extent-clamped) values.
  COUNTBLANK: ({ args, currentRow, currentCol, evaluate }) => {
    const node = args[0];
    if (node?.type === 'range' && node.rangeRef) {
      const rows = Math.abs(resolveRow(node.rangeRef.end, currentRow) - resolveRow(node.rangeRef.start, currentRow)) + 1;
      const cols = Math.abs(resolveCol(node.rangeRef.end, currentCol) - resolveCol(node.rangeRef.start, currentCol)) + 1;
      const nonBlank = flatten([evaluate(node)]).filter((v) => v != null && v !== '').length;
      return Math.max(0, rows * cols - nonBlank);
    }
    return flatten(args.map((a) => evaluate(a))).filter((v) => v == null || v === '').length;
  },

  OFFSET: (rc) => {
    const target = offsetTarget(rc.args, rc);
    return readReference(target, rc.ctx);
  },

  // INDEX is a reference function so `INDEX(A1:A5,3)` can be a range endpoint
  // (`A1:INDEX(A1:A5,3)`) and so `INDEX(A:A,5)` indexes the nominal column rather
  // than its used-range clamp. Over a computed array it is the ordinary INDEX.
  INDEX: (rc) => {
    const target = indexTarget(rc.args, rc);
    if (target) return rc.ctx.getCellValue(target.r1, target.c1, undefined, target.sheetName);
    return lookupFunctions.INDEX(
      rc.args.map((a) => rc.evaluate(a)),
      rc.ctx
    );
  },

  INDIRECT: (rc) => {
    const target = indirectTarget(rc.args, rc);
    if (!target.isRange) return rc.ctx.getCellValue(target.r1, target.c1, undefined, target.sheetName);
    const range: RangeReference = {
      start: absBound(target.r1, target.c1, target.sheetName),
      end: absBound(target.r2, target.c2, target.sheetName),
    };
    if (target.r2 - target.r1 + 1 >= MAX_SHEET_ROWS) range.whole = 'column';
    else if (target.c2 - target.c1 + 1 >= MAX_SHEET_COLS) range.whole = 'row';
    return rc.ctx.getRangeValues(range, undefined, target.sheetName);
  },

  ISREF: ({ args }) => !!args[0] && isReferenceExpression(args[0]),

  ISFORMULA: ({ args, ctx, currentRow, currentCol }) => {
    const node = args[0];
    if (node?.type !== 'cell' || !node.cellRef) throw new Error('#VALUE!');
    if (!ctx.isCellFormula) return false;
    return ctx.isCellFormula(
      resolveRow(node.cellRef, currentRow),
      resolveCol(node.cellRef, currentCol),
      node.cellRef.sheetName
    );
  },

  CELL: ({ args, ctx, currentRow, currentCol, evaluate }) => {
    const infoType = toText(evaluate(args[0])).toLowerCase();

    let row = currentRow;
    let col = currentCol;
    let sheetName: string | undefined;
    if (args[1] !== undefined) {
      const ref = resolveStrict(args[1], { ctx, currentRow, currentCol, evaluate });
      row = ref.r1;
      col = ref.c1;
      sheetName = ref.sheetName;
    }

    switch (infoType) {
      case 'address':
        return `$${columnIndexToLabel(col)}$${row + 1}`;
      case 'row':
        return row + 1;
      case 'col':
        return col + 1;
      case 'contents':
        return ctx.getCellValue(row, col, undefined, sheetName) ?? 0;
      case 'type': {
        const value = ctx.getCellValue(row, col, undefined, sheetName);
        if (value == null || value === '') return 'b';
        return typeof value === 'string' ? 'l' : 'v';
      }
      default:
        throw new Error('#VALUE!');
    }
  },
};
