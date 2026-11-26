// Formula parser and evaluator.
//
// parse() delegates syntax to tokenizer.ts / syntax.ts (precedence climbing with
// Excel's operator order, references recognised as tokens); this file keeps the
// registry of functions and evaluates the numeric AST.

import type { ParsedFormulaNode, ParseResult, EvaluationContext, CellReference } from './types';
import { cellReferenceToKey } from './cell-reference';
import { parseFormulaSyntax } from './syntax';
import { eagerFunctions, lazyFunctions, refFunctions } from './functions';
import type { EagerFn, LazyFn, RefFn } from './functions';
import { compareValues, excelEqual, toText } from './functions/helpers';
import { readReference, resolveReference, type RefEnv } from './functions/reference';

/** Ranges larger than this are not expanded into the best-effort `dependencies` set. */
const MAX_RANGE_DEP_CELLS = 100_000;

export class FormulaParser {
  private functions: Map<string, EagerFn> = new Map();
  private lazyFunctions: Map<string, LazyFn> = new Map();
  private refFunctions: Map<string, RefFn> = new Map();

  constructor() {
    this.registerBuiltInFunctions();
  }

  private registerBuiltInFunctions(): void {
    for (const [name, fn] of Object.entries(eagerFunctions)) {
      this.functions.set(name, fn);
    }
    for (const [name, fn] of Object.entries(lazyFunctions)) {
      this.lazyFunctions.set(name, fn);
    }
    for (const [name, fn] of Object.entries(refFunctions)) {
      this.refFunctions.set(name, fn);
    }
  }

  parse(formula: string, currentRow: number = 0, currentCol: number = 0): ParseResult {
    const dependencies = new Set<string>();
    let error: string | undefined;

    try {
      // Guard against a pathologically long formula, which would freeze the
      // main thread; reject it instead.
      if (formula.length > 10_000) {
        throw new Error('#ERROR! Formula too long');
      }

      // Remove leading = if present
      const expression = formula.startsWith('=') ? formula.slice(1) : formula;

      const ast = parseFormulaSyntax(expression, currentRow, currentCol);
      this.collectDependencies(ast, currentRow, currentCol, dependencies);

      return { ast, dependencies, error };
    } catch (e) {
      error = (e as Error).message; // everything thrown while parsing is an Error
      return {
        ast: { type: 'string', value: '' },
        dependencies,
        error,
      };
    }
  }

  /**
   * Best-effort set of the "row:col" keys (sheet-prefixed when cross-sheet) a
   * formula reads. The recalc graph derives its real dependencies from the stable
   * AST (collectStableDependencies); a whole-column/row range, a 3D reference
   * and an oversized range are not expanded here.
   */
  private collectDependencies(
    node: ParsedFormulaNode,
    currentRow: number,
    currentCol: number,
    dependencies: Set<string>
  ): void {
    const key = (ref: CellReference): string => {
      const k = cellReferenceToKey(ref, currentRow, currentCol);
      return ref.sheetName ? `${ref.sheetName}!${k}` : k;
    };
    if (node.type === 'cell' && node.cellRef) dependencies.add(key(node.cellRef));
    if (node.type === 'range' && node.rangeRef && !node.rangeRef.whole) {
      const { start, end } = node.rangeRef;
      const rows = [start, end].map((r) => (r.rowAbsolute ? r.row : currentRow + r.row));
      const cols = [start, end].map((r) => (r.colAbsolute ? r.col : currentCol + r.col));
      const [rLo, rHi] = [Math.min(...rows), Math.max(...rows)];
      const [cLo, cHi] = [Math.min(...cols), Math.max(...cols)];
      const sheetPrefix = start.sheetName ? `${start.sheetName}!` : '';
      if ((rHi - rLo + 1) * (cHi - cLo + 1) <= MAX_RANGE_DEP_CELLS) {
        for (let r = rLo; r <= rHi; r++) {
          for (let c = cLo; c <= cHi; c++) dependencies.add(`${sheetPrefix}${r}:${c}`);
        }
      }
    }
    for (const arg of node.args ?? []) this.collectDependencies(arg, currentRow, currentCol, dependencies);
    if (node.left) this.collectDependencies(node.left, currentRow, currentCol, dependencies);
    if (node.right) this.collectDependencies(node.right, currentRow, currentCol, dependencies);
  }

  /** Apply a binary operator to two scalar values (throws a `#…` error token on failure). */
  private applyOperator(operator: string, leftVal: unknown, rightVal: unknown): unknown {
    // Operators that accept any types (text-aware).
    switch (operator) {
      // Text concatenation.
      case '&':
        return toText(leftVal) + toText(rightVal);
      // Comparisons: numeric when both look numeric, otherwise
      // case-insensitive lexical (compareValues). Equality is type-aware
      // so "1" != 1 and "" != 0.
      case '>':
        return compareValues(leftVal, rightVal) > 0;
      case '<':
        return compareValues(leftVal, rightVal) < 0;
      case '>=':
        return compareValues(leftVal, rightVal) >= 0;
      case '<=':
        return compareValues(leftVal, rightVal) <= 0;
      case '=':
        return excelEqual(leftVal, rightVal);
      case '<>':
        return !excelEqual(leftVal, rightVal);
    }

    // Arithmetic operators require numbers
    const leftNum = Number(leftVal);
    const rightNum = Number(rightVal);

    if (isNaN(leftNum) || isNaN(rightNum)) {
      throw new Error('#VALUE!');
    }

    switch (operator) {
      case '+':
        return leftNum + rightNum;
      case '-':
        return leftNum - rightNum;
      case '*':
        return leftNum * rightNum;
      case '/':
        if (rightNum === 0) {
          throw new Error('#DIV/0!');
        }
        return leftNum / rightNum;
      case '^':
        return Math.pow(leftNum, rightNum);
      default:
        throw new Error(`#ERROR! Unknown operator: ${operator}`);
    }
  }

  /**
   * Element-wise operator over arrays, with Excel's broadcasting: a scalar
   * applies to every element and a single row/column stretches to the other
   * operand's size. Positions with no counterpart (mismatched sizes) are #N/A;
   * a per-element failure becomes that element's error value, so one bad cell
   * does not discard the whole array.
   */
  private broadcastOperator(operator: string, leftVal: unknown, rightVal: unknown): unknown[][] {
    const toGrid = (v: unknown): unknown[][] =>
      Array.isArray(v) ? (v as unknown[][]).map((row) => (Array.isArray(row) ? row : [row])) : [[v]];
    const l = toGrid(leftVal);
    const r = toGrid(rightVal);
    const dim = (g: unknown[][]) => ({ rows: g.length, cols: g.reduce((m, row) => Math.max(m, row.length), 0) });
    const ld = dim(l);
    const rd = dim(r);
    const rows = ld.rows === 1 ? rd.rows : rd.rows === 1 ? ld.rows : Math.max(ld.rows, rd.rows);
    const cols = ld.cols === 1 ? rd.cols : rd.cols === 1 ? ld.cols : Math.max(ld.cols, rd.cols);
    const at = (g: unknown[][], d: { rows: number; cols: number }, i: number, j: number): unknown | undefined => {
      const row = d.rows === 1 ? 0 : i;
      const col = d.cols === 1 ? 0 : j;
      if (row >= g.length || col >= g[row].length) return undefined;
      return g[row][col];
    };
    const out: unknown[][] = [];
    for (let i = 0; i < rows; i++) {
      const outRow: unknown[] = [];
      for (let j = 0; j < cols; j++) {
        const a = at(l, ld, i, j);
        const b = at(r, rd, i, j);
        if (a === undefined || b === undefined) {
          outRow.push('#N/A');
          continue;
        }
        try {
          outRow.push(this.applyOperator(operator, a, b));
        } catch (e) {
          const msg = (e as Error).message;
          outRow.push(msg.startsWith('#') ? msg.split(' ')[0] : '#VALUE!');
        }
      }
      out.push(outRow);
    }
    return out;
  }

  evaluate(ast: ParsedFormulaNode, ctx: EvaluationContext, currentRow: number = 0, currentCol: number = 0): unknown {
    return this.evaluateNode(ast, ctx, currentRow, currentCol);
  }

  private evaluateNode(
    node: ParsedFormulaNode,
    ctx: EvaluationContext,
    currentRow: number,
    currentCol: number
  ): unknown {
    const env: RefEnv = {
      ctx,
      currentRow,
      currentCol,
      evaluate: (n) => this.evaluateArgument(n, ctx, currentRow, currentCol),
    };
    switch (node.type) {
      case 'number':
      case 'string':
      case 'error':
        return node.value;
      case 'array':
        // A copy, so a function that edits its argument cannot change the formula.
        return (node.value as unknown[][]).map((row) => row.slice());
      case 'cell': {
        const ref = node.cellRef as CellReference;
        const row = ref.rowAbsolute ? ref.row : currentRow + ref.row;
        const col = ref.colAbsolute ? ref.col : currentCol + ref.col;
        // Use sheet name if provided, otherwise use sheetId from context
        return ctx.getCellValue(row, col, undefined, ref.sheetName);
      }
      case 'range': {
        const range = node.rangeRef as NonNullable<ParsedFormulaNode['rangeRef']>;
        return ctx.getRangeValues(range, undefined, range.start.sheetName || range.end.sheetName);
      }
      case 'percent': {
        const value = this.evaluateNode(node.left as ParsedFormulaNode, ctx, currentRow, currentCol);
        return Array.isArray(value)
          ? this.broadcastOperator('/', value, 100)
          : this.applyOperator('/', value, 100);
      }
      case 'implicit':
        return this.implicitIntersection(node.left as ParsedFormulaNode, env, ctx, currentRow, currentCol);
      case 'spillref': {
        const ref = node.cellRef as CellReference;
        const row = ref.rowAbsolute ? ref.row : currentRow + ref.row;
        const col = ref.colAbsolute ? ref.col : currentCol + ref.col;
        const region = ctx.getSpillRegion?.(row, col, ref.sheetName);
        if (!region) throw new Error('#REF!'); // not a spill anchor
        return region.values.map((r) => r.slice());
      }
      case 'variable':
        throw new Error(`#NAME? Name ${node.value} not found`);
      case 'ref3d':
        // A 3D reference is valid only as a function argument (SUM(Sheet1:Sheet3!A1)).
        throw new Error('#VALUE!');
      case 'operator': {
        const operator = node.operator as string;
        if (operator === ':' || operator === ' ') {
          const ref = resolveReference(node, env);
          return readReference(ref as NonNullable<typeof ref>, ctx);
        }
        // A union is valid only as a function argument (SUM((A1:A2,C1))).
        if (operator === ',') throw new Error('#VALUE!');
        const leftVal = this.evaluateNode(node.left as ParsedFormulaNode, ctx, currentRow, currentCol);
        const rightVal = this.evaluateNode(node.right as ParsedFormulaNode, ctx, currentRow, currentCol);

        // Operators work element-wise when an operand is an array (a range or
        // an array result), broadcasting a scalar or a single row/column.
        if (Array.isArray(leftVal) || Array.isArray(rightVal)) {
          return this.broadcastOperator(operator, leftVal, rightVal);
        }
        return this.applyOperator(operator, leftVal, rightVal);
      }
      case 'function': {
        const name = node.functionName as string;
        const args = node.args as ParsedFormulaNode[];
        // Reference functions (ROW, OFFSET, INDIRECT, INDEX, …) need the raw
        // argument AST so they can read an argument's reference.
        const refFunc = this.refFunctions.get(name);
        if (refFunc) {
          return refFunc({ args, ctx, currentRow, currentCol, evaluate: env.evaluate });
        }
        // Lazy functions (IF, IFERROR, IFS, …) receive thunks so they can
        // short-circuit and catch errors thrown while evaluating a branch.
        const lazyFunc = this.lazyFunctions.get(name);
        if (lazyFunc) {
          const thunks = args.map((arg) => () => this.evaluateNode(arg, ctx, currentRow, currentCol));
          return lazyFunc(thunks, ctx);
        }
        const func = this.functions.get(name);
        if (!func) {
          throw new Error(`#NAME? Function ${name} not found`);
        }
        return func(
          args.map((arg) => this.evaluateArgument(arg, ctx, currentRow, currentCol)),
          ctx
        );
      }
      default:
        return null;
    }
  }

  /**
   * Evaluate a function argument. Unlike an operand, an argument may be a
   * multi-area reference: a 3D reference (`Sheet1:Sheet3!A1`) or a union
   * (`(A1:A2,C1)`). Both evaluate to the stacked values of their areas, which is
   * what the aggregate functions (SUM, COUNT, AVERAGE, …) consume.
   */
  private evaluateArgument(
    node: ParsedFormulaNode,
    ctx: EvaluationContext,
    currentRow: number,
    currentCol: number
  ): unknown {
    if (node.type === 'ref3d') return this.evaluate3d(node, ctx, currentRow, currentCol);
    if (node.type === 'operator' && node.operator === ',') {
      const areas: unknown[] = [];
      const collect = (n: ParsedFormulaNode): void => {
        if (n.type === 'operator' && n.operator === ',') {
          collect(n.left as ParsedFormulaNode);
          collect(n.right as ParsedFormulaNode);
        } else {
          areas.push(this.evaluateNode(n, ctx, currentRow, currentCol));
        }
      };
      collect(node);
      return this.stack(areas);
    }
    return this.evaluateNode(node, ctx, currentRow, currentCol);
  }

  /** Stack values (scalars and matrices) on top of each other into one matrix. */
  private stack(values: unknown[]): unknown[][] {
    const rows: unknown[][] = [];
    for (const v of values) {
      if (Array.isArray(v)) {
        for (const row of v as unknown[]) rows.push(Array.isArray(row) ? row : [row]);
      } else {
        rows.push([v]);
      }
    }
    return rows;
  }

  /** The values of a 3D reference: the same cells on every sheet from the first to the last, in workbook order. */
  private evaluate3d(
    node: ParsedFormulaNode,
    ctx: EvaluationContext,
    currentRow: number,
    currentCol: number
  ): unknown[][] {
    const names = ctx.getSheetNames?.() ?? [];
    const first = (node.cellRef ?? (node.rangeRef as NonNullable<ParsedFormulaNode['rangeRef']>).start).sheetName;
    const a = names.indexOf(first as string);
    const b = names.indexOf(node.sheetEnd as string);
    if (a < 0 || b < 0) throw new Error('#REF!'); // an endpoint sheet no longer exists
    const areas: unknown[] = [];
    for (let i = Math.min(a, b); i <= Math.max(a, b); i++) {
      if (node.cellRef) {
        const ref = node.cellRef;
        const row = ref.rowAbsolute ? ref.row : currentRow + ref.row;
        const col = ref.colAbsolute ? ref.col : currentCol + ref.col;
        areas.push(ctx.getCellValue(row, col, undefined, names[i]));
      } else {
        areas.push(ctx.getRangeValues(node.rangeRef as NonNullable<ParsedFormulaNode['rangeRef']>, undefined, names[i]));
      }
    }
    return this.stack(areas);
  }

  /**
   * `@x`: the single value of `x` at the formula's own row/column (a range) or
   * its top-left value (any other array). A scalar passes through.
   */
  private implicitIntersection(
    operand: ParsedFormulaNode,
    env: RefEnv,
    ctx: EvaluationContext,
    currentRow: number,
    currentCol: number
  ): unknown {
    const ref = resolveReference(operand, env);
    if (ref) {
      const row = ref.r1 === ref.r2 ? ref.r1 : currentRow;
      const col = ref.c1 === ref.c2 ? ref.c1 : currentCol;
      if (row < ref.r1 || row > ref.r2 || col < ref.c1 || col > ref.c2) throw new Error('#VALUE!');
      return ctx.getCellValue(row, col, undefined, ref.sheetName);
    }
    const value = this.evaluateNode(operand, ctx, currentRow, currentCol);
    if (!Array.isArray(value)) return value;
    const first = value[0];
    return (Array.isArray(first) ? first[0] : first) ?? null;
  }
}
