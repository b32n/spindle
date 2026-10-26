// Formula parser - simplified but functional

import type { ParsedFormulaNode, ParseResult, EvaluationContext, RangeReference } from './types';
import { parseCellReference, parseRangeReference, cellReferenceToKey } from './cell-reference';
import { eagerFunctions, lazyFunctions, refFunctions } from './functions';
import type { EagerFn, LazyFn, RefFn } from './functions';
import { compareValues, excelEqual, toText } from './functions/helpers';

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
      // Guard against a pathologically long formula. parseExpression is
      // recursive and roughly O(n^2) on a flat expression, so a huge input
      // freezes the main thread; reject it instead.
      if (formula.length > 10_000) {
        throw new Error('#ERROR! Formula too long');
      }

      // Remove leading = if present
      const expression = formula.startsWith('=') ? formula.slice(1) : formula;

      // Simple parser - handles basic cases
      const ast = this.parseExpression(expression, currentRow, currentCol, dependencies);

      return { ast, dependencies, error };
    } catch (e) {
      error = e instanceof Error ? e.message : '#ERROR!';
      return {
        ast: { type: 'string', value: '' },
        dependencies,
        error,
      };
    }
  }

  private parseExpression(
    expr: string,
    currentRow: number,
    currentCol: number,
    dependencies: Set<string>
  ): ParsedFormulaNode {
    expr = expr.trim();

    // Unwrap a fully-enclosing pair of parentheses: (a+b) -> a+b. Grouping
    // parens are otherwise never stripped, so =(1+2), 2*(3+4) and =(1>0)
    // mis-parse (as a string or #VALUE!). The loop collapses nested wraps.
    while (expr.length >= 2 && expr[0] === '(' && this.isFullyParenthesized(expr)) {
      expr = expr.slice(1, -1).trim();
    }

    // Try to parse as arithmetic/comparison expression
    const operatorMatch = this.findLowestPrecedenceOperator(expr);
    if (operatorMatch) {
      const { operator, index } = operatorMatch;
      const leftExpr = expr.substring(0, index).trim();
      const rightExpr = expr.substring(index + operator.length).trim();
      
      return {
        type: 'operator',
        operator,
        left: this.parseExpression(leftExpr, currentRow, currentCol, dependencies),
        right: this.parseExpression(rightExpr, currentRow, currentCol, dependencies),
      };
    }

    // Try to parse as cell reference
    const cellRef = parseCellReference(expr);
    if (cellRef) {
      // For relative references, convert absolute position to offset
      if (!cellRef.rowAbsolute || !cellRef.colAbsolute) {
        // Calculate offset from current cell
        const rowOffset = cellRef.row - currentRow;
        const colOffset = cellRef.col - currentCol;
        // Store as offset for relative references
        const adjustedRef = {
          row: cellRef.rowAbsolute ? cellRef.row : rowOffset,
          col: cellRef.colAbsolute ? cellRef.col : colOffset,
          rowAbsolute: cellRef.rowAbsolute,
          colAbsolute: cellRef.colAbsolute,
          sheetName: cellRef.sheetName, // Preserve sheet name
        };
        const key = cellReferenceToKey(adjustedRef, currentRow, currentCol);
        // For cross-sheet references, include sheet name in dependency key
        const depKey = cellRef.sheetName ? `${cellRef.sheetName}!${key}` : key;
        dependencies.add(depKey);
        return { type: 'cell', cellRef: adjustedRef };
      }
      // Absolute reference - use as-is
      const key = cellReferenceToKey(cellRef, currentRow, currentCol);
      // For cross-sheet references, include sheet name in dependency key
      const depKey = cellRef.sheetName ? `${cellRef.sheetName}!${key}` : key;
      dependencies.add(depKey);
      return { type: 'cell', cellRef };
    }

    // Try to parse as range
    const rangeRef = parseRangeReference(expr);
    if (rangeRef) {
      // Add all cells in range to dependencies
      const { start, end } = rangeRef;
      
      // For relative references, convert absolute position to offset (like we do for single cells)
      // parseCellReference returns absolute positions, so we need to calculate the offset
      const startRowOffset = start.rowAbsolute ? 0 : start.row - currentRow;
      const startColOffset = start.colAbsolute ? 0 : start.col - currentCol;
      const endRowOffset = end.rowAbsolute ? 0 : end.row - currentRow;
      const endColOffset = end.colAbsolute ? 0 : end.col - currentCol;
      
      // Create adjusted range reference with offsets for relative refs
      const sheetName = start.sheetName || end.sheetName; // Get sheet name from either start or end
      const adjustedRangeRef: RangeReference = {
        start: {
          row: start.rowAbsolute ? start.row : startRowOffset,
          col: start.colAbsolute ? start.col : startColOffset,
          rowAbsolute: start.rowAbsolute,
          colAbsolute: start.colAbsolute,
          sheetName: sheetName, // Preserve sheet name
        },
        end: {
          row: end.rowAbsolute ? end.row : endRowOffset,
          col: end.colAbsolute ? end.col : endColOffset,
          rowAbsolute: end.rowAbsolute,
          colAbsolute: end.colAbsolute,
          sheetName: sheetName, // Preserve sheet name
        },
      };
      
      // Calculate absolute positions for dependency tracking
      const startRow = adjustedRangeRef.start.rowAbsolute 
        ? adjustedRangeRef.start.row 
        : currentRow + adjustedRangeRef.start.row;
      const endRow = adjustedRangeRef.end.rowAbsolute 
        ? adjustedRangeRef.end.row 
        : currentRow + adjustedRangeRef.end.row;
      const startCol = adjustedRangeRef.start.colAbsolute 
        ? adjustedRangeRef.start.col 
        : currentCol + adjustedRangeRef.start.col;
      const endCol = adjustedRangeRef.end.colAbsolute 
        ? adjustedRangeRef.end.col 
        : currentCol + adjustedRangeRef.end.col;

      // Add dependencies with sheet name prefix if it's a cross-sheet reference.
      // Skip the per-cell expansion for an oversized range (e.g. A1:A1048576):
      // it would iterate billions of cells and hang the main thread. The recalc
      // graph derives its real dependencies from the AST
      // (collectStableDependencies), so this Set is only a best-effort extra.
      const sheetPrefix = sheetName ? `${sheetName}!` : '';
      const rLo = Math.min(startRow, endRow);
      const rHi = Math.max(startRow, endRow);
      const cLo = Math.min(startCol, endCol);
      const cHi = Math.max(startCol, endCol);
      const MAX_RANGE_DEP_CELLS = 100_000;
      if ((rHi - rLo + 1) * (cHi - cLo + 1) <= MAX_RANGE_DEP_CELLS) {
        for (let r = rLo; r <= rHi; r++) {
          for (let c = cLo; c <= cHi; c++) {
            dependencies.add(`${sheetPrefix}${r}:${c}`);
          }
        }
      }

      return { type: 'range', rangeRef: adjustedRangeRef };
    }

    // Try to parse as function call (e.g., SUM(A1:A10) or ERROR.TYPE(A1)).
    // Case-insensitive: `sum(...)` === `SUM(...)`; the name is upper-cased for
    // lookup below.
    const functionMatch = expr.match(/^([A-Za-z][A-Za-z.]*)\s*\((.*)\)$/);
    if (functionMatch) {
      const [, funcName, argsStr] = functionMatch;
      const args = this.parseArguments(argsStr, currentRow, currentCol, dependencies);
      return {
        type: 'function',
        functionName: funcName.toUpperCase(),
        args,
      };
    }

    // Boolean literals: the barewords TRUE / FALSE (Excel returns booleans,
    // not the strings). Route through the registered TRUE()/FALSE() functions
    // so AND/OR/IF and comparisons see a real boolean.
    const upperExpr = expr.toUpperCase();
    if (upperExpr === 'TRUE') return { type: 'function', functionName: 'TRUE', args: [] };
    if (upperExpr === 'FALSE') return { type: 'function', functionName: 'FALSE', args: [] };

    // Try to parse as number
    const num = Number(expr);
    if (!isNaN(num) && expr.trim() !== '') {
      return { type: 'number', value: num };
    }

    // Try to parse as string (quoted)
    if ((expr.startsWith('"') && expr.endsWith('"')) || (expr.startsWith("'") && expr.endsWith("'"))) {
      return { type: 'string', value: expr.slice(1, -1) };
    }

    // Default to string
    return { type: 'string', value: expr };
  }

  /**
   * True when the expression is wrapped in one pair of parentheses that spans
   * the whole thing — `(1+2)` is, but `(1+2)*(3+4)` is not (the first group
   * closes before the end). Quoted strings are skipped so a `)` inside a
   * string literal doesn't confuse the depth count.
   */
  private isFullyParenthesized(expr: string): boolean {
    if (expr[0] !== '(' || expr[expr.length - 1] !== ')') return false;
    let depth = 0;
    for (let i = 0; i < expr.length; i++) {
      const c = expr[i];
      if (c === '"' || c === "'") {
        const quote = c;
        i++;
        while (i < expr.length && expr[i] !== quote) {
          if (expr[i] === '\\') i++;
          i++;
        }
        continue;
      }
      if (c === '(') depth++;
      else if (c === ')') {
        depth--;
        // The opening paren's match closed before the end → not fully wrapping.
        if (depth === 0 && i !== expr.length - 1) return false;
      }
    }
    return depth === 0;
  }

  /**
   * Find the operator with lowest precedence (to parse correctly)
   * Returns the rightmost operator of lowest precedence to handle left-associativity
   */
  private findLowestPrecedenceOperator(expr: string): { operator: string; index: number } | null {
    let depth = 0;
    let lowestPrecOp: { operator: string; index: number } | null = null;
    let lowestPrec = Infinity;

    // Operator precedence (lower number = lower precedence, evaluated later),
    // matching Excel: comparison < & (concat) < +,- < *,/ < ^ (power).
    const operators = [
      { op: '>=', prec: 0 },
      { op: '<=', prec: 0 },
      { op: '<>', prec: 0 },
      { op: '>', prec: 0 },
      { op: '<', prec: 0 },
      { op: '=', prec: 0 },
      { op: '&', prec: 1 },
      { op: '+', prec: 2 },
      { op: '-', prec: 2 },
      { op: '*', prec: 3 },
      { op: '/', prec: 3 },
      { op: '^', prec: 4 },
    ];

    for (let i = 0; i < expr.length; i++) {
      const char = expr[i];
      
      // Track parentheses depth
      if (char === '(') depth++;
      else if (char === ')') depth--;
      // Skip if inside parentheses or quotes
      else if (depth > 0) continue;
      else if (char === '"' || char === "'") {
        // Skip quoted strings
        const quote = char;
        i++;
        while (i < expr.length && expr[i] !== quote) {
          if (expr[i] === '\\') i++; // Skip escaped characters
          i++;
        }
        continue;
      }
      // Check for operators (but not at start, and not part of a number or cell reference)
      else {
        for (const { op, prec } of operators) {
          // Check for multi-character operators first
          if (op.length === 2 && i < expr.length - 1) {
            const twoChars = expr.substring(i, i + 2);
            if (twoChars === op) {
              if (prec < lowestPrec || (prec === lowestPrec && i > (lowestPrecOp?.index ?? -1))) {
                lowestPrec = prec;
                lowestPrecOp = { operator: op, index: i };
              }
              continue;
            }
          }
          
          // Single character operators
          if (op.length === 1 && char === op) {
            const prevChar = i > 0 ? expr[i - 1] : '';
            const nextChar = i < expr.length - 1 ? expr[i + 1] : '';

            // The +/- in scientific notation (e.g. 1e-5) is part of the number.
            // Require the e/E to actually be a number's exponent (preceded by a
            // digit or '.') so the trailing E of a word like TRUE/FALSE — or a
            // cell/name ending in E — doesn't swallow the following operator
            // (=TRUE+1 must split into TRUE + 1, not parse as one token).
            const charBeforePrev = i > 1 ? expr[i - 2] : '';
            const isPartOfNumber =
              (op === '-' || op === '+') &&
              (prevChar === 'e' || prevChar === 'E') &&
              /[0-9.]/.test(charBeforePrev);

            // A +/- directly following another operator is a unary sign that
            // belongs to the operand on its right (e.g. the -3 in 2*-3), not a
            // binary split point. A leading sign is left alone so it still
            // parses via the empty-left path (0 - x), which also handles -A1.
            let prevNonSpace = '';
            for (let k = i - 1; k >= 0; k--) {
              if (expr[k] !== ' ') {
                prevNonSpace = expr[k];
                break;
              }
            }
            const isUnarySign =
              (op === '-' || op === '+') &&
              prevNonSpace !== '' &&
              '+-*/(<>=&^'.includes(prevNonSpace);

            // Skip a single char that is really one half of a two-char operator
            // — either the first char (<=, >=, <>) or the second (the = in
            // <=/>=, or the > in <>). Without the second-char guard the tie-break
            // toward the rightmost operator makes the lone = win, so <= / >= / <>
            // all mis-parse.
            const isTwoCharOp =
              (op === '<' && (nextChar === '>' || nextChar === '=')) ||
              (op === '>' && nextChar === '=') ||
              (op === '=' && (prevChar === '<' || prevChar === '>')) ||
              (op === '>' && prevChar === '<');

            // A leading sign is a unary minus/plus: in Excel it binds tighter
            // than ^ (=-2^2 is 4), so it must split last. Give it the highest
            // precedence; it still parses via the empty-left path (0 - x).
            const isLeadingSign = (op === '-' || op === '+') && prevNonSpace === '';
            const effPrec = isLeadingSign ? 5 : prec;

            if (
              !isPartOfNumber &&
              !isUnarySign &&
              !isTwoCharOp &&
              (effPrec < lowestPrec || (effPrec === lowestPrec && i > (lowestPrecOp?.index ?? -1)))
            ) {
              lowestPrec = effPrec;
              lowestPrecOp = { operator: op, index: i };
            }
          }
        }
      }
    }

    return lowestPrecOp;
  }

  private parseArguments(
    argsStr: string,
    currentRow: number,
    currentCol: number,
    dependencies: Set<string>
  ): ParsedFormulaNode[] {
    if (!argsStr.trim()) return [];

    const args: ParsedFormulaNode[] = [];
    let depth = 0;
    let current = '';

    for (let i = 0; i < argsStr.length; i++) {
      const char = argsStr[i];
      // Copy quoted strings through verbatim so a comma inside a string
      // literal (e.g. a TEXTJOIN delimiter) is not treated as a separator.
      if (char === '"' || char === "'") {
        const quote = char;
        current += char;
        i++;
        while (i < argsStr.length && argsStr[i] !== quote) {
          current += argsStr[i];
          i++;
        }
        if (i < argsStr.length) current += argsStr[i];
        continue;
      }
      if (char === '(') depth++;
      else if (char === ')') depth--;
      else if (char === ',' && depth === 0) {
        if (current.trim()) {
          args.push(this.parseExpression(current.trim(), currentRow, currentCol, dependencies));
        }
        current = '';
        continue;
      }
      current += char;
    }

    if (current.trim()) {
      args.push(this.parseExpression(current.trim(), currentRow, currentCol, dependencies));
    }

    return args;
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
          const msg = e instanceof Error ? e.message : '';
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
    currentRow: number = 0,
    currentCol: number = 0
  ): unknown {
    switch (node.type) {
      case 'number':
        return node.value;
      case 'string':
        return node.value;
      case 'cell':
        if (node.cellRef) {
          const ref = node.cellRef;
          const row = ref.rowAbsolute ? ref.row : currentRow + ref.row;
          const col = ref.colAbsolute ? ref.col : currentCol + ref.col;
          // Use sheet name if provided, otherwise use sheetId from context
          return ctx.getCellValue(row, col, undefined, ref.sheetName);
        }
        return null;
      case 'range':
        if (node.rangeRef) {
          const sheetName = node.rangeRef.start.sheetName || node.rangeRef.end.sheetName;
          return ctx.getRangeValues(node.rangeRef, undefined, sheetName);
        }
        return null;
      case 'operator':
        if (node.operator && node.left && node.right) {
          const leftVal = this.evaluateNode(node.left, ctx, currentRow, currentCol);
          const rightVal = this.evaluateNode(node.right, ctx, currentRow, currentCol);
          
          // Operators work element-wise when an operand is an array (a range or
          // an array result), broadcasting a scalar or a single row/column.
          if (Array.isArray(leftVal) || Array.isArray(rightVal)) {
            return this.broadcastOperator(node.operator, leftVal, rightVal);
          }
          return this.applyOperator(node.operator, leftVal, rightVal);
        }
        return null;
      case 'function':
        if (node.functionName && node.args) {
          // Reference functions (ROW, OFFSET, INDIRECT, …) need the raw
          // argument AST so they can read an argument's reference.
          const refFunc = this.refFunctions.get(node.functionName);
          if (refFunc) {
            return refFunc({
              args: node.args,
              ctx,
              currentRow,
              currentCol,
              evaluate: (n) => this.evaluateNode(n, ctx, currentRow, currentCol),
            });
          }
          // Lazy functions (IF, IFERROR, IFS, …) receive thunks so they can
          // short-circuit and catch errors thrown while evaluating a branch.
          const lazyFunc = this.lazyFunctions.get(node.functionName);
          if (lazyFunc) {
            const thunks = node.args.map(
              (arg) => () => this.evaluateNode(arg, ctx, currentRow, currentCol)
            );
            return lazyFunc(thunks, ctx);
          }
          const func = this.functions.get(node.functionName);
          if (!func) {
            throw new Error(`#NAME? Function ${node.functionName} not found`);
          }
          const evaluatedArgs = node.args.map((arg) => this.evaluateNode(arg, ctx, currentRow, currentCol));
          return func(evaluatedArgs, ctx);
        }
        return null;
      default:
        return null;
    }
  }
}

