// What the XLSX adapter needs to know about a formula: whether the engine can evaluate it, and an array formula's extent.
// The file's text goes in as the editor form (future-function prefixes already removed, see `fromFileFormula` in the io package).

import { FormulaParser } from '../formula-parser/parser';
import type { ParsedFormulaNode } from '../formula-parser/types';
import { getRegisteredFunctionNames } from '../formula-parser/functions';
import { columnIndexToLabel, columnLabelToIndex } from '../formula-parser/cell-reference';

let parser: FormulaParser | undefined;
let known: ReadonlySet<string> | undefined;

/**
 * True when the engine can evaluate the formula: it parses, every function it calls is implemented and it names no defined name
 * (names are not imported yet). Anything else is kept as the file's saved result (`cachedOnly`) rather than shown as `#NAME?`.
 */
export function isFormulaEvaluable(formula: string): boolean {
  parser ??= new FormulaParser();
  known ??= getRegisteredFunctionNames();
  const parsed = parser.parse(formula);
  if (parsed.error) return false;
  // iterative: a formula nested thousands deep must not need that much stack
  const stack: ParsedFormulaNode[] = [parsed.ast];
  while (stack.length > 0) {
    const node = stack.pop() as ParsedFormulaNode;
    if (node.type === 'variable') return false;
    if (node.type === 'function' && !known.has(node.functionName as string)) return false;
    if (node.args) for (const arg of node.args) stack.push(arg);
    if (node.left) stack.push(node.left);
    if (node.right) stack.push(node.right);
  }
  return true;
}

const RANGE = /^\$?([A-Z]{1,3})\$?([0-9]{1,7})(?::\$?([A-Z]{1,3})\$?([0-9]{1,7}))?$/i;

/** First cell (0-based) and size of an A1 range or cell (`G11:G15` starts at row 10, column 6 and is 5 x 1); undefined when it is not one. */
export function parseRange(ref: string): { row: number; col: number; rows: number; cols: number } | undefined {
  const m = RANGE.exec(ref);
  if (!m) return undefined;
  const r1 = Number(m[2]);
  const c1 = columnLabelToIndex(m[1].toUpperCase());
  const r2 = m[4] === undefined ? r1 : Number(m[4]);
  const c2 = m[3] === undefined ? c1 : columnLabelToIndex(m[3].toUpperCase());
  if (r1 < 1 || r2 < r1 || c2 < c1) return undefined;
  return { row: r1 - 1, col: c1, rows: r2 - r1 + 1, cols: c2 - c1 + 1 };
}

/** Rows and columns of an A1 range or cell (`G11:G15` is 5 x 1); undefined when it is not one. */
export function rangeSpan(ref: string): { rows: number; cols: number } | undefined {
  const range = parseRange(ref);
  return range && { rows: range.rows, cols: range.cols };
}

/** The A1 range of `rows` x `cols` cells whose first cell is at (`row`, `col`), both 0-based (a single cell for 1 x 1). */
export function spanRef(row: number, col: number, rows: number, cols: number): string {
  const first = `${columnIndexToLabel(col)}${row + 1}`;
  return rows === 1 && cols === 1 ? first : `${first}:${columnIndexToLabel(col + cols - 1)}${row + rows}`;
}
