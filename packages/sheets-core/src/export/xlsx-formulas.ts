// What the XLSX adapter needs to know about a formula: whether the engine can evaluate it, and an array formula's extent.
// The file's text goes in as the editor form (future-function prefixes already removed, see `fromFileFormula` in the io package).

import { FormulaParser } from '../formula-parser/parser';
import type { ParsedFormulaNode } from '../formula-parser/types';
import { getRegisteredFunctionNames } from '../formula-parser/functions';
import { columnIndexToLabel, columnLabelToIndex } from '../formula-parser/cell-reference';
import { tokenize, type Token } from '../formula-parser/tokenizer';

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

/**
 * A formula a file saved without the array or dynamic-array mark is a legacy one: where it uses a range of several cells as an
 * operand of its own operators (`=A1:A3*2`, `=A1:A3`), Excel takes the one cell of the range in the formula's row (or column), or
 * `#VALUE!` when there is none. Excel 365 shows that with `@` (`=@A1:A3*2`), and so does this: the `@` is put before each such range.
 *
 * Only a range outside every parenthesis and call is changed. What a function does with a range argument depends on the function
 * (SUM takes the whole range, ROUND the cell in its row), so a call's arguments, a group and an array constant are left as they are,
 * and so is a formula with a union, an intersection (space), a spill reference or anything the tokenizer rejects.
 */
export function withImplicitIntersection(formula: string): string {
  if (!formula.includes(':')) return formula; // every range has a colon: most formulas need no tokenizing
  let tokens: Token[];
  try {
    tokens = tokenize(formula);
  } catch {
    return formula;
  }
  const inserts: number[] = [];
  let depth = 0;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.kind === 'op') {
      if (t.text === '(' || t.text === '{') depth++;
      else if (t.text === ')' || t.text === '}') depth--;
      else if (t.text === '#' || (t.text === ',' && depth === 0)) return formula;
      continue;
    }
    if (i > 0 && t.space && tokens[i - 1].kind !== 'op') return formula;
    if (depth !== 0) continue;
    // an optional sheet prefix, then a range of more than one cell: A1:B2 is `cell : cell`, A:A and 1:1 are one token
    const start = t.kind === 'sheet' ? i + 1 : i;
    if (t.kind === 'sheet' && t.end !== undefined) {
      i++; // a 3D reference is one operand: the cell after its prefix is not a range of its own
      continue;
    }
    const first = tokens[start];
    let length = 0;
    if (first?.kind === 'colrange' || first?.kind === 'rowrange') length = 1;
    else if (first?.kind === 'cell') {
      const colon = tokens[start + 1];
      const last = tokens[start + 2];
      if (colon?.kind === 'op' && colon.text === ':' && last?.kind === 'cell') {
        if (last.row !== first.row || last.col !== first.col) length = 3;
      }
    }
    if (length === 0) continue;
    const before = tokens[i - 1];
    const after = tokens[start + length];
    if ((before?.kind === 'op' && (before.text === ':' || before.text === '@')) || (after?.kind === 'op' && after.text === ':')) return formula;
    inserts.push(t.pos);
    i = start + length - 1;
  }
  let out = formula;
  for (const pos of inserts.reverse()) out = `${out.slice(0, pos)}@${out.slice(pos)}`;
  return out;
}
