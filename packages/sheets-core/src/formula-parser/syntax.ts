// Formula syntax: tokens -> numeric AST.
//
// A precedence-climbing parser with Excel's operator order, loosest to tightest:
//
//   comparison  =  <>  <  >  <=  >=
//   concat      &
//   additive    +  -
//   multiplicative  *  /
//   power       ^                (left-associative: 2^3^2 is 64)
//   percent     %                (postfix)
//   sign        -x  +x           (binds tighter than ^: -2^2 is 4)
//   intersect   (space)          A1:C3 B2:D4
//   range       :                A1:INDEX(A1:A5,3)
//
// The union operator `,` exists only inside parentheses: `(A1:A2,C1)`.
// References keep the parser's convention: absolute components hold absolute
// positions, relative ones hold offsets from the formula's own cell.

import { syntaxError, tokenize, type CellToken, type Token } from './tokenizer';
import { MAX_SHEET_COLS, MAX_SHEET_ROWS } from './types';
import type { CellReference, ParsedFormulaNode } from './types';

const MAX_NESTING = 256;
const COMPARISON = new Set(['=', '<>', '<', '>', '<=', '>=']);

/** True for a node that can denote a reference (so it may be an operand of `:`, ` ` and `,`). */
function isReferenceLike(node: ParsedFormulaNode): boolean {
  switch (node.type) {
    case 'cell':
    case 'range':
    case 'ref3d':
    case 'spillref':
    case 'variable':
    case 'function':
    case 'error': // a deleted reference prints as #REF!, which must still parse as an operand
      return true;
    case 'operator':
      return node.operator === ':' || node.operator === ' ' || node.operator === ',';
    default:
      return false;
  }
}

class SyntaxParser {
  private tokens: Token[];
  private index = 0;
  private depth = 0;

  constructor(
    private readonly text: string,
    private readonly currentRow: number,
    private readonly currentCol: number
  ) {
    this.tokens = tokenize(this.text);
  }

  parse(): ParsedFormulaNode {
    if (this.peek().kind === 'eof') return { type: 'string', value: '' };
    const node = this.parseComparison();
    const next = this.peek();
    if (next.kind !== 'eof') throw this.unexpected(next);
    return node;
  }

  // ---- token helpers ------------------------------------------------------

  private peek(): Token {
    return this.tokens[this.index];
  }

  private next(): Token {
    return this.tokens[this.index++];
  }

  private peekOp(): string | undefined {
    const t = this.peek();
    return t.kind === 'op' ? t.text : undefined;
  }

  private expectOp(text: string): void {
    const t = this.next();
    if (t.kind !== 'op' || t.text !== text) throw this.unexpected(t, `expected "${text}"`);
  }

  private unexpected(t: Token, hint?: string): Error {
    const what = t.kind === 'eof' ? 'end of formula' : `${this.describe(t)} at position ${t.pos}`;
    return syntaxError(`Unexpected ${what}${hint ? ` (${hint})` : ''}`);
  }

  /** The token as written, quoted (a string token already carries its own quotes). */
  private describe(t: Token): string {
    const next = this.tokens[this.tokens.indexOf(t) + 1];
    const raw = this.text.slice(t.pos, next.pos).trimEnd();
    return t.kind === 'str' ? raw : `"${raw}"`;
  }

  private enter(): void {
    if (++this.depth > MAX_NESTING) throw syntaxError('Formula is nested too deeply');
  }

  private leave(): void {
    this.depth--;
  }

  // ---- operator levels ----------------------------------------------------

  private binary(operator: string, left: ParsedFormulaNode, right: ParsedFormulaNode): ParsedFormulaNode {
    return { type: 'operator', operator, left, right };
  }

  private parseComparison(): ParsedFormulaNode {
    let left = this.parseConcat();
    for (let op = this.peekOp(); op !== undefined && COMPARISON.has(op); op = this.peekOp()) {
      this.next();
      left = this.binary(op, left, this.parseConcat());
    }
    return left;
  }

  private parseConcat(): ParsedFormulaNode {
    let left = this.parseAdditive();
    while (this.peekOp() === '&') {
      this.next();
      left = this.binary('&', left, this.parseAdditive());
    }
    return left;
  }

  private parseAdditive(): ParsedFormulaNode {
    let left = this.parseMultiplicative();
    for (let op = this.peekOp(); op === '+' || op === '-'; op = this.peekOp()) {
      this.next();
      left = this.binary(op, left, this.parseMultiplicative());
    }
    return left;
  }

  private parseMultiplicative(): ParsedFormulaNode {
    let left = this.parsePower();
    for (let op = this.peekOp(); op === '*' || op === '/'; op = this.peekOp()) {
      this.next();
      left = this.binary(op, left, this.parsePower());
    }
    return left;
  }

  private parsePower(): ParsedFormulaNode {
    let left = this.parsePercent();
    while (this.peekOp() === '^') {
      this.next();
      left = this.binary('^', left, this.parsePercent());
    }
    return left;
  }

  private parsePercent(): ParsedFormulaNode {
    let node = this.parseSign();
    while (this.peekOp() === '%') {
      this.next();
      node = { type: 'percent', left: node };
    }
    return node;
  }

  /** A leading sign is stored as `"" op x`, which evaluates as 0 op x. */
  private parseSign(): ParsedFormulaNode {
    const op = this.peekOp();
    if (op === '-' || op === '+') {
      this.next();
      this.enter();
      const operand = this.parseSign();
      this.leave();
      return this.binary(op, { type: 'string', value: '' }, operand);
    }
    return this.parseIntersection();
  }

  private startsOperand(t: Token): boolean {
    switch (t.kind) {
      case 'cell':
      case 'colrange':
      case 'rowrange':
      case 'sheet':
      case 'ident':
      case 'err':
        return true;
      case 'op':
        return t.text === '(';
      default:
        return false;
    }
  }

  /** A space between two operands intersects them. */
  private parseIntersection(): ParsedFormulaNode {
    let left = this.parseRange();
    for (let t = this.peek(); t.space && this.startsOperand(t); t = this.peek()) {
      const right = this.parseRange();
      if (!isReferenceLike(left) || !isReferenceLike(right)) {
        throw syntaxError(`The intersection operator needs references on both sides (position ${t.pos})`);
      }
      left = this.binary(' ', left, right);
    }
    return left;
  }

  /** `:` joins two references; two plain cells make a range node, anything else an operator node. */
  private parseRange(): ParsedFormulaNode {
    let left = this.parsePrimary();
    while (this.peekOp() === ':') {
      const colon = this.next();
      const right = this.parsePrimary();
      left = this.joinRange(left, right, colon.pos);
    }
    return left;
  }

  private joinRange(left: ParsedFormulaNode, right: ParsedFormulaNode, pos: number): ParsedFormulaNode {
    if (left.cellRef && right.cellRef && right.type === 'cell') {
      const sheet = left.cellRef.sheetName;
      const rightSheet = right.cellRef.sheetName;
      if (left.type === 'cell' && (rightSheet === undefined || rightSheet === sheet)) {
        return {
          type: 'range',
          rangeRef: {
            start: { ...left.cellRef },
            end: { ...right.cellRef, sheetName: sheet },
          },
        };
      }
      if (left.type === 'ref3d' && rightSheet === undefined) {
        return {
          type: 'ref3d',
          sheetEnd: left.sheetEnd,
          rangeRef: {
            start: left.cellRef,
            end: { ...right.cellRef, sheetName: left.cellRef.sheetName },
          },
        };
      }
    }
    if (!isReferenceLike(left) || !isReferenceLike(right)) {
      throw syntaxError(`The range operator ":" needs references on both sides (position ${pos})`);
    }
    return this.binary(':', left, right);
  }

  // ---- operands -----------------------------------------------------------

  private parsePrimary(): ParsedFormulaNode {
    const t = this.next();
    switch (t.kind) {
      case 'num':
        return { type: 'number', value: t.value };
      case 'str':
        return { type: 'string', value: t.value };
      case 'err':
        return { type: 'error', value: t.value };
      case 'cell': {
        const node: ParsedFormulaNode = { type: 'cell', cellRef: this.cellRef(t) };
        return this.peekOp() === '#' && !this.peek().space ? this.spill(node) : node;
      }
      case 'colrange':
        return this.wholeColumns(t);
      case 'rowrange':
        return this.wholeRows(t);
      case 'sheet':
        return this.sheetReference(t);
      case 'ident':
        return this.identifier(t);
      case 'op':
        return this.operatorOperand(t);
      default:
        throw this.unexpected(t);
    }
  }

  private operatorOperand(t: Token & { kind: 'op' }): ParsedFormulaNode {
    switch (t.text) {
      case '(':
        return this.group();
      case '{':
        return this.arrayConstant();
      case '@': {
        this.enter();
        const operand = this.parseRange();
        this.leave();
        return { type: 'implicit', left: operand };
      }
      default:
        throw this.unexpected(t);
    }
  }

  /** `( expr )`, or a union `( ref , ref ... )`. */
  private group(): ParsedFormulaNode {
    this.enter();
    let node = this.parseComparison();
    while (this.peekOp() === ',') {
      const comma = this.next();
      const right = this.parseComparison();
      if (!isReferenceLike(node) || !isReferenceLike(right)) {
        throw syntaxError(`The union operator "," needs references on both sides (position ${comma.pos})`);
      }
      node = this.binary(',', node, right);
    }
    this.expectOp(')');
    this.leave();
    return node;
  }

  private spill(anchor: ParsedFormulaNode): ParsedFormulaNode {
    this.next(); // '#'
    return { type: 'spillref', cellRef: anchor.cellRef };
  }

  private identifier(t: Token & { kind: 'ident' }): ParsedFormulaNode {
    if (this.peekOp() === '(') return this.functionCall(t.text);
    const upper = t.text.toUpperCase();
    // TRUE / FALSE route through the registered TRUE()/FALSE() so AND/OR and
    // comparisons see a real boolean.
    if (upper === 'TRUE' || upper === 'FALSE') return { type: 'function', functionName: upper, args: [] };
    // A defined name. Names are not resolved yet, so it evaluates to #NAME?, but
    // it is kept as written so the formula text survives (a round trip, an export).
    return { type: 'variable', value: t.text };
  }

  private functionCall(name: string): ParsedFormulaNode {
    this.expectOp('(');
    this.enter();
    const args: ParsedFormulaNode[] = [];
    if (this.peekOp() !== ')') {
      for (;;) {
        // An empty argument (`F(1,,2)`) is skipped, as the engine always has.
        const op = this.peekOp();
        if (op !== ',' && op !== ')') args.push(this.parseComparison());
        const sep = this.next();
        if (sep.kind === 'op' && sep.text === ')') break;
        if (sep.kind !== 'op' || sep.text !== ',') throw this.unexpected(sep, 'expected "," or ")"');
      }
    } else {
      this.next();
    }
    this.leave();
    return { type: 'function', functionName: name.toUpperCase(), args };
  }

  // ---- references ---------------------------------------------------------

  private cellRef(t: CellToken, sheetName?: string): CellReference {
    return {
      row: t.rowAbs ? t.row : t.row - this.currentRow,
      col: t.colAbs ? t.col : t.col - this.currentCol,
      rowAbsolute: t.rowAbs,
      colAbsolute: t.colAbs,
      sheetName,
    };
  }

  private wholeColumns(t: Token & { kind: 'colrange' }, sheetName?: string): ParsedFormulaNode {
    const col = (c: number, abs: boolean) => (abs ? c : c - this.currentCol);
    return {
      type: 'range',
      rangeRef: {
        start: { row: 0, col: col(t.c1, t.abs1), rowAbsolute: true, colAbsolute: t.abs1, sheetName },
        end: { row: MAX_SHEET_ROWS - 1, col: col(t.c2, t.abs2), rowAbsolute: true, colAbsolute: t.abs2, sheetName },
        whole: 'column',
      },
    };
  }

  private wholeRows(t: Token & { kind: 'rowrange' }, sheetName?: string): ParsedFormulaNode {
    const row = (r: number, abs: boolean) => (abs ? r : r - this.currentRow);
    return {
      type: 'range',
      rangeRef: {
        start: { row: row(t.r1, t.abs1), col: 0, rowAbsolute: t.abs1, colAbsolute: true, sheetName },
        end: { row: row(t.r2, t.abs2), col: MAX_SHEET_COLS - 1, rowAbsolute: t.abs2, colAbsolute: true, sheetName },
        whole: 'row',
      },
    };
  }

  /** After a sheet prefix (`Sheet1!`, `'My Sheet'!`, `Sheet1:Sheet3!`) comes a reference. */
  private sheetReference(sheet: Token & { kind: 'sheet' }): ParsedFormulaNode {
    const t = this.next();
    if (t.space) throw this.unexpected(t, `a reference must follow "${sheet.start}!"`);
    let node: ParsedFormulaNode;
    switch (t.kind) {
      case 'cell':
        node = { type: 'cell', cellRef: this.cellRef(t, sheet.start) };
        break;
      case 'colrange':
        node = this.wholeColumns(t, sheet.start);
        break;
      case 'rowrange':
        node = this.wholeRows(t, sheet.start);
        break;
      case 'err':
        return { type: 'error', value: t.value };
      default:
        throw this.unexpected(t, `a reference must follow "${sheet.start}!"`);
    }
    if (sheet.end === undefined) {
      return node.type === 'cell' && this.peekOp() === '#' && !this.peek().space ? this.spill(node) : node;
    }
    // 3D reference: the first sheet is the ref's sheet, the last one is kept beside it.
    return { ...node, type: 'ref3d', sheetEnd: sheet.end };
  }

  // ---- array constants ----------------------------------------------------

  private arrayConstant(): ParsedFormulaNode {
    const open = this.tokens[this.index - 1];
    const rows: unknown[][] = [[]];
    for (;;) {
      rows[rows.length - 1].push(this.arrayElement());
      const sep = this.next();
      if (sep.kind === 'op' && sep.text === '}') break;
      if (sep.kind === 'op' && sep.text === ',') continue;
      if (sep.kind === 'op' && sep.text === ';') {
        rows.push([]);
        continue;
      }
      throw this.unexpected(sep, 'expected ",", ";" or "}" in the array constant');
    }
    const width = rows[0].length;
    const ragged = rows.findIndex((r) => r.length !== width);
    if (ragged >= 0) {
      throw syntaxError(
        `Array constant at position ${open.pos} is ragged: row ${ragged + 1} has ${rows[ragged].length} elements, row 1 has ${width}`
      );
    }
    return { type: 'array', value: rows };
  }

  /** A number (optionally signed), a string, TRUE/FALSE or an error literal. */
  private arrayElement(): unknown {
    let t = this.next();
    let sign = 1;
    if (t.kind === 'op' && (t.text === '-' || t.text === '+')) {
      sign = t.text === '-' ? -1 : 1;
      const signed = this.next();
      if (signed.kind !== 'num') throw this.unexpected(signed, 'a sign must be followed by a number in an array constant');
      t = signed;
    }
    switch (t.kind) {
      case 'num':
        return sign * t.value;
      case 'str':
        return t.value;
      case 'err':
        return t.value;
      case 'ident': {
        const upper = t.text.toUpperCase();
        if (upper === 'TRUE') return true;
        if (upper === 'FALSE') return false;
        throw this.unexpected(t, 'array constants hold only numbers, text, TRUE/FALSE and errors');
      }
      default:
        throw this.unexpected(t, 'array constants hold only numbers, text, TRUE/FALSE and errors');
    }
  }
}

/** Parse formula text (without the leading `=`). Throws an `#ERROR! ...` Error on a syntax error. */
export function parseFormulaSyntax(text: string, currentRow: number, currentCol: number): ParsedFormulaNode {
  return new SyntaxParser(text, currentRow, currentCol).parse();
}
