// Formula tokenizer.
//
// Turns formula text (without the leading `=`) into tokens for syntax.ts. A
// reference is recognised here, as one token, because its shape depends on the
// characters around it (`A:A` is a column range, `A1:B2` two cells and an
// operator, `Sheet1:Sheet3!` a 3D sheet prefix). Every token records whether
// whitespace preceded it: a space between two operands is the intersection
// operator.

import { columnLabelToIndex } from './cell-reference';
import { MAX_SHEET_COLS, MAX_SHEET_ROWS } from './types';

interface Positioned {
  /** Offset of the token's first character in the formula text. */
  pos: number;
  /** Whitespace separated this token from the previous one. */
  space: boolean;
}

export interface CellToken extends Positioned {
  kind: 'cell';
  /** 0-based position. */
  row: number;
  col: number;
  rowAbs: boolean;
  colAbs: boolean;
}

export type Token =
  | (Positioned & { kind: 'num'; value: number })
  | (Positioned & { kind: 'str'; value: string })
  | (Positioned & { kind: 'err'; value: string })
  | (Positioned & { kind: 'ident'; text: string })
  | CellToken
  | (Positioned & { kind: 'colrange'; c1: number; c2: number; abs1: boolean; abs2: boolean })
  | (Positioned & { kind: 'rowrange'; r1: number; r2: number; abs1: boolean; abs2: boolean })
  | (Positioned & { kind: 'sheet'; start: string; end?: string })
  | (Positioned & { kind: 'op'; text: string })
  /** Lenient mode only: a stretch the grammar does not know (a structured reference, an unterminated string, a stray character), kept as written. */
  | (Positioned & { kind: 'raw'; end: number })
  | (Positioned & { kind: 'eof' });

export function syntaxError(message: string): Error {
  return new Error(`#ERROR! ${message}`);
}

// An unquoted sheet name: words of letters, digits, `_` and `.` separated by single spaces
// (`Sheet1`, `Q1 Data`, `2024`). Excel itself writes quotes around anything but a plain
// identifier; the looser form is accepted because earlier versions of the engine did.
const NAME_RUN = String.raw`[\p{L}\p{N}_.\\]+(?: [\p{L}\p{N}_.\\]+)*`;
const SHEET_PREFIX = new RegExp(`^(${NAME_RUN})(?::(${NAME_RUN}))?!`, 'u');
const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w.])/;
const CELL_ADDRESS = /^[A-Za-z]{1,3}\d+$/;

/**
 * The sheet prefix (`Sheet1!`, `Q1 Data!`, `Sheet1:Sheet3!`) at the start of `rest`, if there is one.
 * `A1:Sheet2!B2` is a range from A1 to a cell on Sheet2, so a span of sheets named like cell addresses
 * must be quoted. (At the start of an operand `A1 Sheet2!B2` reads as the sheet "A1 Sheet2".)
 */
function sheetPrefixAt(rest: string, prev: Token | undefined): RegExpExecArray | null {
  // No name is longer than MAX_PREFIX_CHARS, so a longer run of name characters is no prefix: bounding the match keeps a
  // formula of many short words (`a a a a ...`) from rescanning the whole run at every word.
  const m = SHEET_PREFIX.exec(rest.length > MAX_PREFIX_CHARS ? rest.slice(0, MAX_PREFIX_CHARS) : rest);
  if (!m) return null;
  const names = m[2] === undefined ? [m[1]] : [m[1], m[2]];
  // A name of several words can only start where an operand starts: after `Sheet1!A1 ` the words are the
  // next operand of an intersection (`Sheet1!A1 Sheet1!B1`), and after `Sheet1!` a reference.
  const operandStarts = prev === undefined || (prev.kind === 'op' && !')%#}:'.includes(prev.text));
  if (!operandStarts && names.some((n) => n.includes(' '))) return null;
  // A span of sheets is `Sheet1:Sheet3!`: no spaces inside it (quote it for that), and no cell addresses.
  if (m[2] !== undefined && names.some((n) => CELL_ADDRESS.test(n) || n.includes(' '))) return null;
  return m;
}
/** Two names of 255 characters (the longest the importer keeps) joined by `:`, and the `!`. */
const MAX_PREFIX_CHARS = 2 * 255 + 2;
const COL_RANGE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w.(])/;
const ROW_RANGE = /^(\$?)(\d+):(\$?)(\d+)(?![\w.])/;
const NUMBER = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/;
const IDENT = /^[\p{L}_\\][\p{L}\p{N}_.]*/u;
const ERROR_LITERAL = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!)/i;

/** Scan a quoted run (string or sheet name) where a doubled quote stands for one literal quote. */
function scanQuoted(text: string, start: number, quote: string): { value: string; end: number } | undefined {
  let value = '';
  let i = start + 1;
  while (i < text.length) {
    if (text[i] === quote) {
      if (text[i + 1] === quote) {
        value += quote; // a doubled quote is a literal quote
        i += 2;
        continue;
      }
      return { value, end: i + 1 };
    }
    value += text[i];
    i++;
  }
  return undefined;
}

function cellFromMatch(m: RegExpExecArray, pos: number, space: boolean): CellToken | undefined {
  const col = columnLabelToIndex(m[2]);
  const row = parseInt(m[4], 10) - 1;
  if (col >= MAX_SHEET_COLS || row < 0 || row >= MAX_SHEET_ROWS) return undefined;
  return { kind: 'cell', pos, space, row, col, rowAbs: m[3] === '$', colAbs: m[1] === '$' };
}

/** The end of the structured reference (`Table1[[#This Row],[Col]]`) whose `[` is at `start`: brackets nest and `'` escapes the next character. */
function bracketEnd(text: string, start: number): number {
  let depth = 0;
  for (let i = start; i < text.length; i++) {
    const ch = text[i];
    if (ch === "'") i++;
    else if (ch === '[') depth++;
    else if (ch === ']' && --depth === 0) return i + 1;
  }
  return text.length;
}

/** Tokens of a formula (without its `=`); throws a syntax error on anything the grammar does not know. */
export function tokenize(text: string): Token[] {
  return scan(text, false);
}

/**
 * Like `tokenize`, but never throws: what the grammar does not know becomes a `raw` token (a structured reference up to its closing
 * bracket, an unterminated string or quote up to the end, anything else one character) and the scan goes on. For code that rewrites
 * the references of a formula the engine cannot parse and leaves the rest as it is.
 */
export function tokenizeLenient(text: string): Token[] {
  return scan(text, true);
}

function scan(text: string, lenient: boolean): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  let bang = -2; // the next `!` at or after `i` (-1 when there is none): a sheet prefix needs one, so most words skip the prefix match
  const prefixPossible = (): boolean => {
    if (bang !== -1 && bang < i) bang = text.indexOf('!', i);
    return bang !== -1 && bang - i < MAX_PREFIX_CHARS;
  };
  // A stretch the grammar rejects: an error, or in lenient mode a raw token over it. Returns where the scan goes on.
  const bad = (message: string, pos: number, space: boolean, to: number): number => {
    if (!lenient) throw syntaxError(message);
    tokens.push({ kind: 'raw', end: to, pos, space });
    return to;
  };
  while (i < text.length) {
    const start = i;
    while (i < text.length && /\s/.test(text[i])) i++;
    if (i >= text.length) break;
    const space = i > start;
    const pos = i;
    const rest = text.slice(i);
    const ch = text[i];
    const prev = tokens[tokens.length - 1];

    if (ch === '"') {
      const s = scanQuoted(text, i, '"');
      if (!s) {
        i = bad(`Unterminated string starting at position ${pos}`, pos, space, text.length);
        continue;
      }
      tokens.push({ kind: 'str', value: s.value, pos, space });
      i = s.end;
      continue;
    }

    if (ch === "'") {
      const s = scanQuoted(text, i, "'");
      if (!s) {
        i = bad(`Unterminated quote starting at position ${pos}`, pos, space, text.length);
        continue;
      }
      if (text[s.end] === '!') {
        // 'Sheet 1'!  or  'Sheet 1:Sheet 3'!  (a colon is illegal in a sheet name, so it spans sheets)
        const colon = s.value.indexOf(':');
        const [first, last] = colon < 0 ? [s.value] : [s.value.slice(0, colon), s.value.slice(colon + 1)];
        if (first === '' || last === '') {
          i = bad(`Empty sheet name at position ${pos}`, pos, space, s.end + 1);
          continue;
        }
        tokens.push({ kind: 'sheet', start: first, end: last, pos, space });
        i = s.end + 1;
      } else {
        // A single-quoted string literal (accepted for compatibility).
        tokens.push({ kind: 'str', value: s.value, pos, space });
        i = s.end;
      }
      continue;
    }

    if (ch === '#') {
      if (prev && !space && prev.kind === 'cell') {
        tokens.push({ kind: 'op', text: '#', pos, space });
        i++;
        continue;
      }
      const m = ERROR_LITERAL.exec(rest);
      if (!m) {
        i = bad(`Unknown error literal at position ${pos}`, pos, space, i + 1);
        continue;
      }
      tokens.push({ kind: 'err', value: m[0].toUpperCase(), pos, space });
      i += m[0].length;
      continue;
    }

    if (/[\d.$]/.test(ch)) {
      const rr = ROW_RANGE.exec(rest);
      if (rr) {
        const [r1, r2] = [parseInt(rr[2], 10) - 1, parseInt(rr[4], 10) - 1];
        if (r1 < 0 || r2 < 0 || r1 >= MAX_SHEET_ROWS || r2 >= MAX_SHEET_ROWS) {
          i = bad(`Row out of range at position ${pos}`, pos, space, i + rr[0].length);
          continue;
        }
        tokens.push({ kind: 'rowrange', r1, r2, abs1: rr[1] === '$', abs2: rr[3] === '$', pos, space });
        i += rr[0].length;
        continue;
      }
    }

    if (/[\d.]/.test(ch)) {
      const sheet = prefixPossible() ? sheetPrefixAt(rest, prev) : null; // a sheet named with digits: 2024!A1
      if (sheet) {
        tokens.push({ kind: 'sheet', start: sheet[1], end: sheet[2], pos, space });
        i += sheet[0].length;
        continue;
      }
      const m = NUMBER.exec(rest);
      if (!m) {
        i = bad(`Unexpected "${ch}" at position ${pos}`, pos, space, i + 1);
        continue;
      }
      const value = Number(m[0]);
      if (!Number.isFinite(value)) {
        i = bad(`Number out of range at position ${pos}`, pos, space, i + m[0].length);
        continue;
      }
      tokens.push({ kind: 'num', value, pos, space });
      i += m[0].length;
      continue;
    }

    if (ch === '$' || /[\p{L}_\\]/u.test(ch)) {
      const sp = prefixPossible() ? sheetPrefixAt(rest, prev) : null;
      if (sp) {
        tokens.push({ kind: 'sheet', start: sp[1], end: sp[2], pos, space });
        i += sp[0].length;
        continue;
      }
      const cr = COL_RANGE.exec(rest);
      if (cr) {
        const [c1, c2] = [columnLabelToIndex(cr[2]), columnLabelToIndex(cr[4])];
        if (c1 >= MAX_SHEET_COLS || c2 >= MAX_SHEET_COLS) {
          i = bad(`Column out of range at position ${pos}`, pos, space, i + cr[0].length);
          continue;
        }
        tokens.push({ kind: 'colrange', c1, c2, abs1: cr[1] === '$', abs2: cr[3] === '$', pos, space });
        i += cr[0].length;
        continue;
      }
      const cm = CELL.exec(rest);
      const cell = cm && rest[cm[0].length] !== '(' ? cellFromMatch(cm, pos, space) : undefined;
      if (cell && cm) {
        tokens.push(cell);
        i += cm[0].length;
        continue;
      }
      const id = IDENT.exec(rest);
      if (!id) {
        i = bad(`Unexpected "${ch}" at position ${pos}`, pos, space, i + 1);
        continue;
      }
      tokens.push({ kind: 'ident', text: id[0], pos, space });
      i += id[0].length;
      continue;
    }

    const two = rest.slice(0, 2);
    if (two === '<>' || two === '<=' || two === '>=') {
      tokens.push({ kind: 'op', text: two, pos, space });
      i += 2;
      continue;
    }
    if ('+-*/^&=<>%:,;(){}@'.includes(ch)) {
      tokens.push({ kind: 'op', text: ch, pos, space });
      i++;
      continue;
    }
    i = bad(`Unexpected "${ch}" at position ${pos}`, pos, space, ch === '[' ? bracketEnd(text, i) : i + 1);
  }
  tokens.push({ kind: 'eof', pos: text.length, space: false });
  return tokens;
}
