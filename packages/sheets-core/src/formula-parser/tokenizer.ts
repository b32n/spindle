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
  | (Positioned & { kind: 'eof' });

export function syntaxError(message: string): Error {
  return new Error(`#ERROR! ${message}`);
}

// Unquoted sheet names: letters, digits, `_` and `.`, not starting with a digit.
const NAME = String.raw`[\p{L}_\\][\p{L}\p{N}_.]*`;
const SHEET_PREFIX = new RegExp(`^(${NAME})(?::(${NAME}))?!`, 'u');
const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w.])/;
const CELL_ADDRESS = /^[A-Za-z]{1,3}\d+$/;
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

export function tokenize(text: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
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
      if (!s) throw syntaxError(`Unterminated string starting at position ${pos}`);
      tokens.push({ kind: 'str', value: s.value, pos, space });
      i = s.end;
      continue;
    }

    if (ch === "'") {
      const s = scanQuoted(text, i, "'");
      if (!s) throw syntaxError(`Unterminated quote starting at position ${pos}`);
      if (text[s.end] === '!') {
        // 'Sheet 1'!  or  'Sheet 1:Sheet 3'!  (a colon is illegal in a sheet name, so it spans sheets)
        const colon = s.value.indexOf(':');
        const [first, last] = colon < 0 ? [s.value] : [s.value.slice(0, colon), s.value.slice(colon + 1)];
        if (first === '' || last === '') throw syntaxError(`Empty sheet name at position ${pos}`);
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
      if (!m) throw syntaxError(`Unknown error literal at position ${pos}`);
      tokens.push({ kind: 'err', value: m[0].toUpperCase(), pos, space });
      i += m[0].length;
      continue;
    }

    if (/[\d.$]/.test(ch)) {
      const rr = ROW_RANGE.exec(rest);
      if (rr) {
        const [r1, r2] = [parseInt(rr[2], 10) - 1, parseInt(rr[4], 10) - 1];
        if (r1 < 0 || r2 < 0 || r1 >= MAX_SHEET_ROWS || r2 >= MAX_SHEET_ROWS) {
          throw syntaxError(`Row out of range at position ${pos}`);
        }
        tokens.push({ kind: 'rowrange', r1, r2, abs1: rr[1] === '$', abs2: rr[3] === '$', pos, space });
        i += rr[0].length;
        continue;
      }
    }

    if (/[\d.]/.test(ch)) {
      const m = NUMBER.exec(rest);
      if (!m) throw syntaxError(`Unexpected "${ch}" at position ${pos}`);
      const value = Number(m[0]);
      if (!Number.isFinite(value)) throw syntaxError(`Number out of range at position ${pos}`);
      tokens.push({ kind: 'num', value, pos, space });
      i += m[0].length;
      continue;
    }

    if (ch === '$' || /[\p{L}_\\]/u.test(ch)) {
      const sp = SHEET_PREFIX.exec(rest);
      // `A1:Sheet2!B2` is a range from A1 to a cell on Sheet2, not a span of sheets "A1" to "Sheet2":
      // a sheet named like a cell address must be quoted.
      if (sp && !(sp[2] !== undefined && (CELL_ADDRESS.test(sp[1]) || CELL_ADDRESS.test(sp[2])))) {
        tokens.push({ kind: 'sheet', start: sp[1], end: sp[2], pos, space });
        i += sp[0].length;
        continue;
      }
      const cr = COL_RANGE.exec(rest);
      if (cr) {
        const [c1, c2] = [columnLabelToIndex(cr[2]), columnLabelToIndex(cr[4])];
        if (c1 >= MAX_SHEET_COLS || c2 >= MAX_SHEET_COLS) throw syntaxError(`Column out of range at position ${pos}`);
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
      if (!id) throw syntaxError(`Unexpected "${ch}" at position ${pos}`);
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
    throw syntaxError(`Unexpected "${ch}" at position ${pos}`);
  }
  tokens.push({ kind: 'eof', pos: text.length, space: false });
  return tokens;
}
