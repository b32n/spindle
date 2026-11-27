import type {
  SpreadsheetBorder,
  SpreadsheetCellStyle,
  SpreadsheetNumberFormat,
  XlsxBorderSide,
  XlsxCellFormat,
  XlsxColor,
  XlsxStyles,
} from '@b32nio/spindle-shared/export/xlsx';
import { builtinNumberFormat } from '../styles-scan';

// ---------------------------------------------------------------------------
// A cell format (`cellXfs` entry) as the editor's flat cell style and number format. Colours are the resolved `#rrggbb` (theme,
// indexed palette and tint applied). The result is computed once per format index and shared by every cell that uses it:
// treat it as read-only. `isDefault` is the workbook's default format (index 0, what a cell with no `s` has): it carries no font name or size, no bottom alignment and no zero rotation.
// ---------------------------------------------------------------------------

export interface CellLook {
  style?: SpreadsheetCellStyle;
  numberFormat?: SpreadsheetNumberFormat;
  /** The format shows dates or times (judged on the code as the file writes it): the cell's number is a date serial. */
  isDate: boolean;
}

const HORIZONTAL: ReadonlyMap<string, 'left' | 'center' | 'right'> = new Map([['left', 'left'], ['center', 'center'], ['right', 'right']]);
const VERTICAL: ReadonlyMap<string, 'top' | 'middle' | 'bottom'> = new Map([['top', 'top'], ['center', 'middle'], ['bottom', 'bottom']]);
const BORDER_STYLES: ReadonlySet<string> = new Set(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double']);

/** Plain black text as Excel writes it: theme colour 1 (Text 1, untinted) or indexed 8 or 64. That is the application's default text colour, not a choice, so a dark theme still works. */
const isDefaultTextColor = (c: XlsxColor | undefined): boolean =>
  c !== undefined && ((c.theme === 1 && !c.tint) || c.indexed === 8 || c.indexed === 64);

const colorOf = (c: XlsxColor | undefined): string | undefined => c?.resolved;

function sideOf(side: XlsxBorderSide | undefined): SpreadsheetBorder | undefined {
  if (side === undefined) return undefined;
  const out: SpreadsheetBorder = { style: BORDER_STYLES.has(side.style) ? (side.style as SpreadsheetBorder['style']) : 'thin' };
  const color = colorOf(side.color);
  if (color !== undefined) out.color = color;
  return out;
}

/** Text rotation in degrees as the style holds it (-90..90, counter-clockwise positive); the file's 91..180 mean clockwise, 255 is stacked text (not modelled). */
function rotationOf(value: number | undefined): number | undefined {
  if (value === undefined) return undefined;
  if (value >= -90 && value <= 90) return value;
  return value > 90 && value <= 180 ? 90 - value : undefined;
}

/** Longest number format code that is read (characters): Excel's own limit. A longer one is opaque: kept as written, never scanned or copied. */
export const MAX_FORMAT_CODE_CHARS = 255;

/** What a number format id means for a cell: whether it shows dates or times, and the code the editor gets. */
interface NumberFormatLook {
  isDate: boolean;
  /** `undefined` for an id the file does not define and the table does not know. */
  code: string | undefined;
  /** The code is over `MAX_FORMAT_CODE_CHARS`: kept as written, not scanned (so `isDate` is false). */
  opaque: boolean;
}

/**
 * The number format of an id. A code the file defines drops the backslash before a character Excel shows without
 * quotes (`\\$#,##0.00` is `$#,##0.00`, `yyyy\\-mm\\-dd` is `yyyy-mm-dd`: LibreOffice escapes every literal, and the editor's
 * format mapping, `fromExcelNumberFormatCode`, expects the plain spelling) and keeps every other escape, where the backslash is
 * what makes the character a literal (`0\\ \\d` shows a `d`, it is not a day; see `unescapeLiterals`). The workbook's styles keep the
 * code as written, and the date test is on that. A built-in id's code comes from the table as it is. A code over
 * `MAX_FORMAT_CODE_CHARS` is not scanned at all: a file may name a megabyte of one for every cell format.
 */
function numberFormatOf(styles: XlsxStyles, id: number): NumberFormatLook {
  const defined = id === 0 ? undefined : definedCodes(styles).get(id);
  const raw = id === 0 ? 'General' : (defined ?? builtinNumberFormat(id));
  if (raw === undefined) return { isDate: false, code: undefined, opaque: false };
  if (raw.length > MAX_FORMAT_CODE_CHARS) return { isDate: false, code: raw, opaque: true };
  return { isDate: isDateFormat(raw), code: defined === undefined ? raw : unescapeLiterals(defined), opaque: false };
}

/** The codes a workbook's `numFmts` define, by id (the first of a repeated id wins), built once per workbook: a lookup per format must not scan the table. */
const DEFINED = new WeakMap<XlsxStyles, Map<number, string>>();
function definedCodes(styles: XlsxStyles): Map<number, string> {
  let codes = DEFINED.get(styles);
  if (codes === undefined) {
    codes = new Map();
    for (const n of styles.numFmts) if (!codes.has(n.id)) codes.set(n.id, n.code);
    DEFINED.set(styles, codes);
  }
  return codes;
}

/**
 * The looks of one workbook's cell formats, made once and shared by every cell (and every sheet) that uses them: by format index, and
 * the number formats under them by id, so a code is read once however many formats name it (a workbook of 60,000 formats naming
 * one long code costs one read, not 60,000).
 */
export class LookCache {
  private readonly looks = new Map<number, CellLook | undefined>();
  private readonly formats = new Map<number, NumberFormatLook>();

  /** `opaque` is told of each number format id whose code is over `MAX_FORMAT_CODE_CHARS`, the first time a format uses it. */
  constructor(
    readonly styles: XlsxStyles,
    private readonly opaque: (id: number) => void = () => {}
  ) {}

  numberFormat(id: number): NumberFormatLook {
    let format = this.formats.get(id);
    if (format === undefined) {
      format = numberFormatOf(this.styles, id);
      this.formats.set(id, format);
      if (format.opaque) this.opaque(id);
    }
    return format;
  }

  /** The look of the cell format at `index`, or `undefined` when the file has none there. */
  get(index: number): CellLook | undefined {
    if (this.looks.has(index)) return this.looks.get(index);
    const xf = this.styles.cellXfs[index];
    const look = xf === undefined ? undefined : lookOf(xf, this.styles, index === 0, this);
    this.looks.set(index, look);
    return look;
  }
}

export function lookOf(xf: XlsxCellFormat, styles: XlsxStyles, isDefault: boolean, cache: LookCache = new LookCache(styles)): CellLook {
  const style: SpreadsheetCellStyle = {};
  const font = styles.fonts[xf.fontId];
  if (font !== undefined) {
    if (font.bold) style.bold = true;
    if (font.italic) style.italic = true;
    if (font.underline) style.underline = true;
    if (font.strike) style.strikethrough = true;
    // the default format's font is the workbook's default font: not repeated on every cell
    if (font.name && !isDefault) style.fontFamily = font.name;
    if (font.size && !isDefault) style.fontSize = font.size;
    const color = isDefaultTextColor(font.color) ? undefined : colorOf(font.color);
    if (color !== undefined) style.fontColor = color;
  }

  const fill = styles.fills[xf.fillId];
  if (fill?.type === 'pattern' && fill.pattern === 'solid') {
    const color = colorOf(fill.fg);
    if (color !== undefined) style.backgroundColor = color;
  }

  const alignment = xf.alignment;
  if (alignment !== undefined) {
    const horizontal = alignment.horizontal === undefined ? undefined : HORIZONTAL.get(alignment.horizontal);
    if (horizontal !== undefined) style.horizontalAlign = horizontal;
    const vertical = alignment.vertical === undefined ? undefined : VERTICAL.get(alignment.vertical);
    // (the default format does not repeat the defaults: bottom, upright)
    if (vertical !== undefined && !(isDefault && vertical === 'bottom')) style.verticalAlign = vertical;
    if (alignment.wrapText) style.wrapText = true;
    const rotation = rotationOf(alignment.textRotation);
    if (rotation !== undefined && !(isDefault && rotation === 0)) style.textRotation = rotation;
  }

  const border = styles.borders[xf.borderId];
  if (border !== undefined && (border.top || border.right || border.bottom || border.left)) {
    style.border = { top: sideOf(border.top), right: sideOf(border.right), bottom: sideOf(border.bottom), left: sideOf(border.left) };
  }

  const format = cache.numberFormat(xf.numFmtId);
  const look: CellLook = { isDate: format.isDate };
  if (Object.keys(style).length > 0) look.style = style;
  if (format.code !== undefined && format.code !== 'General') look.numberFormat = { code: format.code };
  return look;
}

// ---------------------------------------------------------------------------
// Number format codes are read in one pass by `scanFormat`, which tells apart what is text and what is syntax:
//   - `"..."`: a quoted run, literal (a backslash in it is a backslash); an unclosed quote is an ordinary character
//   - `[...]`: a group (colour, condition, locale, elapsed time); an unclosed bracket is an ordinary character
//   - `\x`: the next character, literal (`\\` is a backslash, `\"` a quote that opens no run)
//   - `_x` and `*x`: a space the width of `x`, and a fill with `x`: `x` is no token
//   - anything else is one character of syntax
// The pass is linear: a closer that is not found after some point is not found after any later one, so it is looked for once.
// ---------------------------------------------------------------------------

const Piece = { Char: 0, Quoted: 1, Group: 2, Escape: 3, Pair: 4 } as const;
type Piece = (typeof Piece)[keyof typeof Piece];

/** Walk `code`, calling `visit(kind, start, end)` for each piece; a visitor that answers `true` ends the walk. */
function scanFormat(code: string, visit: (kind: Piece, start: number, end: number) => boolean | void): void {
  const n = code.length;
  let noBracket = false;
  let noQuote = false;
  let i = 0;
  while (i < n) {
    const ch = code[i];
    let kind: Piece = Piece.Char;
    let end = i + 1;
    if (ch === '"' && !noQuote) {
      const close = code.indexOf('"', i + 1);
      noQuote = close < 0;
      if (close >= 0) [kind, end] = [Piece.Quoted, close + 1];
    } else if (ch === '[' && !noBracket) {
      const close = code.indexOf(']', i + 1);
      noBracket = close < 0;
      if (close >= 0) [kind, end] = [Piece.Group, close + 1];
    } else if (ch === '\\' && i + 1 < n) [kind, end] = [Piece.Escape, i + 2];
    else if ((ch === '_' || ch === '*') && i + 1 < n) [kind, end] = [Piece.Pair, i + 2];
    if (visit(kind, i, end) === true) return;
    i = end;
  }
}

/** The characters Excel shows without quotes: a backslash before one of them is only a spelling of the plain character, except where `keepsEscape` says it is not. */
const PLAIN_LITERALS = new Set('$-+():/!^&\'~{}<>= ');

/**
 * Whether the backslash before `c` (the character at `at + 1` of `text`) must stay although `c` is one of `PLAIN_LITERALS`, because
 * without it `c` would be syntax:
 *   - `/` between digit placeholders (`0\/0`, `#\/#`, `00\/00\/0000`) is a fraction, so it stays in a section that has a placeholder
 *     (`0`, `#`, `?` outside quotes and escapes; `placeholders`), and in `AM\/PM` and `A\/P`, where the slash is part of one token;
 *   - `+` and `-` after an `E` (`0.00E\+00`) are the sign of a scientific exponent.
 * Where it is not kept the character is the same literal either way (a date, a time, a currency).
 */
function keepsEscape(text: string, at: number, c: string, placeholders: boolean): boolean {
  if (c === '/') return placeholders || (/[AaMm]/.test(text[at - 1] ?? '') && /[PpMm]/.test(text[at + 2] ?? ''));
  if (c === '+' || c === '-') return /[Ee]/.test(text[at - 1] ?? '');
  return false;
}

/** Whether each section of `code` (`;` outside quotes, groups and escapes ends one) has a digit placeholder (`0`, `#`, `?`). */
function placeholderSections(code: string): boolean[] {
  const found = [false];
  scanFormat(code, (kind, start) => {
    if (kind !== Piece.Char) return;
    const ch = code[start];
    if (ch === ';') found.push(false);
    else if (ch === '0' || ch === '#' || ch === '?') found[found.length - 1] = true;
  });
  return found;
}

/**
 * `code` without the backslash before a character that needs none (`\$`, `\-`, `\ `, `\(` and the rest of `PLAIN_LITERALS`): the
 * spelling LibreOffice and others write, and the one the editor's format mapping knows. A backslash before anything else is what
 * makes that character a literal and stays (`\d`, `\m`, `\*`, `\@`, `\.`, `\"`, `\\`, a digit, `%`, `#`, `?`, `_`), and so is one whose
 * character would be syntax without it (`keepsEscape`: `0\/0` is the text, `0/0` a fraction; `E\+00`, `AM\/PM`). What is in
 * quotes is never touched (`"C:\Temp"0`). A group (`[$\$-409]`) is unescaped the same way (a `/` in one stays).
 */
export function unescapeLiterals(code: string): string {
  if (!code.includes('\\')) return code;
  const placeholders = placeholderSections(code);
  const out: string[] = [];
  let from = 0;
  let section = 0;
  scanFormat(code, (kind, start, end) => {
    if (kind === Piece.Char && code[start] === ';') section++;
    else if (kind === Piece.Escape && PLAIN_LITERALS.has(code[start + 1]) && !keepsEscape(code, start, code[start + 1], placeholders[section])) {
      out.push(code.slice(from, start));
      from = start + 1;
    } else if (kind === Piece.Group) {
      // inside a group (`[$\\$-409]`) the same rule, with an escaped backslash read as one escape
      const group = code.slice(start, end);
      const plain = group.replace(/\\(.)/g, (whole, c: string, offset: number) => (PLAIN_LITERALS.has(c) && !keepsEscape(group, offset, c, true) ? c : whole));
      if (plain !== group) {
        out.push(code.slice(from, start), plain);
        from = end;
      }
    }
  });
  out.push(code.slice(from));
  return out.join('');
}

/**
 * Whether a number format shows dates or times: the letters `y m d h s` (or `b`, the Buddhist year), in either case, in its syntax,
 * that is outside quotes, `[...]` groups, escapes (`\d`) and the `_x` and `*x` pairs. The test the `exceljs` reader used to decide
 * that a number is a date (which saw only lower-case letters and `M`), on a correct reading of the code. Give it the code as the
 * file writes it.
 */
export function isDateFormat(code: string | undefined): boolean {
  if (!code) return false;
  let found = false;
  scanFormat(code, (kind, start) => {
    if (kind === Piece.Char && DATE_LETTER.test(code[start])) found = true;
    return found;
  });
  return found;
}

/** Format codes are case-insensitive: `DD`, `YYYY`, `H` and `S` are the same tokens as their lower-case spellings (`M` is a month or, after an hour, minutes, and a date either way). */
const DATE_LETTER = /[ymdhsb]/i;
