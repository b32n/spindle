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

/**
 * The number format code the editor gets. A code the file defines drops the backslash before a character Excel shows without
 * quotes (`\\$#,##0.00` is `$#,##0.00`, `yyyy\\-mm\\-dd` is `yyyy-mm-dd`: LibreOffice escapes every literal, and the editor's
 * format mapping, `fromExcelNumberFormatCode`, expects the plain spelling) and keeps every other escape, where the backslash is
 * what makes the character a literal (`0\\ \\d` shows a `d`, it is not a day; see `unescapeLiterals`). The workbook's styles keep the
 * code as written. A built-in id's code comes from the table as it is.
 */
function editorFormatCode(styles: XlsxStyles, id: number): string | undefined {
  if (id === 0) return 'General';
  const defined = definedCodes(styles).get(id);
  return defined === undefined ? builtinNumberFormat(id) : unescapeLiterals(defined);
}

/** The code of an id as the file or the built-in table has it (not unescaped), for the date test. */
function rawFormatCode(styles: XlsxStyles, id: number): string | undefined {
  return id === 0 ? 'General' : (definedCodes(styles).get(id) ?? builtinNumberFormat(id));
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

export function lookOf(xf: XlsxCellFormat, styles: XlsxStyles, isDefault: boolean): CellLook {
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

  const look: CellLook = { isDate: isDateFormat(rawFormatCode(styles, xf.numFmtId)) };
  if (Object.keys(style).length > 0) look.style = style;
  const code = editorFormatCode(styles, xf.numFmtId);
  if (code !== undefined && code !== 'General') look.numberFormat = { code };
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

/** The characters Excel shows without quotes: a backslash before one of them is only a spelling of the plain character. */
const PLAIN_LITERALS = new Set('$-+():/!^&\'~{}<>= ');

/**
 * `code` without the backslash before a character that needs none (`\$`, `\-`, `\ `, `\(` and the rest of `PLAIN_LITERALS`): the
 * spelling LibreOffice and others write, and the one the editor's format mapping knows. A backslash before anything else is what
 * makes that character a literal and stays (`\d`, `\m`, `\*`, `\@`, `\.`, `\"`, `\\`, a digit, `%`, `#`, `?`, `_`), and what is in
 * quotes is never touched (`"C:\Temp"0`). A group (`[$\$-409]`) is unescaped the same way.
 */
export function unescapeLiterals(code: string): string {
  if (!code.includes('\\')) return code;
  const out: string[] = [];
  let from = 0;
  scanFormat(code, (kind, start, end) => {
    if (kind === Piece.Escape && PLAIN_LITERALS.has(code[start + 1])) {
      out.push(code.slice(from, start));
      from = start + 1;
    } else if (kind === Piece.Group) {
      // inside a group (`[$\\$-409]`) the same rule, with an escaped backslash read as one escape
      const group = code.slice(start, end);
      const plain = group.replace(/\\(.)/g, (whole, c: string) => (PLAIN_LITERALS.has(c) ? c : whole));
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
 * Whether a number format shows dates or times: the letters `y m d h s` (or `b`, the Buddhist year) in its syntax, that is
 * outside quotes, `[...]` groups, escapes (`\d`) and the `_x` and `*x` pairs. The test the `exceljs` reader used to decide
 * that a number is a date, on a correct reading of the code. Give it the code as the file writes it.
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

const DATE_LETTER = /[ymdhMsb]/;
