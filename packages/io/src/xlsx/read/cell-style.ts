import type {
  SpreadsheetBorder,
  SpreadsheetCellStyle,
  SpreadsheetNumberFormat,
  XlsxBorderSide,
  XlsxCellFormat,
  XlsxColor,
  XlsxStyles,
} from '@b32nio/spindle-shared/export/xlsx';
import { numberFormatCode } from './styles';

// ---------------------------------------------------------------------------
// A cell format (`cellXfs` entry) as the editor's flat cell style and number format. Colours are the resolved `#rrggbb` (theme,
// indexed palette and tint applied). The result is computed once per format index and shared by every cell that uses it:
// treat it as read-only. `isDefault` is the workbook's default format (index 0, what a cell with no `s` has): it carries no font name or size, no bottom alignment and no zero rotation.
// ---------------------------------------------------------------------------

export interface CellLook {
  style?: SpreadsheetCellStyle;
  numberFormat?: SpreadsheetNumberFormat;
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
 * The number format code the editor gets. A code the file defines drops its backslash escapes (`\\$#,##0.00` is `$#,##0.00`,
 * `yyyy\\-mm\\-dd` is `yyyy-mm-dd`): the spelling the `exceljs` reader gave (`numfmt-xform.js`) and the one the editor's format
 * mapping (`fromExcelNumberFormatCode`) expects, so a code another producer escapes (LibreOffice escapes every literal) still
 * maps onto a structured format or survives a save as the same custom code. The workbook's styles keep the code as written.
 * A built-in id's code comes from the table as it is.
 */
function editorFormatCode(styles: XlsxStyles, id: number): string | undefined {
  const defined = id === 0 ? undefined : definedCodes(styles).get(id);
  return defined === undefined ? numberFormatCode(styles, id) : defined.replace(/\\(.)/g, '$1');
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

  const look: CellLook = {};
  if (Object.keys(style).length > 0) look.style = style;
  const code = editorFormatCode(styles, xf.numFmtId);
  if (code !== undefined && code !== 'General') look.numberFormat = { code };
  return look;
}

/**
 * Whether a number format shows dates or times: the letters `y m d h s` (or `b`, the Buddhist year) outside quotes and
 * `[...]` groups. The same test the `exceljs` reader used to decide that a number is a date.
 */
export function isDateFormat(code: string | undefined): boolean {
  if (!code) return false;
  return /[ymdhMsb]/.test(withoutGroups(code));
}

/** `code` without its `[...]` groups and `"..."` runs (an unclosed one is kept as it is), in one pass: the regex forms of this are quadratic on a long run of `[`. */
function withoutGroups(code: string): string {
  let out = '';
  let i = 0;
  // once a closer is not found after some point it is not found after any later one: no second scan of the rest
  let noBracket = false;
  let noQuote = false;
  while (i < code.length) {
    const ch = code[i];
    let end = -1;
    if (ch === '[' && !noBracket) {
      end = code.indexOf(']', i + 1);
      noBracket = end < 0;
    } else if (ch === '"' && !noQuote) {
      end = code.indexOf('"', i + 1);
      noQuote = end < 0;
    }
    if (end < 0) out += code[i++];
    else i = end + 1;
  }
  return out;
}
