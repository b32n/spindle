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
  const code = numberFormatCode(styles, xf.numFmtId);
  if (code !== undefined && code !== 'General') look.numberFormat = { code };
  return look;
}

/**
 * Whether a number format shows dates or times: the letters `y m d h s` (or `b`, the Buddhist year) outside quotes and
 * `[...]` groups. The same test the `exceljs` reader used to decide that a number is a date.
 */
export function isDateFormat(code: string | undefined): boolean {
  if (!code) return false;
  return /[ymdhMsb]/.test(code.replace(/\[[^\]]*]/g, '').replace(/"[^"]*"/g, ''));
}
