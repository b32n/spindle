// Cell styles between the editor's pooled `CellStyle` and the interchange `SpreadsheetCellStyle` (see export/xlsx.ts, which
// reads and writes the files). Borders are the one place the two differ in shape: the editor holds a CSS shorthand per side
// (`1px dashed #000000`) next to the exact kind the file used, the file format has thirteen line styles.

import type { SpreadsheetBorder, SpreadsheetCellStyle, SpreadsheetDefaultFormat, SpreadsheetCell } from '@b32nio/spindle-shared/export/xlsx';
import type { BorderKind, CellStyle, LineFormat } from '../types';
import type { FormatPool } from '../format-pool';
import type { StylePool } from '../style-pool';
import { CSS_OF_KIND, keptBorderKind, resolveBorderKind } from '../border-kinds';
import { fromExcelNumberFormatCode, toExcelNumberFormatCode } from './xlsx-number-format';

// CellStyle's border fields are CSS shorthand strings, e.g. "1px solid #000000".
const CSS_BORDER = /^(\d+)px\s+(\w+)\s+(#[0-9a-fA-F]{3,8}|\w+)$/;

/** The side of an exported style: the kept kind when the shorthand still says what it did, otherwise what the shorthand says. */
function toBorder(value: string | undefined, kind: BorderKind | undefined): SpreadsheetBorder | undefined {
  if (!value) return undefined;
  const match = value.match(CSS_BORDER);
  if (!match) return { style: 'thin' };
  const [, widthStr, cssStyle, color] = match;
  return { style: resolveBorderKind(Number(widthStr), cssStyle, kind), color: color.startsWith('#') ? color : undefined };
}

/** The shorthand of an imported side. */
const toCssBorder = (border: SpreadsheetBorder): string => `${CSS_OF_KIND[border.style].width}px ${CSS_OF_KIND[border.style].style} ${border.color ?? '#000000'}`;

/** A pooled style as the interchange style (what a file writer takes). */
export function toSpreadsheetStyle(style: CellStyle): SpreadsheetCellStyle {
  const out: SpreadsheetCellStyle = {};
  if (style.bold != null) out.bold = style.bold;
  if (style.italic != null) out.italic = style.italic;
  if (style.textDecoration === 'underline') {
    out.underline = true;
    if (style.underlineKind) out.underlineStyle = style.underlineKind;
    if (style.strikethrough) out.strikethrough = true;
  } else if (style.textDecoration === 'line-through' || style.strikethrough) out.strikethrough = true;
  if (style.fontFamily != null) out.fontFamily = style.fontFamily;
  if (style.fontSize != null) out.fontSize = style.fontSize;
  if (style.fontColor != null) out.fontColor = style.fontColor;
  if (style.backgroundColor != null) out.backgroundColor = style.backgroundColor;
  if (style.fillPattern != null && style.patternColor != null) {
    out.fillPattern = style.fillPattern;
    out.patternColor = style.patternColor;
  }
  if (style.textAlign != null) out.horizontalAlign = style.textAlign;
  if (style.verticalAlign != null) out.verticalAlign = style.verticalAlign;
  if (style.textWrap != null) out.wrapText = style.textWrap;
  if (style.indent != null) out.indent = style.indent;
  if (style.shrinkToFit) out.shrinkToFit = true;
  if (style.textRotation != null) out.textRotation = style.textRotation;

  if (style.borderTop || style.borderRight || style.borderBottom || style.borderLeft) {
    out.border = {
      top: toBorder(style.borderTop, style.borderTopKind),
      right: toBorder(style.borderRight, style.borderRightKind),
      bottom: toBorder(style.borderBottom, style.borderBottomKind),
      left: toBorder(style.borderLeft, style.borderLeftKind),
    };
  }
  return out;
}

/** An interchange style as a pooled style (what a file reader gives). */
export function fromSpreadsheetStyle(style: SpreadsheetCellStyle): CellStyle {
  const out: CellStyle = {};
  if (style.bold != null) out.bold = style.bold;
  if (style.italic != null) out.italic = style.italic;
  if (style.underline) {
    out.textDecoration = 'underline';
    if (style.underlineStyle) out.underlineKind = style.underlineStyle;
    if (style.strikethrough) out.strikethrough = true;
  } else if (style.strikethrough) out.textDecoration = 'line-through';
  if (style.fontFamily != null) out.fontFamily = style.fontFamily;
  if (style.fontSize != null) out.fontSize = style.fontSize;
  if (style.fontColor != null) out.fontColor = style.fontColor;
  if (style.backgroundColor != null) out.backgroundColor = style.backgroundColor;
  if (style.fillPattern != null && style.patternColor != null) {
    out.fillPattern = style.fillPattern;
    out.patternColor = style.patternColor;
  }
  if (style.horizontalAlign != null) out.textAlign = style.horizontalAlign;
  if (style.verticalAlign != null) out.verticalAlign = style.verticalAlign;
  if (style.wrapText != null) out.textWrap = style.wrapText;
  if (style.indent) out.indent = style.indent;
  if (style.shrinkToFit) out.shrinkToFit = true;
  // an explicit 0 is the default: the file writer leaves it out, so the model does not keep it either (open -> save -> open is stable)
  if (style.textRotation) out.textRotation = style.textRotation;

  if (style.border) {
    const sides = [['Top', style.border.top], ['Right', style.border.right], ['Bottom', style.border.bottom], ['Left', style.border.left]] as const;
    for (const [name, side] of sides) {
      if (!side) continue;
      out[`border${name}`] = toCssBorder(side);
      const kind = keptBorderKind(side.style);
      if (kind) out[`border${name}Kind`] = kind;
    }
  }
  return out;
}

/**
 * What a file import takes from the reader into the pools: styles and number formats looked up once per distinct object (the reader gives
 * every cell of one cell format the same object), and the pools capped. A million styled cells cost a million lookups, not a million
 * conversions; a hostile file with more distinct formats than the pool may hold (Excel itself stops at 64,000) keeps the first ones
 * and its other cells unformatted, counted in `cut`.
 */
export class StyleImport {
  private readonly styleIds = new Map<object, string | undefined>();
  private readonly formatIds = new Map<object, string | undefined>();
  /** Distinct styles and number formats the import has refused (the pools are full), and row/column formats past the sheet's extent. */
  cut = { styles: 0, formats: 0, lines: 0 };

  constructor(
    readonly styles: StylePool,
    readonly formats: FormatPool,
    private readonly limit: number = MAX_IMPORTED_STYLES
  ) {}

  style(style: SpreadsheetCellStyle): string | undefined {
    if (this.styleIds.has(style)) return this.styleIds.get(style);
    const converted = fromSpreadsheetStyle(style);
    const known = this.styles.has(converted);
    const id = !known && this.styles.size() >= this.limit ? undefined : this.styles.getOrCreate(converted);
    if (id === undefined) this.cut.styles++;
    this.styleIds.set(style, id);
    return id;
  }

  format(numberFormat: NonNullable<SpreadsheetCell['numberFormat']>): string | undefined {
    if (this.formatIds.has(numberFormat)) return this.formatIds.get(numberFormat);
    const format = fromExcelNumberFormatCode(numberFormat.code);
    let id: string | undefined;
    if (format.type) {
      id = !this.formats.has(format) && this.formats.size() >= this.limit ? undefined : this.formats.getOrCreate(format);
      if (id === undefined) this.cut.formats++;
    }
    this.formatIds.set(numberFormat, id);
    return id;
  }

  /** The default format of a column or row as the model holds it (`undefined`: nothing to keep). */
  line(format: SpreadsheetDefaultFormat): LineFormat | undefined {
    const styleId = format.style ? this.style(format.style) : undefined;
    const formatId = format.numberFormat ? this.format(format.numberFormat) : undefined;
    return styleId || formatId ? { ...(styleId ? { styleId } : {}), ...(formatId ? { formatId } : {}) } : undefined;
  }
}

/** Most distinct styles, and as many number formats, one import puts in the pools (Excel's own limit is 64,000 cell formats). */
export const MAX_IMPORTED_STYLES = 65_536;

/** The default format of a column or row as the interchange format (a file writer takes it); `undefined` when it names nothing that exists. */
export function toDefaultFormat(format: LineFormat, styles: StylePool, formats: FormatPool): SpreadsheetDefaultFormat | undefined {
  const style = format.styleId ? styles.get(format.styleId) : undefined;
  const numberFormat = format.formatId ? formats.get(format.formatId) : undefined;
  if (!style && !numberFormat) return undefined;
  return {
    ...(style ? { style: toSpreadsheetStyle(style) } : {}),
    ...(numberFormat ? { numberFormat: { code: toExcelNumberFormatCode(numberFormat) } } : {}),
  };
}
