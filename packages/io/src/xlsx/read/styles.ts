import type {
  XlsxAlignment,
  XlsxBorder,
  XlsxBorderSide,
  XlsxCellFormat,
  XlsxDifferentialFormat,
  XlsxFill,
  XlsxNamedStyle,
  XlsxProtection,
  XlsxStyles,
} from '@b32nio/spindle-shared/export/xlsx';
import { kid, kids, type XmlElement } from '../xml';
import { builtinNumberFormat } from '../styles-scan';
import { DEFAULT_INDEXED_COLORS, parseRgb, readColor, type ColorContext } from './colors';
import { readFont } from './font';
import { boolAttr, intAttr, numAttr, put } from './values';

// ---------------------------------------------------------------------------
// The styles part (ECMA-376 18.8): number formats, fonts, fills (pattern and gradient), borders with colours, the two format
// tables (`cellStyleXfs`, `cellXfs`), named styles and `dxfs`. Every colour carries its resolved `#rrggbb` (theme, indexed
// palette and tint applied). Tables keep the file's order and length, so the ids a cell's `s` attribute and an xf's
// `fontId`/`fillId`/`borderId` hold index straight into them.
// ---------------------------------------------------------------------------

const list = (root: XmlElement | undefined, group: string, item: string): XmlElement[] => kids(kid(root, group), item);

/** The workbook's own `indexedColors`; an entry that is not a colour keeps the default for its index. `undefined` when there is none. A short palette stays short: `resolveColor` falls back to the default per index. */
function readPalette(root: XmlElement | undefined): string[] | undefined {
  const entries = list(kid(root, 'colors'), 'indexedColors', 'rgbColor');
  if (entries.length === 0) return undefined;
  return entries.map((e, i) => parseRgb(e.attrs.rgb) ?? DEFAULT_INDEXED_COLORS[i] ?? '#000000');
}

function readPatternOrGradient(el: XmlElement | undefined, ctx: ColorContext): XlsxFill {
  const gradient = kid(el, 'gradientFill');
  if (gradient) {
    const out: XlsxFill = {
      type: 'gradient',
      gradientType: gradient.attrs.type === 'path' ? 'path' : 'linear',
      stops: kids(gradient, 'stop').map((s) => ({ position: numAttr(s.attrs.position) ?? 0, color: readColor(kid(s, 'color'), ctx) ?? {} })),
    };
    put(out, 'degree', numAttr(gradient.attrs.degree));
    put(out, 'left', numAttr(gradient.attrs.left));
    put(out, 'right', numAttr(gradient.attrs.right));
    put(out, 'top', numAttr(gradient.attrs.top));
    put(out, 'bottom', numAttr(gradient.attrs.bottom));
    return out;
  }
  const pattern = kid(el, 'patternFill');
  const out: XlsxFill = { type: 'pattern', pattern: pattern?.attrs.patternType ?? 'none' };
  put(out, 'fg', readColor(kid(pattern, 'fgColor'), ctx));
  put(out, 'bg', readColor(kid(pattern, 'bgColor'), ctx));
  return out;
}

function readSide(el: XmlElement | undefined, ctx: ColorContext): XlsxBorderSide | undefined {
  const style = el?.attrs.style;
  if (!el || style === undefined) return undefined;
  const side: XlsxBorderSide = { style };
  put(side, 'color', readColor(kid(el, 'color'), ctx));
  return side;
}

function readBorder(el: XmlElement | undefined, ctx: ColorContext): XlsxBorder {
  const border: XlsxBorder = {};
  // `start`/`end` are the right-to-left-aware names some producers (and the strict schema) use for left/right
  put(border, 'left', readSide(kid(el, 'left') ?? kid(el, 'start'), ctx));
  put(border, 'right', readSide(kid(el, 'right') ?? kid(el, 'end'), ctx));
  put(border, 'top', readSide(kid(el, 'top'), ctx));
  put(border, 'bottom', readSide(kid(el, 'bottom'), ctx));
  put(border, 'diagonal', readSide(kid(el, 'diagonal'), ctx));
  put(border, 'diagonalUp', boolAttr(el?.attrs.diagonalUp));
  put(border, 'diagonalDown', boolAttr(el?.attrs.diagonalDown));
  put(border, 'outline', boolAttr(el?.attrs.outline));
  return border;
}

function readAlignment(el: XmlElement | undefined): XlsxAlignment | undefined {
  if (!el) return undefined;
  const a = el.attrs;
  const out: XlsxAlignment = {};
  put(out, 'horizontal', a.horizontal);
  put(out, 'vertical', a.vertical);
  put(out, 'textRotation', intAttr(a.textRotation));
  put(out, 'wrapText', boolAttr(a.wrapText));
  put(out, 'indent', intAttr(a.indent));
  put(out, 'relativeIndent', intAttr(a.relativeIndent));
  put(out, 'justifyLastLine', boolAttr(a.justifyLastLine));
  put(out, 'shrinkToFit', boolAttr(a.shrinkToFit));
  put(out, 'readingOrder', intAttr(a.readingOrder));
  return out;
}

function readProtection(el: XmlElement | undefined): XlsxProtection | undefined {
  if (!el) return undefined;
  const out: XlsxProtection = {};
  put(out, 'locked', boolAttr(el.attrs.locked));
  put(out, 'hidden', boolAttr(el.attrs.hidden));
  return out;
}

function readFormat(xf: XmlElement): XlsxCellFormat {
  const a = xf.attrs;
  const out: XlsxCellFormat = {
    numFmtId: intAttr(a.numFmtId) ?? 0,
    fontId: intAttr(a.fontId) ?? 0,
    fillId: intAttr(a.fillId) ?? 0,
    borderId: intAttr(a.borderId) ?? 0,
  };
  put(out, 'xfId', intAttr(a.xfId));
  put(out, 'applyNumberFormat', boolAttr(a.applyNumberFormat));
  put(out, 'applyFont', boolAttr(a.applyFont));
  put(out, 'applyFill', boolAttr(a.applyFill));
  put(out, 'applyBorder', boolAttr(a.applyBorder));
  put(out, 'applyAlignment', boolAttr(a.applyAlignment));
  put(out, 'applyProtection', boolAttr(a.applyProtection));
  put(out, 'quotePrefix', boolAttr(a.quotePrefix));
  put(out, 'pivotButton', boolAttr(a.pivotButton));
  put(out, 'alignment', readAlignment(kid(xf, 'alignment')));
  put(out, 'protection', readProtection(kid(xf, 'protection')));
  return out;
}

function readDxf(dxf: XmlElement, ctx: ColorContext): XlsxDifferentialFormat {
  const out: XlsxDifferentialFormat = {};
  put(out, 'font', readFont(kid(dxf, 'font'), ctx));
  const numFmt = kid(dxf, 'numFmt');
  const id = intAttr(numFmt?.attrs.numFmtId);
  if (numFmt && id !== undefined) out.numFmt = { id, code: numFmt.attrs.formatCode ?? builtinNumberFormat(id) ?? '' };
  const fill = kid(dxf, 'fill');
  if (fill) out.fill = readPatternOrGradient(fill, ctx);
  put(out, 'alignment', readAlignment(kid(dxf, 'alignment')));
  const border = kid(dxf, 'border');
  if (border) out.border = readBorder(border, ctx);
  put(out, 'protection', readProtection(kid(dxf, 'protection')));
  return out;
}

function readNamedStyle(el: XmlElement): XlsxNamedStyle | undefined {
  const name = el.attrs.name;
  if (name === undefined) return undefined;
  const style: XlsxNamedStyle = { name, xfId: intAttr(el.attrs.xfId) ?? 0 };
  put(style, 'builtinId', intAttr(el.attrs.builtinId));
  put(style, 'hidden', boolAttr(el.attrs.hidden));
  return style;
}

/**
 * Read a styles part; `themeColors` is the theme's 12 colours (`parseTheme(...).colors`). A missing root gives empty tables
 * (a cell then falls back to the defaults of whoever reads them).
 */
export function parseStyles(root: XmlElement | undefined, themeColors: readonly string[]): XlsxStyles {
  const indexedColors = readPalette(root);
  const ctx: ColorContext = { themeColors, indexedColors };
  const numFmts: XlsxStyles['numFmts'] = [];
  for (const n of list(root, 'numFmts', 'numFmt')) {
    const id = intAttr(n.attrs.numFmtId);
    if (id !== undefined) numFmts.push({ id, code: n.attrs.formatCode ?? '' });
  }
  const styles: XlsxStyles = {
    numFmts,
    fonts: list(root, 'fonts', 'font').map((f) => readFont(f, ctx) as NonNullable<ReturnType<typeof readFont>>),
    fills: list(root, 'fills', 'fill').map((f) => readPatternOrGradient(f, ctx)),
    borders: list(root, 'borders', 'border').map((b) => readBorder(b, ctx)),
    cellStyleXfs: list(root, 'cellStyleXfs', 'xf').map(readFormat),
    cellXfs: list(root, 'cellXfs', 'xf').map(readFormat),
    cellStyles: list(root, 'cellStyles', 'cellStyle').flatMap((s) => readNamedStyle(s) ?? []),
    dxfs: list(root, 'dxfs', 'dxf').map((d) => readDxf(d, ctx)),
  };
  put(styles, 'indexedColors', indexedColors);
  return styles;
}

/** The format code of a number format id: a custom code from the file, else the built-in's (en-US) code, else `undefined` (a locale-specific built-in). */
export function numberFormatCode(styles: Pick<XlsxStyles, 'numFmts'>, id: number): string | undefined {
  if (id === 0) return 'General';
  return styles.numFmts.find((n) => n.id === id)?.code ?? builtinNumberFormat(id);
}
