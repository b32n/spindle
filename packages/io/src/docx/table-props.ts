import type {
  DocBorder,
  DocBorders,
  DocCellVerticalAlign,
  DocRunBase,
  DocTableLook,
  DocTableStyle,
  DocTableStylePart,
} from '@b32nio/spindle-shared/export/docx';

import { preferRecorded, themeColor, type Theme } from './theme';
import { attrsOf, childrenOf, find, findAll, type XmlNode } from './xml-tree';

// ---------------------------------------------------------------------------
// Table properties: borders, shading, vertical alignment, `w:tblLook`, and the formatting a table style
// gives the parts of a table (`w:tblStylePr`), plus the rule that picks the parts a cell belongs to.
// ---------------------------------------------------------------------------

const TWIPS_TO_PX = 96 / 1440;

/** Border line styles: Word has ~190, the model has five. */
function borderStyle(val: string | undefined): DocBorder['style'] {
  if (!val || val === 'nil' || val === 'none') return 'none';
  if (val === 'dotted') return 'dotted';
  if (/^dash|^dotDash|^dotDotDash/.test(val)) return 'dashed';
  if (/^double|^triple|^thinThick|^thickThin|^wave|^doubleWave/.test(val)) return 'double';
  return 'single';
}

function parseBorder(el: XmlNode, theme: Theme | undefined): DocBorder {
  const a = attrsOf(el);
  const style = borderStyle(a['@_w:val']);
  if (style === 'none') return { style };
  // `w:sz` is in eighths of a point; a hairline still shows as one pixel.
  const sz = Number(a['@_w:sz']);
  const widthPx = Number.isFinite(sz) && sz > 0 ? Math.max(1, Math.round((sz / 8) * (96 / 72) * 100) / 100) : 1;
  const fromTheme = themeColor(theme, a['@_w:themeColor'], a['@_w:themeTint'], a['@_w:themeShade']);
  const val = a['@_w:color'];
  const color = fromTheme ? preferRecorded(fromTheme, val) : val && val !== 'auto' && /^[0-9a-fA-F]{6}$/.test(val) ? `#${val.toLowerCase()}` : undefined;
  return { style, ...(widthPx !== 1 && { widthPx }), ...(color && { color }) };
}

/** `w:tblBorders` / `w:tcBorders`: the sides that are present. */
export function parseBorders(el: XmlNode | undefined, theme: Theme | undefined): DocBorders | undefined {
  if (!el) return undefined;
  const kids = childrenOf(el);
  const out: DocBorders = {};
  const sides: Array<[keyof DocBorders, string[]]> = [
    ['top', ['w:top']],
    ['left', ['w:left', 'w:start']],
    ['bottom', ['w:bottom']],
    ['right', ['w:right', 'w:end']],
    ['insideH', ['w:insideH']],
    ['insideV', ['w:insideV']],
  ];
  for (const [key, tags] of sides) {
    const node = tags.map((t) => find(kids, t)).find(Boolean);
    if (node) out[key] = parseBorder(node, theme);
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/** The fill of a `w:shd` as a hex colour; automatic or absent fills are no shading. */
export function parseShading(el: XmlNode | undefined, theme: Theme | undefined): string | undefined {
  if (!el) return undefined;
  const a = attrsOf(el);
  const fromTheme = themeColor(theme, a['@_w:themeFill'], a['@_w:themeFillTint'], a['@_w:themeFillShade']);
  const fill = a['@_w:fill'];
  if (fromTheme) return preferRecorded(fromTheme, fill);
  if (fill && fill !== 'auto' && /^[0-9a-fA-F]{6}$/.test(fill)) return `#${fill.toLowerCase()}`;
  const color = a['@_w:color'];
  if (a['@_w:val'] === 'solid' && color && color !== 'auto' && /^[0-9a-fA-F]{6}$/.test(color)) return `#${color.toLowerCase()}`;
  return undefined;
}

export function parseVAlign(el: XmlNode | undefined): DocCellVerticalAlign | undefined {
  const v = el ? attrsOf(el)['@_w:val'] : undefined;
  return v === 'center' || v === 'bottom' ? v : v === 'top' ? 'top' : undefined;
}

const LOOK_BITS: Array<[keyof DocTableLook, number, string]> = [
  ['firstRow', 0x0020, 'w:firstRow'],
  ['lastRow', 0x0040, 'w:lastRow'],
  ['firstColumn', 0x0080, 'w:firstColumn'],
  ['lastColumn', 0x0100, 'w:lastColumn'],
  ['noHBand', 0x0200, 'w:noHBand'],
  ['noVBand', 0x0400, 'w:noVBand'],
];

/** `w:tblLook`: the hex `w:val` bit mask, or the separate attributes newer Word writes next to it. */
export function parseLook(el: XmlNode | undefined): DocTableLook | undefined {
  if (!el) return undefined;
  const a = attrsOf(el);
  const mask = a['@_w:val'] ? parseInt(a['@_w:val'], 16) : NaN;
  const out: DocTableLook = {};
  for (const [key, bit, attrName] of LOOK_BITS) {
    const attr = a[`@_${attrName}`];
    const on = attr !== undefined ? attr === '1' || attr === 'true' : Number.isNaN(mask) ? false : (mask & bit) !== 0;
    out[key] = on;
  }
  return out;
}

/** Table-style parts, lowest precedence first (a later part overrides an earlier one). */
const PART_ORDER = ['wholeTable', 'band1Vert', 'band2Vert', 'band1Horz', 'band2Horz', 'lastCol', 'firstCol', 'lastRow', 'firstRow', 'seCell', 'swCell', 'neCell', 'nwCell'] as const;

export type TableStyleCondition = (typeof PART_ORDER)[number];

export interface TablePartParsers {
  /** `w:rPr` -> the run formatting the interchange carries (explicit "off" included). */
  text: (rPr: XmlNode | undefined) => DocRunBase;
}

function parsePart(props: XmlNode[], theme: Theme | undefined, parsers: TablePartParsers): DocTableStylePart | undefined {
  const tcPr = find(props, 'w:tcPr');
  const tcKids = tcPr ? childrenOf(tcPr) : [];
  const text = parsers.text(find(props, 'w:rPr'));
  const backgroundColor = parseShading(find(tcKids, 'w:shd'), theme);
  const borders = parseBorders(find(tcKids, 'w:tcBorders'), theme);
  const verticalAlign = parseVAlign(find(tcKids, 'w:vAlign'));
  const part: DocTableStylePart = {
    ...(Object.keys(text).length > 0 && { text }),
    ...(backgroundColor && { backgroundColor }),
    ...(borders && { borders }),
    ...(verticalAlign && { verticalAlign }),
  };
  return Object.keys(part).length > 0 ? part : undefined;
}

/** The table-specific content of a `w:style w:type="table"` element. */
export function parseTableStyle(styleKids: XmlNode[], theme: Theme | undefined, parsers: TablePartParsers): DocTableStyle | undefined {
  const tblPr = find(styleKids, 'w:tblPr');
  const tblKids = tblPr ? childrenOf(tblPr) : [];
  const borders = parseBorders(find(tblKids, 'w:tblBorders'), theme);
  const whole = parsePart(styleKids, theme, parsers);
  const rowBand = Number(attrsOf(find(tblKids, 'w:tblStyleRowBandSize') ?? {})['@_w:val']);
  const colBand = Number(attrsOf(find(tblKids, 'w:tblStyleColBandSize') ?? {})['@_w:val']);
  const conditions: Record<string, DocTableStylePart> = {};
  for (const pr of findAll(styleKids, 'w:tblStylePr')) {
    const type = attrsOf(pr)['@_w:type'];
    if (!type) continue;
    const kids = childrenOf(pr);
    const part = parsePart(kids, theme, parsers);
    // A conditional part may restate table-level borders (`w:tblPr/w:tblBorders`); they apply to its cells.
    const partTbl = find(kids, 'w:tblPr');
    const tblBorders = partTbl ? parseBorders(find(childrenOf(partTbl), 'w:tblBorders'), theme) : undefined;
    const merged: DocTableStylePart | undefined = tblBorders ? { ...part, borders: { ...tblBorders, ...part?.borders } } : part;
    if (merged) conditions[type] = merged;
  }
  const out: DocTableStyle = {
    ...(borders && { borders }),
    ...(whole && { whole }),
    ...(Object.keys(conditions).length > 0 && { conditions }),
    ...(rowBand > 1 && { rowBandSize: rowBand }),
    ...(colBand > 1 && { colBandSize: colBand }),
  };
  return Object.keys(out).length > 0 ? out : undefined;
}

export interface CellPlace {
  row: number;
  rows: number;
  /** First grid column the cell occupies. */
  col: number;
  colspan: number;
  cols: number;
}

/** The parts of a table style that apply to a cell, lowest precedence first, as Word's `tblLook` and banding rules select them. */
export function applicableParts(style: DocTableStyle, look: DocTableLook | undefined, place: CellPlace): DocTableStylePart[] {
  const conditions = style.conditions ?? {};
  const rowBand = style.rowBandSize ?? 1;
  const colBand = style.colBandSize ?? 1;
  const firstRow = !!look?.firstRow;
  const lastRow = !!look?.lastRow;
  const firstCol = !!look?.firstColumn;
  const lastCol = !!look?.lastColumn;
  const isFirstRow = place.row === 0;
  const isLastRow = place.row === place.rows - 1;
  const isFirstCol = place.col === 0;
  const isLastCol = place.col + place.colspan >= place.cols;
  const active = new Set<TableStyleCondition>(['wholeTable']);

  if (!look?.noHBand && !(firstRow && isFirstRow) && !(lastRow && isLastRow)) {
    const index = place.row - (firstRow ? 1 : 0);
    active.add(Math.floor(index / rowBand) % 2 === 0 ? 'band1Horz' : 'band2Horz');
  }
  if (!look?.noVBand && !(firstCol && isFirstCol) && !(lastCol && isLastCol)) {
    const index = place.col - (firstCol ? 1 : 0);
    active.add(Math.floor(index / colBand) % 2 === 0 ? 'band1Vert' : 'band2Vert');
  }
  if (firstCol && isFirstCol) active.add('firstCol');
  if (lastCol && isLastCol) active.add('lastCol');
  if (firstRow && isFirstRow) active.add('firstRow');
  if (lastRow && isLastRow) active.add('lastRow');
  if (firstRow && firstCol && isFirstRow && isFirstCol) active.add('nwCell');
  if (firstRow && lastCol && isFirstRow && isLastCol) active.add('neCell');
  if (lastRow && firstCol && isLastRow && isFirstCol) active.add('swCell');
  if (lastRow && lastCol && isLastRow && isLastCol) active.add('seCell');

  const parts: DocTableStylePart[] = [];
  for (const key of PART_ORDER) {
    if (!active.has(key)) continue;
    const part = key === 'wholeTable' ? mergeParts(style.whole, conditions.wholeTable) : conditions[key];
    if (part) parts.push(part);
  }
  return parts;
}

function mergeParts(a: DocTableStylePart | undefined, b: DocTableStylePart | undefined): DocTableStylePart | undefined {
  if (!a || !b) return a ?? b;
  return { ...a, ...b, ...(a.borders || b.borders ? { borders: { ...a.borders, ...b.borders } } : {}), ...(a.text || b.text ? { text: { ...a.text, ...b.text } } : {}) };
}

const mergeBorders = (a: DocBorders | undefined, b: DocBorders | undefined): DocBorders | undefined => (a || b ? { ...a, ...b } : undefined);

/** A table style layered over the one it is based on: later values replace earlier ones side by side, part by part. */
export function mergeTableStyle(base: DocTableStyle | undefined, top: DocTableStyle | undefined): DocTableStyle | undefined {
  if (!base || !top) return top ?? base;
  const conditions: Record<string, DocTableStylePart> = { ...base.conditions };
  for (const [k, v] of Object.entries(top.conditions ?? {})) conditions[k] = mergeParts(conditions[k], v) as DocTableStylePart;
  const borders = mergeBorders(base.borders, top.borders);
  const whole = mergeParts(base.whole, top.whole);
  const rowBandSize = top.rowBandSize ?? base.rowBandSize;
  const colBandSize = top.colBandSize ?? base.colBandSize;
  return {
    ...(borders && { borders }),
    ...(whole && { whole }),
    ...(Object.keys(conditions).length > 0 && { conditions }),
    ...(rowBandSize && { rowBandSize }),
    ...(colBandSize && { colBandSize }),
  };
}

export const pxFromTwips = (twips: number): number => Math.round(twips * TWIPS_TO_PX * 100) / 100;
