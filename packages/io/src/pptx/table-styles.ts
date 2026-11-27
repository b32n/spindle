import type { PptxCellBorders, PptxColor, PptxColorMap, PptxFill, PptxLevelStyle, PptxLine } from '@b32nio/spindle-shared/export/pptx';
import { kid, kids, type XmlElement } from '../xlsx/xml';
import { colorElement, parseColor } from './color';
import { parseLine, type Rels } from './fill';
import { pct } from './units';

// Table styles (plan K7 and K10): `ppt/tableStyles.xml` read into the pieces a cell resolves, so that the cell carries its own
// look (the editor has no style gallery). Corner cells (`nwCell` and friends) are not resolved.

type EdgeKey = 'left' | 'right' | 'top' | 'bottom' | 'insideH' | 'insideV';

/** What one region of a table style (`a:wholeTbl`, `a:band1H`, `a:firstRow`...) sets. */
export interface TablePart {
  fill?: PptxFill;
  /** A `null` edge is an explicit no-line. */
  edges: Partial<Record<EdgeKey, PptxLine | null>>;
  text: Pick<PptxLevelStyle, 'bold' | 'italic' | 'color'>;
}

export interface TableStyle {
  id: string;
  name: string;
  parts: Partial<Record<'wholeTbl' | 'band1H' | 'band2H' | 'band1V' | 'band2V' | 'firstCol' | 'lastCol' | 'firstRow' | 'lastRow', TablePart>>;
}

const PART_NAMES = ['wholeTbl', 'band1H', 'band2H', 'band1V', 'band2V', 'firstCol', 'lastCol', 'firstRow', 'lastRow'] as const;
const EDGES: EdgeKey[] = ['left', 'right', 'top', 'bottom', 'insideH', 'insideV'];

/** A colour in a table style: a `tint` (the share of the colour kept, the rest being white) reads as the lighter variant `lumMod` / `lumOff` give. */
function styleColor(holder: XmlElement | undefined, map: PptxColorMap): PptxColor | undefined {
  const el = colorElement(holder);
  const color = parseColor(el, map);
  const tint = pct(kid(el, 'tint')?.attrs.val);
  if (color?.kind === 'theme' && tint !== undefined && color.lumMod === undefined) return { ...color, lumMod: tint, lumOff: 1 - tint };
  return color;
}

function readPart(el: XmlElement, map: PptxColorMap, rels: Rels): TablePart {
  const part: TablePart = { edges: {}, text: {} };
  const style = kid(el, 'tcStyle');
  const fill = kid(kid(style, 'fill'), 'solidFill');
  if (fill) {
    const color = styleColor(fill, map);
    if (color) part.fill = { kind: 'solid', color };
  } else if (kid(kid(style, 'fill'), 'noFill')) part.fill = { kind: 'none' };
  const borders = kid(style, 'tcBdr');
  for (const key of EDGES) {
    const ln = kid(kid(borders, key), 'ln');
    if (!ln) continue;
    const line = parseLine(ln, map, rels) as PptxLine;
    const solid = ln.children.find((c) => c.name === 'solidFill');
    if (solid) line.fill = { kind: 'solid', color: styleColor(solid, map) ?? { kind: 'rgb', hex: '#000000' } };
    part.edges[key] = line.fill?.kind === 'none' ? null : line;
  }
  const tx = kid(el, 'tcTxStyle');
  if (tx?.attrs.b === 'on') part.text.bold = true;
  else if (tx?.attrs.b === 'off') part.text.bold = false;
  if (tx?.attrs.i === 'on') part.text.italic = true;
  else if (tx?.attrs.i === 'off') part.text.italic = false;
  const color = styleColor(tx, map);
  if (color) part.text.color = color;
  return part;
}

/** The styles of `ppt/tableStyles.xml`, by `styleId`. */
export function readTableStyles(root: XmlElement | undefined, map: PptxColorMap, rels: Rels): Map<string, TableStyle> {
  const out = new Map<string, TableStyle>();
  for (const s of kids(root, 'tblStyle')) {
    const id = s.attrs.styleId ?? '';
    const parts: TableStyle['parts'] = {};
    for (const name of PART_NAMES) {
      const el = kid(s, name);
      if (el) parts[name] = readPart(el, map, rels);
    }
    if (id !== '') out.set(id, { id, name: s.attrs.styleName ?? '', parts });
  }
  return out;
}

export interface CellPosition {
  row: number;
  col: number;
  rows: number;
  cols: number;
}

export interface TableFlags {
  firstRow: boolean;
  firstCol: boolean;
  lastRow: boolean;
  lastCol: boolean;
  bandRow: boolean;
  bandCol: boolean;
}

export interface ResolvedCell {
  fill?: PptxFill;
  borders: PptxCellBorders;
  text: TablePart['text'];
}

/** The style regions that apply to a cell, lowest priority first: the whole table, banding, the first and last column, the first and last row. */
function regions(style: TableStyle, at: CellPosition, flags: TableFlags): TablePart[] {
  const { parts } = style;
  const list: (TablePart | undefined)[] = [parts.wholeTbl];
  const rowBand = at.row - (flags.firstRow ? 1 : 0);
  const colBand = at.col - (flags.firstCol ? 1 : 0);
  const inBody = !(flags.firstRow && at.row === 0) && !(flags.lastRow && at.row === at.rows - 1);
  if (flags.bandRow && inBody) list.push(rowBand % 2 === 0 ? parts.band1H : parts.band2H);
  const inCols = !(flags.firstCol && at.col === 0) && !(flags.lastCol && at.col === at.cols - 1);
  if (flags.bandCol && inCols) list.push(colBand % 2 === 0 ? parts.band1V : parts.band2V);
  if (flags.firstCol && at.col === 0) list.push(parts.firstCol);
  if (flags.lastCol && at.col === at.cols - 1) list.push(parts.lastCol);
  if (flags.firstRow && at.row === 0) list.push(parts.firstRow);
  if (flags.lastRow && at.row === at.rows - 1) list.push(parts.lastRow);
  return list.filter((p): p is TablePart => p !== undefined);
}

/** What a table style gives one cell: each region states only what it sets and a higher one wins. An outer edge of the table takes the region's `left` / `top`..., an inner edge its `insideV` / `insideH`. */
export function resolveCellStyle(style: TableStyle, at: CellPosition, flags: TableFlags): ResolvedCell {
  const out: ResolvedCell = { borders: {}, text: {} };
  const edgeKey = {
    l: at.col === 0 ? 'left' : 'insideV',
    r: at.col === at.cols - 1 ? 'right' : 'insideV',
    t: at.row === 0 ? 'top' : 'insideH',
    b: at.row === at.rows - 1 ? 'bottom' : 'insideH',
  } as const;
  for (const part of regions(style, at, flags)) {
    if (part.fill) out.fill = part.fill;
    for (const side of ['l', 't', 'r', 'b'] as const) {
      const edge = part.edges[edgeKey[side]];
      if (edge !== undefined) out.borders[side] = edge;
    }
    Object.assign(out.text, part.text);
  }
  return out;
}
