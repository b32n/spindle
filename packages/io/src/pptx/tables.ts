import type { PptxCellBorders, PptxLevelStyles, PptxTable, PptxTableCell, PptxTextBody } from '@b32nio/spindle-shared/export/pptx';
import { isTrue, kid, kids, type XmlElement } from '../xlsx/xml';
import { parseFill, parseLine } from './fill';
import { frameOf, mergeLevel } from './levels';
import { readTextBody } from './text';
import { type ShapeContext } from './shapes';
import { resolveCellStyle, type TableFlags, type TableStyle } from './table-styles';
import { emuToPx, num } from './units';

// Tables (plan K7 and K10): `a:tbl` in a `p:graphicFrame`, with spans, per-cell borders, margins, anchors and the table style
// resolved into each cell.

export interface TableContext extends ShapeContext {
  styles: ReadonlyMap<string, TableStyle>;
}

const TABLE_URI = /\/table$/;
const ANCHOR = { t: 'top', ctr: 'middle', b: 'bottom' } as const;
const SIDES = [['l', 'lnL'], ['t', 'lnT'], ['r', 'lnR'], ['b', 'lnB']] as const;
const MARGINS = [['l', 'marL'], ['t', 'marT'], ['r', 'marR'], ['b', 'marB']] as const;

/** Is this `p:graphicFrame` a table? */
export const tableOf = (frame: XmlElement): XmlElement | undefined => {
  const data = kid(kid(frame, 'graphic'), 'graphicData');
  return TABLE_URI.test(data?.attrs.uri ?? '') ? kid(data, 'tbl') : undefined;
};

const spanOf = (v: string | undefined): number | undefined => {
  const n = num(v);
  return n !== undefined && n >= 2 ? Math.floor(n) : undefined;
};

function readCell(tc: XmlElement, at: { row: number; col: number; rows: number; cols: number }, flags: TableFlags, style: TableStyle | undefined, ctx: TableContext): PptxTableCell {
  const pr = kid(tc, 'tcPr');
  const resolved = style ? resolveCellStyle(style, at, flags) : undefined;
  const textStyle = resolved?.text ?? {};
  const base: PptxLevelStyles = ctx.defaultLevels.map((l) => mergeLevel(l, textStyle));
  const body = kid(tc, 'txBody');
  const text: PptxTextBody = body ? readTextBody(body, base, undefined, ctx.map, undefined, { rels: ctx.rels, warn: ctx.warn }) : { props: {}, paragraphs: [] };

  const borders: PptxCellBorders = { ...resolved?.borders };
  for (const [side, name] of SIDES) {
    const ln = kid(pr, name);
    const line = ln ? parseLine(ln, ctx.map, ctx.rels) : undefined;
    if (line) borders[side] = line.fill?.kind === 'none' ? null : line;
  }
  const fill = parseFill(pr, ctx.map, ctx.rels) ?? resolved?.fill;
  const margins: NonNullable<PptxTableCell['margins']> = {};
  for (const [side, attr] of MARGINS) {
    const v = emuToPx(pr?.attrs[attr]);
    if (v !== undefined && v >= 0) margins[side] = v;
  }
  const anchor = ANCHOR[(pr?.attrs.anchor ?? '') as keyof typeof ANCHOR];
  const gridSpan = spanOf(tc.attrs.gridSpan);
  const rowSpan = spanOf(tc.attrs.rowSpan);
  return {
    text,
    ...(gridSpan ? { gridSpan } : {}),
    ...(rowSpan ? { rowSpan } : {}),
    ...(isTrue(tc.attrs.hMerge) ? { hMerge: true } : {}),
    ...(isTrue(tc.attrs.vMerge) ? { vMerge: true } : {}),
    ...(fill ? { fill } : {}),
    ...(Object.keys(borders).length > 0 ? { borders } : {}),
    ...(Object.keys(margins).length > 0 ? { margins } : {}),
    ...(anchor ? { anchor } : {}),
    ...(pr?.attrs.vert && pr.attrs.vert !== 'horz' ? { vert: pr.attrs.vert } : {}),
  };
}

export function readTable(frame: XmlElement, tbl: XmlElement, ctx: TableContext): PptxTable {
  const cNvPr = kid(kid(frame, 'nvGraphicFramePr'), 'cNvPr');
  const pr = kid(tbl, 'tblPr');
  const columns = kids(kid(tbl, 'tblGrid'), 'gridCol').map((c) => Math.max(0, emuToPx(c.attrs.w) ?? 0));
  const trs = kids(tbl, 'tr');
  const flags: TableFlags = {
    firstRow: isTrue(pr?.attrs.firstRow),
    firstCol: isTrue(pr?.attrs.firstCol),
    lastRow: isTrue(pr?.attrs.lastRow),
    lastCol: isTrue(pr?.attrs.lastCol),
    bandRow: isTrue(pr?.attrs.bandRow),
    bandCol: isTrue(pr?.attrs.bandCol),
  };
  const styleId = (kid(pr, 'tableStyleId')?.text ?? '').trim();
  const style = styleId !== '' ? ctx.styles.get(styleId) : undefined;
  if (styleId !== '' && !style) ctx.warn(`A table uses the style ${styleId}, which the file does not define; its cells keep only their own formatting.`);
  // the grid is rectangular: a row with fewer cells than the grid has columns is padded, extra cells are dropped
  const cols = Math.max(columns.length, ...trs.map((tr) => kids(tr, 'tc').length));
  while (columns.length < cols) columns.push(0);
  const rows = trs.map((tr, row) => ({
    height: Math.max(0, emuToPx(tr.attrs.h) ?? 0),
    cells: Array.from({ length: cols }, (_, col) => {
      const tc = kids(tr, 'tc')[col];
      return readCell(tc ?? { name: 'tc', ns: 'main', attrs: {}, children: [], text: '' }, { row, col, rows: trs.length, cols }, flags, style, ctx);
    }),
  }));
  const altText = (cNvPr?.attrs.descr ?? '').trim();
  const frameBox = frameOf(kid(frame, 'xfrm'));
  const flagsOn = Object.fromEntries(Object.entries(flags).filter(([, on]) => on).map(([k]) => [k, true as const]));
  return {
    kind: 'table',
    pptxId: cNvPr?.attrs.id ?? '',
    name: (cNvPr?.attrs.name ?? '').trim(),
    ...(altText !== '' ? { altText } : {}),
    ...(isTrue(cNvPr?.attrs.hidden) ? { hidden: true } : {}),
    ...(frameBox ? { frame: frameBox } : {}),
    columns,
    rows,
    ...(Object.keys(flagsOn).length > 0 ? { flags: flagsOn } : {}),
    ...(styleId !== '' ? { styleId } : {}),
    ...(style?.name ? { styleName: style.name } : {}),
  };
}
