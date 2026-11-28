import type { PptxConnection, PptxConnector, PptxElement, PptxFrame, PptxGroup, PptxOpaque, PptxPicture, PptxTable, PptxTableCell } from '@b32nio/spindle-shared/export/pptx';
import { hlinkXml, type WriteContext } from './writer-context';
import { txBodyXml } from './writer-body';
import { opaqueXml } from './writer-opaque';
import { fillXml, geometryXml, lineXml, shapeXml } from './writer-shapes';
import { angle, attrs, emu, esc, xfrmXml } from './writer-xml';

// The slide elements other than `p:sp` (plan K13): pictures, tables, connectors and groups, and the dispatcher that writes a list
// of elements in drawing order. The inverses of `readPicture`, `readTable`, `readConnector` and `readGroup`. Opaque elements
// (charts, diagrams, objects, media, ink) are written by `writer-opaque.ts` (K14).

const CELL_SIDES = [['l', 'a:lnL'], ['r', 'a:lnR'], ['t', 'a:lnT'], ['b', 'a:lnB']] as const;
const MARGINS = [['l', 'marL'], ['r', 'marR'], ['t', 'marT'], ['b', 'marB']] as const;
const ANCHORS = { top: 't', middle: 'ctr', bottom: 'b' } as const;
const TABLE_FLAGS = ['firstRow', 'firstCol', 'lastRow', 'lastCol', 'bandRow', 'bandCol'] as const;
const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';

/** What writing a list of elements needs besides the context: the ids chosen for them and, for connectors, the new id of each source id. */
export interface ElementIds {
  ids: Map<PptxElement, number>;
  /** The new id of the element a source `pptxId` named (the first, when a malformed deck repeats one). */
  bySource: Map<string, number>;
}

/**
 * The ids of `allocateElementIds`, with the lookup connectors need to glue their ends to the shapes' new ids. Only elements that
 * are written have an entry: opaque elements the store holds no parts for and pictures with no image to embed leave a connector end free, not dangling.
 */
export function elementIds(ids: Map<PptxElement, number>, ctx: WriteContext): ElementIds {
  const bySource = new Map<string, number>();
  for (const [e, id] of ids) {
    const written = e.kind === 'opaque' ? ctx.opaque?.place(e) !== undefined : e.kind === 'picture' ? e.dataUrl !== undefined && ctx.image(e.dataUrl, e.part) !== undefined : true;
    if (written && !bySource.has(e.pptxId)) bySource.set(e.pptxId, id);
  }
  return { ids, bySource };
}

/** `p:cNvPr` with the element's name, alternative text and visibility, and a hyperlink inside it when there is one. */
function cNvPr(e: { name: string; altText?: string; hidden?: boolean; link?: Parameters<typeof hlinkXml>[0] }, id: number, ctx: WriteContext): string {
  const a = attrs({ id, name: e.name, descr: e.altText, hidden: e.hidden ? 1 : undefined });
  const link = hlinkXml(e.link, ctx);
  return link === '' ? `<p:cNvPr${a}/>` : `<p:cNvPr${a}>${link}</p:cNvPr>`;
}

/** `p:pic`: the image through a media part, the crop, the transparency, the frame, geometry and outline. A picture with no image to embed is left out. */
function pictureXml(pic: PptxPicture, id: number, ctx: WriteContext): string {
  const target = pic.dataUrl !== undefined ? ctx.image(pic.dataUrl, pic.part) : undefined;
  if (target === undefined) {
    ctx.lossy('picture');
    return '';
  }
  const rId = ctx.rels.add('image', target);
  const ph = pic.placeholder;
  const crop = pic.crop ? `<a:srcRect${attrs({ l: pctOf(pic.crop.l), t: pctOf(pic.crop.t), r: pctOf(pic.crop.r), b: pctOf(pic.crop.b) })}/>` : '';
  const alpha = pic.alpha !== undefined ? `<a:alphaModFix amt="${Math.round(pic.alpha * 100000)}"/>` : '';
  const spPr = (pic.frame ? xfrmXml(pic.frame) : '') + geometryXml(pic.geometry, ctx) + lineXml(pic.line, ctx);
  return (
    `<p:pic><p:nvPicPr>${cNvPr(pic, id, ctx)}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
    `<p:nvPr>${ph ? `<p:ph${attrs({ type: ph.type === 'obj' ? undefined : ph.type, idx: ph.idx })}/>` : ''}</p:nvPr></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rId}">${alpha}</a:blip>${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${spPr}</p:spPr></p:pic>`
  );
}

/** Percent of the source (`25`) to the thousandths of a percent of `a:srcRect` (`25000`), omitted when 0; a negative value (the crop grown outward) is written as it is. */
const pctOf = (v: number): number | undefined => (v !== 0 ? Math.round(v * 1000) : undefined);

function cellXml(cell: PptxTableCell, ctx: WriteContext): string {
  const borders = CELL_SIDES.map(([side, tag]) => {
    const b = cell.borders?.[side];
    return b === undefined ? '' : b === null ? `<${tag}><a:noFill/></${tag}>` : lineXml(b, ctx, tag);
  }).join('');
  const a = attrs({
    ...Object.fromEntries(MARGINS.map(([side, name]) => [name, cell.margins?.[side] !== undefined ? emu(cell.margins[side] as number) : undefined])),
    vert: cell.vert,
    anchor: cell.anchor ? ANCHORS[cell.anchor] : undefined,
  });
  const body = borders + fillXml(cell.fill, ctx);
  const span = attrs({ gridSpan: cell.gridSpan, rowSpan: cell.rowSpan, hMerge: cell.hMerge ? 1 : undefined, vMerge: cell.vMerge ? 1 : undefined });
  return `<a:tc${span}><a:txBody>${txBodyXml(cell.text, ctx)}</a:txBody>${body === '' ? `<a:tcPr${a}/>` : `<a:tcPr${a}>${body}</a:tcPr>`}</a:tc>`;
}

/** A table as a `p:graphicFrame` holding `a:tbl`; its style is named by id and registered with the context, the cells state their resolved formatting themselves. */
function tableXml(t: PptxTable, id: number, ctx: WriteContext): string {
  const flags = Object.fromEntries(TABLE_FLAGS.map((f) => [f, t.flags?.[f] ? 1 : undefined]));
  if (t.styleId !== undefined) ctx.tableStyle(t.styleId, t.styleName);
  const style = t.styleId !== undefined ? `<a:tableStyleId>${esc(t.styleId)}</a:tableStyleId>` : '';
  const grid = t.columns.map((w) => `<a:gridCol w="${emu(w)}"/>`).join('');
  const rows = t.rows.map((r) => `<a:tr h="${emu(r.height)}">${r.cells.map((c) => cellXml(c, ctx)).join('')}</a:tr>`).join('');
  const frame = t.frame ?? { x: 0, y: 0, w: t.columns.reduce((s, w) => s + w, 0), h: t.rows.reduce((s, r) => s + r.height, 0) };
  return (
    `<p:graphicFrame><p:nvGraphicFramePr>${cNvPr(t, id, ctx)}<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${emu(frame.x)}" y="${emu(frame.y)}"/><a:ext cx="${emu(frame.w)}" cy="${emu(frame.h)}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${TABLE_URI}"><a:tbl><a:tblPr${attrs(flags)}>${style}</a:tblPr><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  );
}

/** `a:stCxn` or `a:endCxn`: the end is glued to the shape's new id; a shape the slide does not hold leaves the end free. */
function connectionXml(tag: string, c: PptxConnection | undefined, bySource: Map<string, number>): string {
  const id = c ? bySource.get(c.id) : undefined;
  return c && id !== undefined ? `<a:${tag}${attrs({ id, idx: c.idx })}/>` : '';
}

function connectorXml(c: PptxConnector, id: number, ctx: WriteContext, bySource: Map<string, number>): string {
  const spPr = (c.frame ? xfrmXml(c.frame) : '') + geometryXml(c.geometry, ctx) + lineXml(c.line, ctx);
  const cxn = connectionXml('stCxn', c.start, bySource) + connectionXml('endCxn', c.end, bySource);
  return `<p:cxnSp><p:nvCxnSpPr>${cNvPr(c, id, ctx)}<p:cNvCxnSpPr>${cxn}</p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr>${spPr}</p:spPr></p:cxnSp>`;
}

const normDeg = (d: number): number => ((d % 360) + 360) % 360;

/**
 * A frame of slide space (or of the space of the group around) in the child space of the group `g`, which is written with its child
 * space equal to its own box (scale 1, so a child needs no scaling): the inverse of `childFrameInParent` for that case. The reader
 * composes the group's rotation and flips back into the child, so each group carries them and its children are un-rotated here.
 */
export function toChildSpace(f: PptxFrame, g: PptxFrame): PptxFrame {
  const gc = { x: g.x + g.w / 2, y: g.y + g.h / 2 };
  const rad = (-(g.rot ?? 0) * Math.PI) / 180;
  const dx = f.x + f.w / 2 - gc.x;
  const dy = f.y + f.h / 2 - gc.y;
  const ux = dx * Math.cos(rad) - dy * Math.sin(rad);
  const uy = dx * Math.sin(rad) + dy * Math.cos(rad);
  const cx = gc.x + (g.flipH ? -ux : ux);
  const cy = gc.y + (g.flipV ? -uy : uy);
  const sign = (g.flipH ? -1 : 1) * (g.flipV ? -1 : 1);
  const rot = sign * ((f.rot ?? 0) - (g.rot ?? 0));
  const flipH = (f.flipH ?? false) !== (g.flipH ?? false);
  const flipV = (f.flipV ?? false) !== (g.flipV ?? false);
  const r = rot === 0 ? 0 : normDeg(rot);
  return { x: cx - f.w / 2, y: cy - f.h / 2, w: f.w, h: f.h, ...(r !== 0 ? { rot: r } : {}), ...(flipH ? { flipH } : {}), ...(flipV ? { flipV } : {}) };
}

/** `p:grpSp`: the group's own box as its transform with the child space equal to it, then its children in that space. A group with no frame states no transform, so its children keep the space they are in. */
function groupXml(g: PptxGroup, frame: PptxFrame | undefined, id: number, ctx: WriteContext, all: ElementIds, chain: PptxFrame[]): string {
  const xfrm = frame
    ? `<a:xfrm${attrs({ rot: frame.rot ? angle(frame.rot) : undefined, flipH: frame.flipH ? 1 : undefined, flipV: frame.flipV ? 1 : undefined })}>` +
      `<a:off x="${emu(frame.x)}" y="${emu(frame.y)}"/><a:ext cx="${emu(frame.w)}" cy="${emu(frame.h)}"/><a:chOff x="${emu(frame.x)}" y="${emu(frame.y)}"/><a:chExt cx="${emu(frame.w)}" cy="${emu(frame.h)}"/></a:xfrm>`
    : '';
  return `<p:grpSp><p:nvGrpSpPr>${cNvPr(g, id, ctx)}<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${xfrm}</p:grpSpPr>${elementsXml(g.children, ctx, all, frame ? [...chain, frame] : chain)}</p:grpSp>`;
}

/**
 * The elements of a slide in drawing order: shapes, pictures, tables, connectors, groups and opaque content. The frames of the model are in slide space; `chain` holds the groups the list is
 * in, outermost first, each frame as written (in the child space of the one before it), and the frames are mapped through them.
 */
export function elementsXml(elements: PptxElement[], ctx: WriteContext, all: ElementIds, chain: PptxFrame[] = []): string {
  return elements
    .map((e) => {
      const id = all.ids.get(e) as number;
      const frame = e.frame && chain.length > 0 ? chain.reduce(toChildSpace, e.frame) : e.frame;
      const placed = frame === e.frame ? e : { ...e, frame };
      switch (placed.kind) {
        case 'shape':
          return shapeXml(placed, id, ctx);
        case 'picture':
          return pictureXml(placed, id, ctx);
        case 'table':
          return tableXml(placed, id, ctx);
        case 'connector':
          return connectorXml(placed, id, ctx, all.bySource);
        case 'group':
          return groupXml(e as PptxGroup, placed.frame, id, ctx, all, chain);
        case 'opaque':
          return opaqueXml(e as PptxOpaque, placed.frame, id, ctx);
      }
    })
    .join('');
}
