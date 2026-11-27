import { composeGroupChain, type GroupFrame, type GroupXfrm } from '@b32nio/spindle-shared';
import type { PptxElement, PptxFrame, PptxGroup } from '@b32nio/spindle-shared/export/pptx';
import { isTrue, kid, type XmlElement } from '../xlsx/xml';
import { readConnector } from './connectors';
import { parseFill } from './fill';
import { frameOf } from './levels';
import { readPicture, type PictureContext } from './pictures';
import { readShape, treeChildren, type ShapeContext } from './shapes';
import { readTable, tableOf, type TableContext } from './tables';
import { emuToPx } from './units';

// The elements of a slide's shape tree in drawing order (plan K9c, K10): `p:sp` shapes, `p:pic` pictures, the table
// `p:graphicFrame`s, `p:cxnSp` connectors and `p:grpSp` groups (K10b). A `p:graphicFrame` of any other kind (chart, SmartArt, OLE: K11) is not read yet.

export type ElementContext = ShapeContext & PictureContext & TableContext;

/** How deep groups are followed: a deeper group is dropped with a warning (a legitimate deck never nests this far). */
export const MAX_GROUP_DEPTH = 64;

const toFrame = (f: PptxFrame): GroupFrame & { flipH: boolean; flipV: boolean } => ({ x: f.x, y: f.y, w: f.w, h: f.h, rotation: f.rot ?? 0, flipH: f.flipH ?? false, flipV: f.flipV ?? false });

/** A frame of a group's child space mapped to slide space through `chain` (innermost group first). */
function composed(f: PptxFrame, chain: GroupXfrm[]): PptxFrame {
  if (chain.length === 0) return f;
  const c = composeGroupChain(toFrame(f), chain);
  return { x: c.x, y: c.y, w: c.w, h: c.h, ...(c.rotation !== 0 ? { rot: c.rotation } : {}), ...(c.flipH ? { flipH: true } : {}), ...(c.flipV ? { flipV: true } : {}) };
}

/** The group transform of a `p:grpSp` (`a:xfrm` of `p:grpSpPr`) in px; `undefined` when it has none or its box is unreadable. A missing child space (`a:chOff`, `a:chExt`) is the group's own box, so the children map to themselves. */
function groupTransform(xfrm: XmlElement | undefined): { own: PptxFrame; xfrm: GroupXfrm } | undefined {
  const own = frameOf(xfrm);
  if (!own) return undefined;
  const chOff = kid(xfrm, 'chOff');
  const chExt = kid(xfrm, 'chExt');
  const [x, y, w, h] = [emuToPx(chOff?.attrs.x), emuToPx(chOff?.attrs.y), emuToPx(chExt?.attrs.cx), emuToPx(chExt?.attrs.cy)];
  return {
    own,
    xfrm: {
      off: { x: own.x, y: own.y },
      ext: { w: own.w, h: own.h },
      chOff: { x: x ?? own.x, y: y ?? own.y },
      chExt: { w: w ?? own.w, h: h ?? own.h },
      rotation: own.rot ?? 0,
      ...(own.flipH ? { flipH: true } : {}),
      ...(own.flipV ? { flipV: true } : {}),
    },
  };
}

export async function readElements(tree: XmlElement | undefined, ctx: ElementContext, chain: GroupXfrm[] = [], depth = 0): Promise<PptxElement[]> {
  const out: PptxElement[] = [];
  const place = <T extends { frame?: PptxFrame }>(el: T): T => (el.frame ? { ...el, frame: composed(el.frame, chain) } : el);
  for (const c of treeChildren(tree)) {
    if (c.name === 'sp') out.push(place(readShape(c, ctx)));
    else if (c.name === 'pic') out.push(place(await readPicture(c, ctx)));
    else if (c.name === 'cxnSp') out.push(place(readConnector(c, ctx)));
    else if (c.name === 'grpSp') {
      const group = await readGroup(c, ctx, chain, depth);
      if (group) out.push(group);
    } else if (c.name === 'graphicFrame') {
      const tbl = tableOf(c);
      if (tbl) out.push(place(readTable(c, tbl, ctx)));
    }
  }
  return out;
}

/** A `p:grpSp`: its own box and its children, read in the child space of the group and composed into slide space. */
async function readGroup(e: XmlElement, ctx: ElementContext, chain: GroupXfrm[], depth: number): Promise<PptxGroup | undefined> {
  if (depth >= MAX_GROUP_DEPTH) {
    ctx.warn(`A group is nested more than ${MAX_GROUP_DEPTH} levels deep; it and what it holds are dropped.`);
    return undefined;
  }
  const cNvPr = kid(kid(e, 'nvGrpSpPr'), 'cNvPr');
  const grpSpPr = kid(e, 'grpSpPr');
  const t = groupTransform(kid(grpSpPr, 'xfrm'));
  // a group without a readable transform maps its children to themselves
  const inner = t ? [t.xfrm, ...chain] : chain;
  // a `grpFill` inside takes the fill of this group; a group's own `grpFill` is the fill of the group around it
  const groupFill = kid(grpSpPr, 'grpFill') ? ctx.groupFill : parseFill(grpSpPr, ctx.map, ctx.rels);
  const { groupFill: _outer, ...rest } = ctx;
  const children = await readElements(e, { ...rest, ...(groupFill ? { groupFill } : {}) }, inner, depth + 1);
  const altText = (cNvPr?.attrs.descr ?? '').trim();
  return {
    kind: 'group',
    pptxId: cNvPr?.attrs.id ?? '',
    name: (cNvPr?.attrs.name ?? '').trim(),
    ...(altText !== '' ? { altText } : {}),
    ...(isTrue(cNvPr?.attrs.hidden) ? { hidden: true } : {}),
    ...(t ? { frame: composed(t.own, chain) } : {}),
    children,
  };
}
