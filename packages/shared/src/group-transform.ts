// Group transforms (DrawingML `a:xfrm` of a `p:grpSp`), pure maths on frames. They live here, below the editor packages, because
// both the PPTX reader (`@b32nio/spindle-io/pptx`, which composes a group's children into slide space) and slides-core (which
// re-exports them from `scene/group.ts`) need the same composition.

interface Point {
  x: number;
  y: number;
}

/** A rectangle with its rotation (degrees, clockwise, about its centre), as slides-core's `Frame`. */
export interface GroupFrame {
  x: number;
  y: number;
  w: number;
  h: number;
  rotation: number;
}

/** A DrawingML group transform (`a:xfrm` of a `p:grpSp`): the group box, the child coordinate space it maps, rotation and flips. */
export interface GroupXfrm {
  off: Point;
  ext: { w: number; h: number };
  chOff: Point;
  chExt: { w: number; h: number };
  /** Degrees, clockwise, about the group centre. */
  rotation: number;
  flipH?: boolean;
  flipV?: boolean;
}

/** A child frame with its flips. */
export type FlippedFrame = GroupFrame & { flipH: boolean; flipV: boolean };

function rotatePointDeg(p: Point, pivot: Point, deg: number): Point {
  const rad = (deg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const dx = p.x - pivot.x;
  const dy = p.y - pivot.y;
  return { x: pivot.x + dx * cos - dy * sin, y: pivot.y + dx * sin + dy * cos };
}

/** ext/chExt, or 1 when chExt is zero or negative or either extent is not finite (malformed files); a zero ext collapses to scale 0. */
const extentScale = (ext: number, chExt: number): number =>
  Number.isFinite(ext) && Number.isFinite(chExt) && chExt > 0 && ext >= 0 ? ext / chExt : 1;

const normDeg = (d: number): number => ((d % 360) + 360) % 360;

/**
 * Map a frame given in a group's child space to the group's parent space: scale by ext/chExt about the child-space
 * centre, mirror for the group flips, rotate about the group centre. A flip mirrors the child (its flip toggles and
 * its rotation changes sign once per flipped axis). A zero or negative chExt or a non-finite extent maps with scale 1;
 * a zero ext collapses the children onto the group centre. Known limit: a rotated child under non-uniform scale only has
 * its w and h scaled (a 90° child gets the wrong extent), like the lossy case noted for `scaleGroup`.
 */
export function childFrameInParent(child: GroupFrame & { flipH?: boolean; flipV?: boolean }, g: GroupXfrm): FlippedFrame {
  const sx = extentScale(g.ext.w, g.chExt.w);
  const sy = extentScale(g.ext.h, g.chExt.h);
  const c = { x: child.x + child.w / 2, y: child.y + child.h / 2 };
  const chCx = g.chOff.x + g.chExt.w / 2;
  const chCy = g.chOff.y + g.chExt.h / 2;
  const d = { x: (c.x - chCx) * sx * (g.flipH ? -1 : 1), y: (c.y - chCy) * sy * (g.flipV ? -1 : 1) };
  const n = rotatePointDeg({ x: g.off.x + g.ext.w / 2 + d.x, y: g.off.y + g.ext.h / 2 + d.y }, { x: g.off.x + g.ext.w / 2, y: g.off.y + g.ext.h / 2 }, g.rotation);
  const w = child.w * sx;
  const h = child.h * sy;
  const sign = (g.flipH ? -1 : 1) * (g.flipV ? -1 : 1);
  return {
    x: n.x - w / 2,
    y: n.y - h / 2,
    w,
    h,
    rotation: normDeg(g.rotation + sign * child.rotation),
    flipH: (child.flipH ?? false) !== (g.flipH ?? false),
    flipV: (child.flipV ?? false) !== (g.flipV ?? false),
  };
}

/** Compose a child frame through nested group transforms, innermost group first, to slide space. */
export function composeGroupChain(child: GroupFrame & { flipH?: boolean; flipV?: boolean }, chain: GroupXfrm[]): FlippedFrame {
  let f: FlippedFrame = { ...child, flipH: child.flipH ?? false, flipV: child.flipV ?? false };
  for (const g of chain) f = childFrameInParent(f, g);
  return f;
}
