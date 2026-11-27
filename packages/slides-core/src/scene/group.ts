// Group transforms. Members share a groupId (nested groups: see group-tree.ts); a group has no frame of its
// own — its bounds are the union of member AABBs. Members always hold absolute slide frames, so a nested group
// transforms as the flat list of its leaves; `childFrameInParent` composes a DrawingML group transform for import.
// Rotating a group orbits each member's centre about the group pivot and adds Δθ to each member's rotation.
// Scaling is lossy for rotated members under non-uniform factors, so the UI
// aspect-locks group corner handles (documented in the plan).

import type { Frame } from './types';
import { frameCenter, unionAABB, type Point, type Rect } from './geometry';
import type { FrameItem } from './align';

export function groupAABB(frames: Frame[]): Rect | null {
  return unionAABB(frames);
}

function rotatePointDeg(p: Point, pivot: Point, deg: number): Point {
  const rad = (deg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const dx = p.x - pivot.x;
  const dy = p.y - pivot.y;
  return { x: pivot.x + dx * cos - dy * sin, y: pivot.y + dx * sin + dy * cos };
}

/** Rotate a group by `deltaDeg` about `pivot` (defaults to the group centre). */
export function rotateGroup(items: FrameItem[], deltaDeg: number, pivot?: Point): FrameItem[] {
  const bounds = unionAABB(items.map((it) => it.frame));
  const p = pivot ?? (bounds ? { x: bounds.x + bounds.w / 2, y: bounds.y + bounds.h / 2 } : { x: 0, y: 0 });
  return items.map(({ id, frame }) => {
    const c = frameCenter(frame);
    const nc = rotatePointDeg(c, p, deltaDeg);
    return {
      id,
      frame: {
        ...frame,
        x: nc.x - frame.w / 2,
        y: nc.y - frame.h / 2,
        rotation: frame.rotation + deltaDeg,
      },
    };
  });
}

/**
 * Scale a group about `origin` by (sx, sy): member centres scale relative to
 * the origin and member w/h scale by the same factors. Use sx === sy (corner
 * handles) to stay lossless for rotated members.
 */
export function scaleGroup(items: FrameItem[], sx: number, sy: number, origin: Point): FrameItem[] {
  return items.map(({ id, frame }) => {
    const c = frameCenter(frame);
    const ncx = origin.x + (c.x - origin.x) * sx;
    const ncy = origin.y + (c.y - origin.y) * sy;
    const w = Math.max(1, frame.w * sx);
    const h = Math.max(1, frame.h * sy);
    return { id, frame: { ...frame, x: ncx - w / 2, y: ncy - h / 2, w, h } };
  });
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
export type FlippedFrame = Frame & { flipH: boolean; flipV: boolean };

/** ext/chExt, or 1 when either extent is missing, zero, negative or not finite (malformed files). */
const extentScale = (ext: number, chExt: number): number =>
  Number.isFinite(ext) && Number.isFinite(chExt) && chExt > 0 && ext >= 0 ? ext / chExt : 1;

const normDeg = (d: number): number => ((d % 360) + 360) % 360;

/**
 * Map a frame given in a group's child space to the group's parent space: scale by ext/chExt about the child-space
 * centre, mirror for the group flips, rotate about the group centre. A flip mirrors the child (its flip toggles and
 * its rotation changes sign once per flipped axis). A zero, negative or non-finite extent maps with scale 1.
 */
export function childFrameInParent(child: Frame & { flipH?: boolean; flipV?: boolean }, g: GroupXfrm): FlippedFrame {
  const sx = extentScale(g.ext.w, g.chExt.w);
  const sy = extentScale(g.ext.h, g.chExt.h);
  const c = frameCenter(child);
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
export function composeGroupChain(child: Frame & { flipH?: boolean; flipV?: boolean }, chain: GroupXfrm[]): FlippedFrame {
  let f: FlippedFrame = { ...child, flipH: child.flipH ?? false, flipV: child.flipV ?? false };
  for (const g of chain) f = childFrameInParent(f, g);
  return f;
}
