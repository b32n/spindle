// Group transforms. Members share a groupId (nested groups: see group-tree.ts); a group has no frame of its
// own — its bounds are the union of member AABBs. Members always hold absolute slide frames, so a nested group
// transforms as the flat list of its leaves; `childFrameInParent` composes a DrawingML group transform for import.
// Rotating a group orbits each member's centre about the group pivot and adds Δθ to each member's rotation.
// Scaling is lossy for rotated members under non-uniform factors, so the UI
// aspect-locks group corner handles (documented in the plan).

import type { Frame } from './types';
import { frameCenter, unionAABB, type Point, type Rect } from './geometry';
import type { FrameItem } from './align';

// The DrawingML group composition lives in shared (the PPTX reader composes group children with it too).
export { childFrameInParent, composeGroupChain, type FlippedFrame, type GroupXfrm } from '@b32nio/spindle-shared';

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
