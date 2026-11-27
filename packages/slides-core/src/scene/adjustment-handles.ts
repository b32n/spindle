// Canvas adjustment handles (K5b): where each preset's yellow-diamond style handles sit for the current guide values,
// and the guide value a dragged handle position stands for. Positions are in the shape's own box (origin top-left,
// before rotation and flips); the guide semantics mirror the geometry in slides-react (`presetPath`), so a handle
// always sits on the feature it moves. Presets without an entry here simply show no handles.

import { isShapePreset, isUnknownShape, resolveAdjustments } from './shape-presets';
import type { Point } from './geometry';
import type { ShapeKind, ShapePreset } from './types';

export interface AdjustmentHandle {
  /** Which of the shape's handles this is (stable for a preset; passed back to `adjustmentsAfterDrag`). */
  id: number;
  x: number;
  y: number;
}

interface HandleDef {
  /** The handle position for the resolved guide values in a `w` x `h` box. */
  at(w: number, h: number, adj: readonly number[]): Point;
  /** The raw (unclamped) guide values a box position stands for, as `[index, value]` pairs; none when the box is too small to tell. */
  from(p: Point, w: number, h: number): Array<[number, number]>;
}

type Len = (w: number, h: number) => number;
const W: Len = (w) => w;
const H: Len = (_w, h) => h;
const SS: Len = (w, h) => Math.min(w, h);
type Cross = (w: number, h: number, adj: readonly number[]) => number;

const clampTo = (v: number, max: number): number => Math.min(Math.max(v, 0), max);

/**
 * A handle that slides along one axis: the position is `origin + sign * adj * factor * len`, the other coordinate
 * comes from `cross`. The position stays inside the box (the geometry clamps its features there too).
 */
function along(
  index: number,
  axis: 'x' | 'y',
  origin: 'start' | 'mid' | 'end',
  sign: 1 | -1,
  len: Len,
  factor: number,
  cross: Cross,
): HandleDef {
  const base = (w: number, h: number): number => {
    const extent = axis === 'x' ? w : h;
    return origin === 'start' ? 0 : origin === 'mid' ? extent / 2 : extent;
  };
  return {
    at(w, h, adj) {
      const pos = clampTo(base(w, h) + sign * adj[index] * factor * len(w, h), axis === 'x' ? w : h);
      const other = cross(w, h, adj);
      return axis === 'x' ? { x: pos, y: other } : { x: other, y: pos };
    },
    from(p, w, h) {
      const scale = factor * len(w, h);
      if (scale <= 0) return [];
      return [[index, (sign * ((axis === 'x' ? p.x : p.y) - base(w, h))) / scale]];
    },
  };
}

/** A handle on the box-inscribed ellipse at `adj[index]` turns of a full circle, clockwise from three o'clock. */
function onEllipseAt(index: number): HandleDef {
  return {
    at(w, h, adj) {
      const t = adj[index] * 2 * Math.PI;
      return { x: w / 2 + (w / 2) * Math.cos(t), y: h / 2 + (h / 2) * Math.sin(t) };
    },
    from(p, w, h) {
      if (w <= 0 || h <= 0) return [];
      const turns = Math.atan2((p.y - h / 2) / (h / 2), (p.x - w / 2) / (w / 2)) / (2 * Math.PI);
      return [[index, ((turns % 1) + 1) % 1]];
    },
  };
}

const at = (n: number): Cross => () => n;
const top = at(0);
const centreX: Cross = (w) => w / 2;
const centreY: Cross = (_w, h) => h / 2;
const rightEdge: Cross = (w) => w;
const bottomEdge: Cross = (_w, h) => h;

/** A callout tail tip: guides 0 and 1 are its offset from the centre in box fractions, so one handle drives both. */
const tip: HandleDef = {
  at: (w, h, adj) => ({ x: w / 2 + adj[0] * w, y: h / 2 + adj[1] * h }),
  from: (p, w, h) => (w > 0 && h > 0 ? [[0, (p.x - w / 2) / w], [1, (p.y - h / 2) / h]] : []),
};

const HANDLES: Partial<Record<ShapePreset, HandleDef[]>> = {
  roundRect: [along(0, 'x', 'start', 1, SS, 1, top)],
  snip1Rect: [along(0, 'x', 'end', -1, SS, 1, top)],
  snip2SameRect: [along(0, 'x', 'start', 1, SS, 1, top), along(1, 'x', 'start', 1, SS, 1, bottomEdge)],
  snip2DiagRect: [along(0, 'x', 'start', 1, SS, 1, top), along(1, 'x', 'end', -1, SS, 1, top)],
  snipRoundRect: [along(0, 'x', 'start', 1, SS, 1, top), along(1, 'x', 'end', -1, SS, 1, top)],
  round1Rect: [along(0, 'x', 'end', -1, SS, 1, top)],
  round2SameRect: [along(0, 'x', 'end', -1, SS, 1, top), along(1, 'x', 'end', -1, SS, 1, bottomEdge)],
  round2DiagRect: [along(0, 'x', 'start', 1, SS, 1, top), along(1, 'x', 'end', -1, SS, 1, top)],
  foldedCorner: [along(0, 'x', 'end', -1, SS, 1, bottomEdge)],
  frame: [along(0, 'x', 'start', 1, SS, 1, top)],
  plaque: [along(0, 'x', 'start', 1, SS, 1, top)],
  bevel: [along(0, 'x', 'start', 1, SS, 1, top)],
  can: [along(0, 'y', 'start', 1, SS, 0.5, centreX)],
  cube: [along(0, 'y', 'start', 1, SS, 1, at(0))],
  donut: [along(0, 'x', 'start', 1, SS, 1, centreY)],
  noSmoking: [along(0, 'x', 'start', 1, SS, 1, centreY)],
  teardrop: [along(0, 'x', 'mid', 1, W, 0.5, (_w, h, adj) => h / 2 - (h / 2) * adj[0])],
  pie: [onEllipseAt(0), onEllipseAt(1)],
  chord: [onEllipseAt(0), onEllipseAt(1)],
  blockArc: [onEllipseAt(0), onEllipseAt(1)],
  moon: [along(0, 'x', 'end', -1, W, 1, centreY)],

  rightArrow: [
    along(0, 'y', 'mid', -1, H, 0.5, (w, h, adj) => w - Math.min(SS(w, h) * adj[1], w)),
    along(1, 'x', 'end', -1, SS, 1, top),
  ],
  leftArrow: [
    along(0, 'y', 'mid', -1, H, 0.5, (w, h, adj) => Math.min(SS(w, h) * adj[1], w)),
    along(1, 'x', 'start', 1, SS, 1, top),
  ],
  upArrow: [
    along(0, 'x', 'mid', 1, W, 0.5, (w, h, adj) => Math.min(SS(w, h) * adj[1], h)),
    along(1, 'y', 'start', 1, SS, 1, at(0)),
  ],
  downArrow: [
    along(0, 'x', 'mid', 1, W, 0.5, (w, h, adj) => h - Math.min(SS(w, h) * adj[1], h)),
    along(1, 'y', 'end', -1, SS, 1, at(0)),
  ],
  leftRightArrow: [
    along(0, 'y', 'mid', -1, H, 0.5, (w, h, adj) => Math.min(SS(w, h) * adj[1], w / 2)),
    along(1, 'x', 'start', 1, SS, 1, top),
  ],
  upDownArrow: [
    along(0, 'x', 'mid', 1, W, 0.5, centreY),
    along(1, 'y', 'start', 1, SS, 1, at(0)),
  ],
  homePlate: [along(0, 'x', 'end', -1, SS, 1, top)],

  star4: [along(0, 'x', 'mid', 1, W, 1, centreY)],
  star6: [along(0, 'x', 'mid', 1, W, 2 / Math.sqrt(3), centreY)],
  star8: [along(0, 'x', 'mid', 1, W, 1, centreY)],
  star12: [along(0, 'x', 'mid', 1, W, 1, centreY)],

  wedgeRectCallout: [tip],
  wedgeRoundRectCallout: [tip, along(2, 'x', 'start', 1, SS, 1, top)],
  wedgeEllipseCallout: [tip],

  leftBracket: [along(0, 'y', 'start', 1, SS, 1, rightEdge)],
  rightBracket: [along(0, 'y', 'start', 1, SS, 1, at(0))],
  leftBrace: [along(0, 'y', 'start', 1, SS, 1, rightEdge), along(1, 'y', 'start', 1, H, 1, at(0))],
  rightBrace: [along(0, 'y', 'start', 1, SS, 1, at(0)), along(1, 'y', 'start', 1, H, 1, rightEdge)],
  bracketPair: [along(0, 'x', 'start', 1, SS, 1, top)],

  mathPlus: [along(0, 'x', 'mid', 1, SS, 0.5, (_w, h) => h / 2 - 0.36745 * h)],
  mathMinus: [along(0, 'y', 'mid', -1, H, 0.5, centreX)],
};

/** The handles to show for a shape in a `w` x `h` box with these guide values (none for presets without adjustable features and for unknown shapes). */
export function adjustmentHandles(shape: ShapeKind, w: number, h: number, adjustments?: readonly number[]): AdjustmentHandle[] {
  const defs = handleDefs(shape);
  if (!defs.length) return [];
  const adj = resolveAdjustments(shape, adjustments);
  return defs.map((def, id) => ({ id, ...def.at(w, h, adj) }));
}

/**
 * The full guide list after dragging handle `id` to `p` (in the shape's box, unrotated): the guides that handle drives
 * take the value `p` stands for (clamped to their range), the rest keep their current value. Undefined when there is
 * no such handle or the box is too small to tell.
 */
export function adjustmentsAfterDrag(shape: ShapeKind, id: number, p: Point, w: number, h: number, adjustments?: readonly number[]): number[] | undefined {
  const def = handleDefs(shape)[id];
  if (!def) return undefined;
  const set = def.from(p, w, h);
  if (!set.length) return undefined;
  const next = resolveAdjustments(shape, adjustments);
  for (const [index, value] of set) next[index] = value;
  return resolveAdjustments(shape, next);
}

function handleDefs(shape: ShapeKind): readonly HandleDef[] {
  return !isUnknownShape(shape) && isShapePreset(shape) ? (HANDLES[shape] ?? []) : [];
}
