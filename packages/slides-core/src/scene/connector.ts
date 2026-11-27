// Connector geometry — resolving a line's endpoints when they are bound to
// other elements' anchors, and reconstructing the line's box
// (x/y/w/h/flipH/flipV) from two endpoints. A line occupies its box diagonal;
// flipH/flipV pick which corner the *start* sits at, so all four start→end
// orderings map to a distinct (box, flipH, flipV) triple and the direction
// (hence which end carries the arrowhead) survives the round-trip. Pure + tested.

import type { AnchorId, ConnectionFrame, LineElement, ShapeKind, SlideElement } from './types';
import { frameCenter, rotatePoint, type Point } from './geometry';

export const ANCHOR_IDS: readonly AnchorId[] = ['n', 'e', 's', 'w', 'nw', 'ne', 'se', 'sw'];

type UV = { u: number; v: number };

/** Fractional (u, v) position of each anchor within the un-rotated box (the rectangle's connection sites). */
const ANCHOR_UV: Record<AnchorId, UV> = {
  nw: { u: 0, v: 0 }, n: { u: 0.5, v: 0 }, ne: { u: 1, v: 0 },
  w: { u: 0, v: 0.5 }, e: { u: 1, v: 0.5 },
  sw: { u: 0, v: 1 }, s: { u: 0.5, v: 1 }, se: { u: 1, v: 1 },
};

const EDGE_MIDS: readonly AnchorId[] = ['n', 'e', 's', 'w'];
const DIAG = 0.5 - Math.SQRT1_2 / 2; // an ellipse's 45 degree point sits this far in from the box corner

/** Where an ellipse's diagonal sites sit: on the outline, not at the box corners. */
const ELLIPSE_UV: Partial<Record<AnchorId, UV>> = {
  nw: { u: DIAG, v: DIAG }, ne: { u: 1 - DIAG, v: DIAG }, se: { u: 1 - DIAG, v: 1 - DIAG }, sw: { u: DIAG, v: 1 - DIAG },
};

type SiteMap = Partial<Record<AnchorId, UV>>;

const edgeMids = (): SiteMap => Object.fromEntries(EDGE_MIDS.map((a) => [a, ANCHOR_UV[a]]));
const ELLIPSE_SITES: SiteMap = { ...edgeMids(), ...ELLIPSE_UV };
/** Outlines with a point at each edge midpoint and nothing at the box corners (diamond, terminator, can). */
const FOUR_EDGE_SITES: SiteMap = edgeMids();
/** Isosceles triangle: apex, the midpoints of both slanted sides, and the three base points. */
const TRIANGLE_SITES: SiteMap = { n: { u: 0.5, v: 0 }, nw: { u: 0.25, v: 0.5 }, ne: { u: 0.75, v: 0.5 }, sw: { u: 0, v: 1 }, s: { u: 0.5, v: 1 }, se: { u: 1, v: 1 } };

/** Shapes whose connection sites differ from the rectangle's eight (OOXML `cxnLst`, mapped onto the eight anchor ids).
 * The triangle's apex is fixed at 0.5: `triangle` has no adjustment in `SHAPE_PRESET_INFO`, so these sites are exact. */
const SHAPE_SITES: Partial<Record<string, SiteMap>> = {
  ellipse: ELLIPSE_SITES,
  flowChartConnector: ELLIPSE_SITES,
  donut: ELLIPSE_SITES,
  noSmoking: ELLIPSE_SITES,
  smileyFace: ELLIPSE_SITES,
  wedgeEllipseCallout: ELLIPSE_SITES,
  diamond: FOUR_EDGE_SITES,
  flowChartDecision: FOUR_EDGE_SITES,
  flowChartTerminator: FOUR_EDGE_SITES,
  flowChartPreparation: FOUR_EDGE_SITES,
  flowChartDelay: FOUR_EDGE_SITES,
  can: FOUR_EDGE_SITES,
  triangle: TRIANGLE_SITES,
};

/** The anchors a shape offers as connection sites (all eight for a rectangle or an unlisted shape). */
export function connectionAnchors(shape?: ShapeKind): readonly AnchorId[] {
  const set = shape ? SHAPE_SITES[shape] : undefined;
  return set ? (Object.keys(set) as AnchorId[]) : ANCHOR_IDS;
}

function anchorUV(shape: ShapeKind | undefined, anchor: AnchorId): UV {
  const set = shape ? SHAPE_SITES[shape] : undefined;
  return set?.[anchor] ?? ANCHOR_UV[anchor];
}

/** Resolve an anchor on a (possibly rotated) frame to slide coordinates. A shape-specific site follows the outline. */
export function anchorPoint(f: ConnectionFrame, anchor: AnchorId): Point {
  const site = anchorUV(f.shape, anchor);
  const u = f.flipH ? 1 - site.u : site.u;
  const v = f.flipV ? 1 - site.v : site.v;
  const local = { x: f.x + u * f.w, y: f.y + v * f.h };
  return f.rotation ? rotatePoint(local, frameCenter(f), f.rotation) : local;
}

/** All connection sites of a frame, keyed by anchor id (screen order irrelevant). */
export function anchorPoints(f: ConnectionFrame): Array<{ anchor: AnchorId; point: Point }> {
  return connectionAnchors(f.shape).map((anchor) => ({ anchor, point: anchorPoint(f, anchor) }));
}

/** The frame of an element as a connector sees it: its box plus, for a shape, the preset that fixes its sites. */
export function connectionFrame(el: SlideElement): ConnectionFrame {
  const f: ConnectionFrame = { x: el.x, y: el.y, w: el.w, h: el.h, rotation: el.rotation };
  if (el.type === 'shape') f.shape = el.shape;
  if (el.type === 'shape' || el.type === 'image') {
    if (el.flipH) f.flipH = true;
    if (el.flipV) f.flipV = true;
  }
  return f;
}

/** The box corner an unbound start/end occupies, given the line's flipH/flipV. */
function freeStart(line: LineElement): Point {
  return {
    x: line.flipH ? line.x + line.w : line.x,
    y: line.flipV ? line.y + line.h : line.y,
  };
}
function freeEnd(line: LineElement): Point {
  return {
    x: line.flipH ? line.x : line.x + line.w,
    y: line.flipV ? line.y : line.y + line.h,
  };
}

export type FrameLookup = (elementId: string) => ConnectionFrame | undefined;

/**
 * Resolve a line's two endpoints in slide coordinates. Priority per end:
 *   1. bound → the target's anchor (via `getFrame`);
 *   2. an explicit free point (startPoint/endPoint) — a connector's free end;
 *   3. the box corner (a plain, unbound line).
 * A binding whose target is missing falls through to (2)/(3).
 */
export function resolveEndpoints(line: LineElement, getFrame: FrameLookup): { start: Point; end: Point } {
  const startFrame = line.startBind && getFrame(line.startBind.elementId);
  const endFrame = line.endBind && getFrame(line.endBind.elementId);
  const start = startFrame ? anchorPoint(startFrame, line.startBind!.anchor) : line.startPoint ?? freeStart(line);
  const end = endFrame ? anchorPoint(endFrame, line.endBind!.anchor) : line.endPoint ?? freeEnd(line);
  return { start, end };
}

/**
 * Reconstruct a line's axis-aligned box + diagonal orientation from two
 * endpoints, so the (start → end) diagonal renders through those exact points.
 */
export function connectorBox(start: Point, end: Point): { x: number; y: number; w: number; h: number; flipH: boolean; flipV: boolean } {
  const x = Math.min(start.x, end.x);
  const y = Math.min(start.y, end.y);
  const w = Math.abs(end.x - start.x);
  const h = Math.abs(end.y - start.y);
  // flipH/flipV record which corner the start occupies (right/bottom of end),
  // so freeStart/freeEnd rebuild the exact endpoints — including for axis-aligned
  // lines, where the single-sign approach used to collapse direction and land the
  // arrowhead on the wrong end.
  const flipH = start.x > end.x;
  const flipV = start.y > end.y;
  return { x, y, w, h, flipH, flipV };
}

/** Effective box for a (possibly bound) line, resolved against `getFrame`. */
export function resolveConnectorFrame(line: LineElement, getFrame: FrameLookup): { x: number; y: number; w: number; h: number; flipH: boolean; flipV: boolean } {
  const { start, end } = resolveEndpoints(line, getFrame);
  return connectorBox(start, end);
}

// ── Routing ─────────────────────────────────────────────────────────────────

/** A unit axis direction (one of four) in slide coordinates. */
export type Direction = Point;

/** How far a route leaves a shape before turning when its ends face away from each other. */
const STUB = 20;

const OUTWARD: Record<AnchorId, Point> = {
  n: { x: 0, y: -1 }, s: { x: 0, y: 1 }, e: { x: 1, y: 0 }, w: { x: -1, y: 0 },
  // Diagonal sites leave through a side: horizontally, like PowerPoint's corner sites on rounded shapes.
  nw: { x: -1, y: 0 }, sw: { x: -1, y: 0 }, ne: { x: 1, y: 0 }, se: { x: 1, y: 0 },
};

/** The axis direction a connector leaves a shape's anchor in, turned with the shape's rotation to the nearest axis. */
export function anchorDirection(f: ConnectionFrame, anchor: AnchorId): Direction {
  const base = OUTWARD[anchor];
  const o = { x: f.flipH ? -base.x : base.x, y: f.flipV ? -base.y : base.y };
  const rad = (f.rotation * Math.PI) / 180;
  const x = o.x * Math.cos(rad) - o.y * Math.sin(rad);
  const y = o.x * Math.sin(rad) + o.y * Math.cos(rad);
  return Math.abs(x) >= Math.abs(y) ? { x: Math.sign(x), y: 0 } : { x: 0, y: Math.sign(y) };
}

export type ConnectorRoute =
  | { kind: 'poly'; points: Point[] }
  | { kind: 'cubic'; points: [Point, Point, Point, Point] };

const eq = (a: Point, b: Point): boolean => Math.abs(a.x - b.x) < 1e-6 && Math.abs(a.y - b.y) < 1e-6;

/** Drop repeated points and points that sit in the middle of a straight run. */
function tidy(points: Point[]): Point[] {
  const out: Point[] = [];
  for (const p of points) if (!out.length || !eq(out[out.length - 1], p)) out.push(p);
  const res: Point[] = [];
  for (const p of out) {
    while (res.length >= 2) {
      const a = res[res.length - 2];
      const b = res[res.length - 1];
      // Only a point in the middle of a run goes; one where the route doubles back is a real turn.
      const through = (a.x === b.x && b.x === p.x && (b.y - a.y) * (p.y - b.y) > 0) || (a.y === b.y && b.y === p.y && (b.x - a.x) * (p.x - b.x) > 0);
      if (!through) break;
      res.pop();
    }
    res.push(p);
  }
  return res;
}

/** Axis (0 = x, 1 = y) a direction runs along. */
const axisOf = (d: Direction): 0 | 1 => (d.x !== 0 ? 0 : 1);
const comp = (p: Point, axis: 0 | 1): number => (axis === 0 ? p.x : p.y);
const mk = (a: number, b: number, axis: 0 | 1): Point => (axis === 0 ? { x: a, y: b } : { x: b, y: a });

/** An elbow route: right-angle legs from `s` (leaving along `d1`) to `e` (arriving against `d2`, which points out of the end shape). */
function elbow(s: Point, e: Point, d1: Direction, d2: Direction, adj: number): Point[] {
  const a1 = axisOf(d1);
  const a2 = axisOf(d2);
  const cross: 0 | 1 = a1 === 0 ? 1 : 0;
  const sgn = (n: number): number => Math.sign(n);
  if (a1 === a2) {
    const delta = comp(e, a1) - comp(s, a1);
    const facing = delta !== 0 && comp(d1, a1) === sgn(delta) && comp(d2, a1) === -sgn(delta);
    if (facing) {
      const mid = comp(s, a1) + delta * adj;
      return tidy([s, mk(mid, comp(s, cross), a1), mk(mid, comp(e, cross), a1), e]);
    }
    const p1 = { x: s.x + d1.x * STUB, y: s.y + d1.y * STUB };
    const p2 = { x: e.x + d2.x * STUB, y: e.y + d2.y * STUB };
    const mid = (comp(s, cross) + comp(e, cross)) / 2;
    return tidy([s, p1, mk(comp(p1, a1), mid, a1), mk(comp(p2, a1), mid, a1), p2, e]);
  }
  // Perpendicular ends: one corner when both legs head the right way, otherwise go round.
  const corner = mk(comp(e, a1), comp(s, cross), a1);
  const leg1 = comp(corner, a1) - comp(s, a1);
  const leg2 = comp(e, a2) - comp(corner, a2);
  if (comp(d1, a1) * leg1 >= 0 && comp(d2, a2) * leg2 <= 0) return tidy([s, corner, e]);
  const p1 = { x: s.x + d1.x * STUB, y: s.y + d1.y * STUB };
  const p2 = { x: e.x + d2.x * STUB, y: e.y + d2.y * STUB };
  return tidy([s, p1, mk(comp(p1, a1), comp(p2, cross), a1), p2, e]);
}

/**
 * Route a connector between two points. `d1`/`d2` are the directions the route leaves the start and the end shape in
 * (outward from the shape); a missing one (a free end) is chosen so the route runs straight at the other end.
 */
export function routeConnector(
  kind: 'straight' | 'elbow' | 'curved',
  start: Point,
  end: Point,
  d1?: Direction,
  d2?: Direction,
  adj = 0.5
): ConnectorRoute {
  if (kind === 'straight' || (start.x === end.x && start.y === end.y)) return { kind: 'poly', points: [start, end] };
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const along: Direction = Math.abs(dx) >= Math.abs(dy) ? { x: Math.sign(dx), y: 0 } : { x: 0, y: Math.sign(dy) };
  const out1 = d1 ?? along;
  const out2 = d2 ?? { x: -along.x, y: -along.y };
  if (kind === 'elbow') return { kind: 'poly', points: elbow(start, end, out1, out2, Math.min(1, Math.max(0, adj))) };
  const k = Math.max(STUB, Math.max(Math.abs(dx), Math.abs(dy)) / 2);
  return {
    kind: 'cubic',
    points: [start, { x: start.x + out1.x * k, y: start.y + out1.y * k }, { x: end.x + out2.x * k, y: end.y + out2.y * k }, end],
  };
}

/** A line's route given its resolved endpoints, leaving bound ends perpendicular to the bound side. */
export function lineRoute(line: LineElement, start: Point, end: Point, getFrame: FrameLookup): ConnectorRoute {
  const dirOf = (bind: LineElement['startBind']): Direction | undefined => {
    const f = bind && getFrame(bind.elementId);
    return f ? anchorDirection(f, bind!.anchor) : undefined;
  };
  return routeConnector(line.connector ?? 'straight', start, end, dirOf(line.startBind), dirOf(line.endBind), line.adjustments?.[0]);
}

/** The SVG path data of a route, shifted by (-ox, -oy) into box-local coordinates. */
export function routePathD(route: ConnectorRoute, ox = 0, oy = 0): string {
  const f = (p: Point): string => `${p.x - ox} ${p.y - oy}`;
  if (route.kind === 'cubic') {
    const [p0, c1, c2, p1] = route.points;
    return `M${f(p0)} C${f(c1)} ${f(c2)} ${f(p1)}`;
  }
  return route.points.map((p, i) => `${i ? 'L' : 'M'}${f(p)}`).join(' ');
}
