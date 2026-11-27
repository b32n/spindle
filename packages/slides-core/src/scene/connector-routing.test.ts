import {
  anchorDirection,
  anchorPoint,
  anchorPoints,
  connectionAnchors,
  connectionFrame,
  lineRoute,
  resolveEndpoints,
  routeConnector,
  routePathD,
  ANCHOR_IDS,
} from './connector';
import { createLineElement, createShapeElement, createTextElement } from './elements';
import type { AnchorId, ConnectionFrame, LineElement } from './types';
import type { Point } from './geometry';

const f = (x: number, y: number, w: number, h: number, shape?: ConnectionFrame['shape'], rotation = 0): ConnectionFrame => ({ x, y, w, h, rotation, ...(shape ? { shape } : {}) });
const pts = (r: ReturnType<typeof routeConnector>): Point[] => r.points;
const isOrthogonal = (p: Point[]): boolean => p.every((q, i) => i === 0 || q.x === p[i - 1].x || q.y === p[i - 1].y);
const E = { x: 1, y: 0 }, W = { x: -1, y: 0 }, N = { x: 0, y: -1 }, S = { x: 0, y: 1 };

describe('shape-specific connection sites', () => {
  it('a rectangle, an unlisted preset and no shape offer the eight anchors', () => {
    expect(connectionAnchors()).toEqual(ANCHOR_IDS);
    expect(connectionAnchors('rect')).toEqual(ANCHOR_IDS);
    expect(connectionAnchors('star6')).toEqual(ANCHOR_IDS);
    expect(connectionAnchors('unknown:wave')).toEqual(ANCHOR_IDS);
    expect(anchorPoints(f(0, 0, 100, 100, 'rect'))).toHaveLength(8);
  });

  it('an ellipse puts its diagonal sites on the outline', () => {
    const k = 0.5 - Math.SQRT1_2 / 2;
    const p = anchorPoint(f(0, 0, 200, 100, 'ellipse'), 'nw');
    expect(p.x).toBeCloseTo(200 * k);
    expect(p.y).toBeCloseTo(100 * k);
    expect(anchorPoint(f(0, 0, 200, 100, 'ellipse'), 'se').x).toBeCloseTo(200 * (1 - k));
    expect(anchorPoint(f(0, 0, 200, 100, 'ellipse'), 'n')).toEqual({ x: 100, y: 0 });
    expect(connectionAnchors('ellipse')).toHaveLength(8);
  });

  it('a diamond offers four edge midpoints and no corners; a bound corner falls back to the box corner', () => {
    expect(connectionAnchors('diamond').slice().sort()).toEqual(['e', 'n', 's', 'w']);
    expect(connectionAnchors('flowChartDecision')).toHaveLength(4);
    expect(anchorPoints(f(0, 0, 100, 60, 'diamond')).map((a) => a.anchor).sort()).toEqual(['e', 'n', 's', 'w']);
    expect(anchorPoint(f(0, 0, 100, 60, 'diamond'), 'ne')).toEqual({ x: 100, y: 0 });
  });

  it('a triangle has its apex, slanted-side midpoints and base points', () => {
    const t = f(0, 0, 100, 80, 'triangle');
    expect(connectionAnchors('triangle').slice().sort()).toEqual(['n', 'ne', 'nw', 's', 'se', 'sw']);
    expect(anchorPoint(t, 'n')).toEqual({ x: 50, y: 0 });
    expect(anchorPoint(t, 'nw')).toEqual({ x: 25, y: 40 });
    expect(anchorPoint(t, 'ne')).toEqual({ x: 75, y: 40 });
    expect(anchorPoint(t, 'sw')).toEqual({ x: 0, y: 80 });
  });

  it('sites rotate with the frame', () => {
    const p = anchorPoint(f(0, 0, 100, 100, 'ellipse', 180), 'nw');
    const k = 0.5 - Math.SQRT1_2 / 2;
    expect(p.x).toBeCloseTo(100 * (1 - k));
    expect(p.y).toBeCloseTo(100 * (1 - k));
  });

  it('connectionFrame carries the preset of shapes only', () => {
    const shape = createShapeElement({ containerId: 's', index: 'a', shape: 'triangle', x: 1, y: 2, w: 3, h: 4 });
    expect(connectionFrame(shape)).toEqual({ x: 1, y: 2, w: 3, h: 4, rotation: 0, shape: 'triangle' });
    const text = createTextElement({ containerId: 's', index: 'a', x: 1, y: 2, w: 3, h: 4 });
    expect(connectionFrame(text)).toEqual({ x: 1, y: 2, w: 3, h: 4, rotation: 0 });
  });

  it('a bound line resolves to the shape site through the frame lookup', () => {
    const line = createLineElement({ containerId: 's', index: 'a', startBind: { elementId: 'a', anchor: 'nw' }, endPoint: { x: 0, y: 0 } });
    const { start } = resolveEndpoints(line, () => f(0, 0, 100, 80, 'triangle'));
    expect(start).toEqual({ x: 25, y: 40 });
  });
});

describe('anchorDirection', () => {
  const box = f(0, 0, 100, 100);
  it.each<[AnchorId, { x: number; y: number }]>([
    ['n', N], ['s', S], ['e', E], ['w', W], ['nw', W], ['sw', W], ['ne', E], ['se', E],
  ])('%s leaves outward', (a, d) => expect(anchorDirection(box, a)).toEqual(d));

  it('turns with the rotation to the nearest axis', () => {
    expect(anchorDirection({ ...box, rotation: 90 }, 'n')).toEqual(E);
    expect(anchorDirection({ ...box, rotation: 180 }, 'e')).toEqual(W);
    expect(anchorDirection({ ...box, rotation: 270 }, 'e')).toEqual(N);
    expect(anchorDirection({ ...box, rotation: 100 }, 'n')).toEqual(E);
    expect(anchorDirection({ ...box, rotation: 10 }, 'e')).toEqual(E);
  });
});

describe('routeConnector', () => {
  it('straight is the segment itself', () => {
    expect(routeConnector('straight', { x: 0, y: 0 }, { x: 5, y: 5 })).toEqual({ kind: 'poly', points: [{ x: 0, y: 0 }, { x: 5, y: 5 }] });
  });

  it('elbow facing the same axis: three legs with the middle at adj', () => {
    const r = routeConnector('elbow', { x: 0, y: 0 }, { x: 100, y: 50 }, E, W);
    expect(pts(r)).toEqual([{ x: 0, y: 0 }, { x: 50, y: 0 }, { x: 50, y: 50 }, { x: 100, y: 50 }]);
    expect(pts(routeConnector('elbow', { x: 0, y: 0 }, { x: 100, y: 50 }, E, W, 0.25))[1]).toEqual({ x: 25, y: 0 });
    // adj is clamped to the span
    expect(pts(routeConnector('elbow', { x: 0, y: 0 }, { x: 100, y: 50 }, E, W, 7))[1]).toEqual({ x: 100, y: 0 });
    expect(pts(routeConnector('elbow', { x: 0, y: 0 }, { x: 100, y: 50 }, E, W, -3))).toEqual([{ x: 0, y: 0 }, { x: 0, y: 50 }, { x: 100, y: 50 }]);
  });

  it('elbow facing the same axis vertically', () => {
    const r = routeConnector('elbow', { x: 0, y: 0 }, { x: 40, y: 100 }, S, N);
    expect(pts(r)).toEqual([{ x: 0, y: 0 }, { x: 0, y: 50 }, { x: 40, y: 50 }, { x: 40, y: 100 }]);
  });

  it('aligned ends collapse to a straight run', () => {
    expect(pts(routeConnector('elbow', { x: 0, y: 10 }, { x: 100, y: 10 }, E, W))).toEqual([{ x: 0, y: 10 }, { x: 100, y: 10 }]);
  });

  it('elbow turning away goes round with stubs', () => {
    // both ends open to the right but the end is to the left of the start
    const r = pts(routeConnector('elbow', { x: 100, y: 0 }, { x: 0, y: 60 }, E, E));
    expect(r[0]).toEqual({ x: 100, y: 0 });
    expect(r[r.length - 1]).toEqual({ x: 0, y: 60 });
    expect(isOrthogonal(r)).toBe(true);
    expect(r[1]).toEqual({ x: 120, y: 0 }); // leaves east first
    expect(r[r.length - 2]).toEqual({ x: 20, y: 60 }); // arrives from the east
  });

  it('keeps the turns where a route doubles back on itself', () => {
    expect(pts(routeConnector('elbow', { x: 0, y: 0 }, { x: -100, y: 0 }, E, E))).toEqual([{ x: 0, y: 0 }, { x: 20, y: 0 }, { x: -100, y: 0 }]);
    expect(pts(routeConnector('elbow', { x: 0, y: 0 }, { x: 0, y: -50 }, S, S))).toEqual([{ x: 0, y: 0 }, { x: 0, y: 20 }, { x: 0, y: -50 }]);
  });

  it('elbow perpendicular ends: one corner when both legs head the right way', () => {
    expect(pts(routeConnector('elbow', { x: 0, y: 0 }, { x: 100, y: 80 }, E, N))).toEqual([{ x: 0, y: 0 }, { x: 100, y: 0 }, { x: 100, y: 80 }]);
    expect(pts(routeConnector('elbow', { x: 0, y: 0 }, { x: 100, y: 80 }, S, W))).toEqual([{ x: 0, y: 0 }, { x: 0, y: 80 }, { x: 100, y: 80 }]);
  });

  it('elbow perpendicular ends that face the wrong way go round', () => {
    const r = pts(routeConnector('elbow', { x: 0, y: 0 }, { x: 100, y: 80 }, W, N));
    expect(isOrthogonal(r)).toBe(true);
    expect(r).toEqual([{ x: 0, y: 0 }, { x: -20, y: 0 }, { x: -20, y: 60 }, { x: 100, y: 60 }, { x: 100, y: 80 }]);
    const v = pts(routeConnector('elbow', { x: 0, y: 0 }, { x: 100, y: 80 }, N, W));
    expect(isOrthogonal(v)).toBe(true);
    expect(v).toEqual([{ x: 0, y: 0 }, { x: 0, y: -20 }, { x: 80, y: -20 }, { x: 80, y: 80 }, { x: 100, y: 80 }]);
    // only the end is wrong: s leaves east, e must be entered from below
    expect(pts(routeConnector('elbow', { x: 0, y: 0 }, { x: 100, y: 80 }, E, S))).toEqual([{ x: 0, y: 0 }, { x: 20, y: 0 }, { x: 20, y: 100 }, { x: 100, y: 100 }, { x: 100, y: 80 }]);
  });

  it('a free end is routed straight at the other end', () => {
    expect(pts(routeConnector('elbow', { x: 0, y: 0 }, { x: 100, y: 50 }))).toEqual([{ x: 0, y: 0 }, { x: 50, y: 0 }, { x: 50, y: 50 }, { x: 100, y: 50 }]);
    expect(pts(routeConnector('elbow', { x: 0, y: 0 }, { x: 10, y: 50 }))).toEqual([{ x: 0, y: 0 }, { x: 0, y: 25 }, { x: 10, y: 25 }, { x: 10, y: 50 }]);
    // coincident ends have no direction: still a valid route
    expect(pts(routeConnector('elbow', { x: 5, y: 5 }, { x: 5, y: 5 }))).toEqual([{ x: 5, y: 5 }, { x: 5, y: 5 }]);
  });

  it('curved leaves along the start direction and arrives against the end direction', () => {
    const r = routeConnector('curved', { x: 0, y: 0 }, { x: 100, y: 50 }, E, W);
    expect(r).toEqual({ kind: 'cubic', points: [{ x: 0, y: 0 }, { x: 50, y: 0 }, { x: 50, y: 50 }, { x: 100, y: 50 }] });
    const short = routeConnector('curved', { x: 0, y: 0 }, { x: 4, y: 2 }, E, W);
    expect(short.points[1]).toEqual({ x: 20, y: 0 });
  });
});

describe('routePathD and lineRoute', () => {
  it('writes polylines and cubics, shifted into box-local coordinates', () => {
    expect(routePathD({ kind: 'poly', points: [{ x: 10, y: 20 }, { x: 30, y: 20 }, { x: 30, y: 50 }] }, 10, 20)).toBe('M0 0 L20 0 L20 30');
    expect(routePathD({ kind: 'cubic', points: [{ x: 0, y: 0 }, { x: 1, y: 0 }, { x: 2, y: 3 }, { x: 4, y: 5 }] })).toBe('M0 0 C1 0 2 3 4 5');
  });

  const line = (over: Partial<LineElement>): LineElement => createLineElement({ containerId: 's', index: 'a', ...over });
  const frames: Record<string, ConnectionFrame> = { a: f(0, 0, 100, 100), b: f(300, 200, 100, 100) };
  const lookup = (id: string) => frames[id];

  it('leaves bound ends perpendicular to the bound side', () => {
    const l = line({ connector: 'elbow', startBind: { elementId: 'a', anchor: 'e' }, endBind: { elementId: 'b', anchor: 'w' } });
    const { start, end } = resolveEndpoints(l, lookup);
    expect(pts(lineRoute(l, start, end, lookup))).toEqual([{ x: 100, y: 50 }, { x: 200, y: 50 }, { x: 200, y: 250 }, { x: 300, y: 250 }]);
  });

  it('treats a missing target and a plain line as free ends', () => {
    const l = line({ connector: 'curved', startBind: { elementId: 'gone', anchor: 'e' }, endPoint: { x: 50, y: 0 } });
    expect(lineRoute(l, { x: 0, y: 0 }, { x: 50, y: 0 }, lookup).kind).toBe('cubic');
    expect(lineRoute(line({}), { x: 0, y: 0 }, { x: 50, y: 0 }, lookup)).toEqual({ kind: 'poly', points: [{ x: 0, y: 0 }, { x: 50, y: 0 }] });
  });

  it('reads the adjustment as the elbow position', () => {
    const l = line({ connector: 'elbow', adjustments: [0.2], startBind: { elementId: 'a', anchor: 'e' }, endBind: { elementId: 'b', anchor: 'w' } });
    const { start, end } = resolveEndpoints(l, lookup);
    expect(pts(lineRoute(l, start, end, lookup))[1]).toEqual({ x: 140, y: 50 });
  });
});

describe('createLineElement connector fields', () => {
  it('keeps elbow, curved and adjustments; drops straight and empty adjustments', () => {
    expect(createLineElement({ containerId: 's', index: 'a', connector: 'elbow', adjustments: [0.3] })).toMatchObject({ connector: 'elbow', adjustments: [0.3] });
    expect(createLineElement({ containerId: 's', index: 'a', connector: 'curved' }).connector).toBe('curved');
    const plain = createLineElement({ containerId: 's', index: 'a', connector: 'straight', adjustments: [] });
    expect(plain).not.toHaveProperty('connector');
    expect(plain).not.toHaveProperty('adjustments');
  });
});

describe('deck keeps a bound connector on a shape-specific site', () => {
  it('follows an ellipse site when the shape moves, and survives deleting it', async () => {
    const { DeckImpl } = await import('../deck');
    const d = new DeckImpl('room', 'Deck');
    const s = d.getActiveSlideId();
    const a = d.addElement(s, { type: 'shape', shape: 'ellipse', x: 0, y: 0, w: 100, h: 100 }).id;
    const l = d.addElement(s, { type: 'line', connector: 'elbow', startBind: { elementId: a, anchor: 'se' }, endPoint: { x: 300, y: 300 }, x: 0, y: 0, w: 1, h: 1 }).id;
    const k = 0.5 - Math.SQRT1_2 / 2;
    d.updateElement(a, { x: 10 });
    const el = d.getElement(l) as LineElement;
    expect(el.x).toBeCloseTo(10 + 100 * (1 - k));
    expect(el.connector).toBe('elbow');
    d.deleteElements([a]);
    expect((d.getElement(l) as LineElement).startBind).toBeUndefined();
    expect((d.getElement(l) as LineElement).startPoint?.x).toBeCloseTo(10 + 100 * (1 - k));
  });
});
