import { adjustmentHandles, adjustmentsAfterDrag } from './adjustment-handles';
import { SHAPE_PRESETS, SHAPE_PRESET_INFO, resolveAdjustments, unknownShape } from './shape-presets';
import type { ShapePreset } from './types';

const BOXES: Array<[number, number]> = [[400, 200], [200, 400], [300, 300]];
// presets whose guides have no handle (kept honest here so a new preset is a conscious choice)
const NO_HANDLE_FOR_GUIDE: Partial<Record<ShapePreset, number[]>> = { blockArc: [2] };

describe('adjustmentHandles', () => {
  it('shows no handles for presets without guides, unknown shapes, and names outside the library', () => {
    expect(adjustmentHandles('rect', 100, 50)).toEqual([]);
    expect(adjustmentHandles(unknownShape('wave'), 100, 50, [0.3])).toEqual([]);
    expect(adjustmentHandles('nonsense' as never, 100, 50)).toEqual([]);
    expect(adjustmentsAfterDrag('rect', 0, { x: 1, y: 1 }, 100, 50)).toBeUndefined();
    expect(adjustmentsAfterDrag(unknownShape('wave'), 0, { x: 1, y: 1 }, 100, 50)).toBeUndefined();
  });

  it('gives every preset that has guides a handle for each of its guides (except the listed few)', () => {
    for (const p of SHAPE_PRESETS) {
      const adj = SHAPE_PRESET_INFO[p].adj;
      if (!adj) continue;
      const handles = adjustmentHandles(p, 400, 200);
      expect(handles.length).toBeGreaterThan(0);
      expect(handles.map((h) => h.id)).toEqual(handles.map((_, i) => i));
      const skipped = NO_HANDLE_FOR_GUIDE[p]?.length ?? 0;
      expect(handles.length).toBeGreaterThanOrEqual(adj.length - skipped - (p.startsWith('wedge') ? 1 : 0));
    }
  });

  it('places every handle inside the box (callout tips excepted) at the defaults', () => {
    for (const p of SHAPE_PRESETS) {
      if (p.startsWith('wedge')) continue;
      for (const [w, h] of BOXES) {
        for (const hd of adjustmentHandles(p, w, h)) {
          expect(hd.x).toBeGreaterThanOrEqual(0);
          expect(hd.x).toBeLessThanOrEqual(w);
          expect(hd.y).toBeGreaterThanOrEqual(0);
          expect(hd.y).toBeLessThanOrEqual(h);
        }
      }
    }
  });

  it('is a round trip: dragging a handle onto its own position keeps the guides (within the clamp ranges)', () => {
    for (const p of SHAPE_PRESETS) {
      for (const [w, h] of BOXES) {
        const base = resolveAdjustments(p);
        for (const hd of adjustmentHandles(p, w, h)) {
          const next = adjustmentsAfterDrag(p, hd.id, { x: hd.x, y: hd.y }, w, h);
          expect(next).toBeDefined();
          // a position clamped to the box (arrow heads on a narrow box) can differ; the defaults never are
          next!.forEach((v, i) => expect(v).toBeCloseTo(base[i], 5));
        }
      }
    }
  });

  it('maps a drag onto the guide it moves and clamps to the preset range', () => {
    // roundRect: x = a * min(w, h)
    expect(adjustmentsAfterDrag('roundRect', 0, { x: 40, y: 0 }, 400, 200)).toEqual([0.2]);
    expect(adjustmentsAfterDrag('roundRect', 0, { x: 9999, y: 0 }, 400, 200)).toEqual([0.5]);
    expect(adjustmentsAfterDrag('roundRect', 0, { x: -50, y: 0 }, 400, 200)).toEqual([0]);
    // rightArrow: handle 0 is the shaft thickness (vertical), handle 1 the head length
    const arrow = adjustmentsAfterDrag('rightArrow', 0, { x: 300, y: 0 }, 400, 200)!;
    expect(arrow[0]).toBeCloseTo(1);
    expect(arrow[1]).toBe(0.5);
    expect(adjustmentsAfterDrag('rightArrow', 1, { x: 300, y: 0 }, 400, 200, [0.2, 0.5])).toEqual([0.2, 0.5]);
    // left/up/down and the double arrows
    expect(adjustmentsAfterDrag('leftArrow', 1, { x: 100, y: 0 }, 400, 200)![1]).toBeCloseTo(0.5);
    expect(adjustmentsAfterDrag('upArrow', 0, { x: 300, y: 100 }, 400, 200)![0]).toBeCloseTo(0.5);
    expect(adjustmentsAfterDrag('downArrow', 1, { x: 0, y: 100 }, 400, 200)![1]).toBeCloseTo(0.5);
  });

  it('moves both tip guides with a callout handle and keeps the radius guide', () => {
    const next = adjustmentsAfterDrag('wedgeRoundRectCallout', 0, { x: 100, y: 250 }, 400, 200)!;
    expect(next[0]).toBeCloseTo(-0.25);
    expect(next[1]).toBeCloseTo(0.75);
    expect(next[2]).toBeCloseTo(1 / 6);
    expect(adjustmentsAfterDrag('wedgeRectCallout', 0, { x: 99999, y: -99999 }, 400, 200)).toEqual([2, -1]);
    expect(adjustmentHandles('wedgeRoundRectCallout', 400, 200).length).toBe(2);
  });

  it('maps angle handles of pie, chord and block arc through the ellipse', () => {
    const [start, end] = adjustmentHandles('pie', 400, 200);
    expect(start.x).toBeCloseTo(400);
    expect(start.y).toBeCloseTo(100);
    expect(end.x).toBeCloseTo(200);
    expect(end.y).toBeCloseTo(0);
    // a point straight above the centre is three quarters of the way round
    expect(adjustmentsAfterDrag('pie', 0, { x: 200, y: 0 }, 400, 200)![0]).toBeCloseTo(0.75);
    // the angle wraps into [0, 1)
    expect(adjustmentsAfterDrag('chord', 1, { x: 200, y: 50 }, 400, 200)![1]).toBeCloseTo(0.75);
    expect(adjustmentsAfterDrag('blockArc', 0, { x: 0, y: 100 }, 400, 200)![0]).toBeCloseTo(0.5);
    expect(adjustmentsAfterDrag('blockArc', 1, { x: 400, y: 100 }, 400, 200)![1]).toBeCloseTo(0);
  });

  it('keeps other guides when one moves, and honours the star6 horizontal stretch', () => {
    expect(adjustmentsAfterDrag('pie', 0, { x: 200, y: 0 }, 400, 200, [0, 0.4])![1]).toBeCloseTo(0.4);
    const [hd] = adjustmentHandles('star6', 300, 300);
    expect(hd.x).toBeCloseTo(150 + 300 * 0.28868 * (2 / Math.sqrt(3)));
    expect(adjustmentsAfterDrag('star6', 0, { x: hd.x, y: hd.y }, 300, 300)![0]).toBeCloseTo(0.28868);
  });

  it('does not divide by a zero-sized box', () => {
    expect(adjustmentsAfterDrag('roundRect', 0, { x: 1, y: 1 }, 0, 10)).toBeUndefined();
    expect(adjustmentsAfterDrag('pie', 0, { x: 1, y: 1 }, 0, 10)).toBeUndefined();
    expect(adjustmentsAfterDrag('wedgeRectCallout', 0, { x: 1, y: 1 }, 10, 0)).toBeUndefined();
  });

  it('rejects an unknown handle id', () => {
    expect(adjustmentsAfterDrag('roundRect', 3, { x: 1, y: 1 }, 100, 100)).toBeUndefined();
    expect(adjustmentsAfterDrag('roundRect', -1, { x: 1, y: 1 }, 100, 100)).toBeUndefined();
  });
});
