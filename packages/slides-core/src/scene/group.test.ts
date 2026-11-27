import { groupAABB, rotateGroup, scaleGroup, childFrameInParent, composeGroupChain, type GroupXfrm } from './group';
import { frameCenter } from './geometry';
import type { FrameItem } from './align';
import type { Frame } from './types';

const item = (id: string, x: number, y: number, w = 40, h = 40, rotation = 0): FrameItem => ({
  id,
  frame: { x, y, w, h, rotation } as Frame,
});

describe('group', () => {
  it('computes the group AABB', () => {
    expect(groupAABB([item('a', 0, 0).frame, item('b', 60, 60).frame])).toEqual({ x: 0, y: 0, w: 100, h: 100 });
  });

  it('rotating a group orbits member centres and adds Δθ', () => {
    const items = [item('a', 0, 0, 40, 40), item('b', 160, 160, 40, 40)];
    const pivot = { x: 100, y: 100 };
    const res = rotateGroup(items, 180, pivot);
    // 180° about (100,100): centre (20,20) → (180,180); (180,180) → (20,20).
    expect(frameCenter(res[0].frame).x).toBeCloseTo(180);
    expect(frameCenter(res[0].frame).y).toBeCloseTo(180);
    expect(res[0].frame.rotation).toBe(180);
    expect(frameCenter(res[1].frame).x).toBeCloseTo(20);
  });

  it('rotates about the group centre by default', () => {
    const res = rotateGroup([item('a', 0, 0, 40, 40), item('b', 160, 160, 40, 40)], 180);
    expect(frameCenter(res[0].frame).x).toBeCloseTo(180);
  });

  it('rotating nothing is a no-op', () => {
    expect(rotateGroup([], 45)).toEqual([]);
  });

  it('scaling a group scales centres about the origin and sizes', () => {
    const items = [item('a', 100, 100, 40, 40)];
    const res = scaleGroup(items, 2, 2, { x: 0, y: 0 });
    // Centre (120,120) → (240,240); size 40 → 80.
    expect(res[0].frame.w).toBe(80);
    expect(res[0].frame.h).toBe(80);
    expect(frameCenter(res[0].frame).x).toBeCloseTo(240);
    expect(frameCenter(res[0].frame).y).toBeCloseTo(240);
  });

  describe('group transforms (import composition)', () => {
    const base: GroupXfrm = { off: { x: 100, y: 100 }, ext: { w: 200, h: 100 }, chOff: { x: 0, y: 0 }, chExt: { w: 200, h: 100 }, rotation: 0 };
    const child = { x: 10, y: 20, w: 40, h: 30, rotation: 0 };

    it('maps through a translation-only group', () => {
      expect(childFrameInParent(child, base)).toEqual({ x: 110, y: 120, w: 40, h: 30, rotation: 0, flipH: false, flipV: false });
    });

    it('scales when the child extent differs from the group extent', () => {
      const g = { ...base, chExt: { w: 100, h: 50 }, chOff: { x: 10, y: 10 } };
      const f = childFrameInParent({ x: 10, y: 10, w: 50, h: 25, rotation: 0 }, g);
      expect(f).toMatchObject({ x: 100, y: 100, w: 100, h: 50 });
    });

    it('treats a zero child extent as scale 1', () => {
      const f = childFrameInParent(child, { ...base, chExt: { w: 0, h: 0 } });
      expect(f).toMatchObject({ w: 40, h: 30 });
    });

    it('collapses the children onto the group centre when the group ext is zero', () => {
      const f = childFrameInParent(child, { ...base, ext: { w: 0, h: 0 } });
      expect(f).toMatchObject({ x: 100, y: 100, w: 0, h: 0 });
    });

    it.each([
      ['negative child extent', { chExt: { w: -10, h: -10 } }],
      ['NaN child extent', { chExt: { w: NaN, h: NaN } }],
      ['negative group extent', { ext: { w: -5, h: -5 } }],
      ['NaN group extent', { ext: { w: NaN, h: NaN } }],
      ['infinite group extent', { ext: { w: Infinity, h: Infinity } }],
    ])('treats a %s as scale 1', (_n, over) => {
      const f = childFrameInParent(child, { ...base, ...over });
      expect(f).toMatchObject({ w: 40, h: 30 });
    });

    it('rotates children about the group centre and adds the group rotation', () => {
      // Group centre (200,150); child centre (130,135) is (-70,-15) from it; 180 degrees gives (+70,+15).
      const f = childFrameInParent({ ...child, rotation: 30 }, { ...base, rotation: 180 });
      expect(f.x + f.w / 2).toBeCloseTo(270);
      expect(f.y + f.h / 2).toBeCloseTo(165);
      expect(f.rotation).toBeCloseTo(210);
    });

    it('mirrors for group flips: positions reflect, flips toggle, rotation changes sign once per flipped axis', () => {
      const h = childFrameInParent({ ...child, rotation: 30, flipV: true }, { ...base, flipH: true });
      expect(h.x + h.w / 2).toBeCloseTo(270);
      expect(h.y + h.h / 2).toBeCloseTo(135);
      expect(h).toMatchObject({ rotation: 330, flipH: true, flipV: true });
      const both = childFrameInParent({ ...child, rotation: 30 }, { ...base, flipH: true, flipV: true });
      expect(both.x + both.w / 2).toBeCloseTo(270);
      expect(both.y + both.h / 2).toBeCloseTo(165);
      expect(both).toMatchObject({ rotation: 30, flipH: true, flipV: true });
      const v = childFrameInParent({ ...child, flipH: true }, { ...base, flipV: true });
      expect(v).toMatchObject({ flipH: true, flipV: true });
      expect(v.y + v.h / 2).toBeCloseTo(165);
    });

    it('normalises the rotation into 0-360', () => {
      expect(childFrameInParent({ ...child, rotation: 350 }, { ...base, rotation: 20 }).rotation).toBeCloseTo(10);
      expect(childFrameInParent({ ...child, rotation: 10 }, { ...base, rotation: 0, flipH: true }).rotation).toBeCloseTo(350);
    });

    it('composes nested groups innermost first, and an empty chain is the identity', () => {
      const inner: GroupXfrm = { off: { x: 50, y: 0 }, ext: { w: 100, h: 100 }, chOff: { x: 0, y: 0 }, chExt: { w: 100, h: 100 }, rotation: 0 };
      const outer: GroupXfrm = { off: { x: 0, y: 0 }, ext: { w: 400, h: 200 }, chOff: { x: 0, y: 0 }, chExt: { w: 200, h: 100 }, rotation: 0 };
      const f = composeGroupChain({ x: 0, y: 0, w: 10, h: 10, rotation: 0 }, [inner, outer]);
      expect(f).toEqual({ x: 100, y: 0, w: 20, h: 20, rotation: 0, flipH: false, flipV: false });
      expect(composeGroupChain({ x: 1, y: 2, w: 3, h: 4, rotation: 5 }, [])).toEqual({ x: 1, y: 2, w: 3, h: 4, rotation: 5, flipH: false, flipV: false });
    });
  });
});
