import { cropLayout, isCropped, MIN_CROP_VISIBLE, normalizeCrop, NO_CROP, setCropEdge } from './crop';
import { createImageElement } from './elements';

describe('normalizeCrop', () => {
  it('returns undefined for nothing, non-objects and an all-zero crop', () => {
    for (const v of [undefined, null, 5, 'x', {}, { l: 0, t: 0, r: 0, b: 0 }, { l: NaN, t: Infinity }]) {
      expect(normalizeCrop(v)).toBeUndefined();
    }
  });
  it('keeps a valid crop and clamps edges to 0-100', () => {
    expect(normalizeCrop({ l: 10, t: 20, r: 30, b: 40 })).toEqual({ l: 10, t: 20, r: 30, b: 40 });
    expect(normalizeCrop({ l: -5, t: 200, r: 0, b: 0 })).toEqual({ l: 0, t: 99, r: 0, b: 0 });
  });
  it('never empties an axis: the second edge gives way', () => {
    expect(normalizeCrop({ l: 60, r: 60, t: 0, b: 0 })).toEqual({ l: 60, r: 39, t: 0, b: 0 });
    expect(normalizeCrop({ l: 100, r: 100, t: 100, b: 100 })).toEqual({ l: 99, r: 0, t: 99, b: 0 });
    const c = normalizeCrop({ l: 50, r: 50, t: 50, b: 50 })!;
    expect(100 - c.l - c.r).toBeGreaterThanOrEqual(MIN_CROP_VISIBLE);
  });
  it('treats non-numeric edges as zero', () => {
    expect(normalizeCrop({ l: '10', t: 5 })).toEqual({ l: 0, t: 5, r: 0, b: 0 });
  });
  it('isCropped follows it', () => {
    expect(isCropped(undefined)).toBe(false);
    expect(isCropped(NO_CROP)).toBe(false);
    expect(isCropped({ l: 1 })).toBe(true);
  });
});

describe('setCropEdge', () => {
  it('sets one edge from nothing', () => {
    expect(setCropEdge(undefined, 'l', 12)).toEqual({ l: 12, t: 0, r: 0, b: 0 });
  });
  it('keeps the other edges and clamps against the opposite one, for every edge', () => {
    const base = { l: 10, t: 10, r: 10, b: 10 };
    expect(setCropEdge(base, 'l', 95)).toEqual({ l: 89, t: 10, r: 10, b: 10 });
    expect(setCropEdge(base, 'r', 95)).toEqual({ l: 10, t: 10, r: 89, b: 10 });
    expect(setCropEdge(base, 't', 95)).toEqual({ l: 10, t: 89, r: 10, b: 10 });
    expect(setCropEdge(base, 'b', 95)).toEqual({ l: 10, t: 10, r: 10, b: 89 });
  });
  it('clamps negatives and NaN to zero and can return to an all-zero crop', () => {
    expect(setCropEdge({ l: 5, t: 0, r: 0, b: 0 }, 'l', -3)).toEqual({ l: 0, t: 0, r: 0, b: 0 });
    expect(setCropEdge({ l: 5, t: 0, r: 0, b: 0 }, 'l', NaN)).toEqual({ l: 0, t: 0, r: 0, b: 0 });
  });
});

describe('cropLayout', () => {
  const crop = { l: 10, t: 20, r: 30, b: 0 }; // visible 60% x 80%
  it('fill (and an undefined fit) stretches the region over the box and positions the whole image behind it', () => {
    for (const fit of ['fill', undefined] as const) {
      const { region, image } = cropLayout(crop, fit, 100, 100, 200, 100);
      expect(region).toEqual({ x: 0, y: 0, w: 100, h: 100 });
      expect(image.w).toBeCloseTo(100 / 0.6);
      expect(image.h).toBeCloseTo(100 / 0.8);
      expect(image.x).toBeCloseTo(-10 / 0.6);
      expect(image.y).toBeCloseTo(-20 / 0.8);
    }
  });
  it('contain letterboxes a wider region vertically', () => {
    // region 600x400 (aspect 1.5) in a 300x300 box: full width, height 200 => 66.67%
    const { region } = cropLayout({ l: 0, t: 0, r: 0, b: 0 }, 'contain', 600, 400, 300, 300);
    expect(region.w).toBe(100);
    expect(region.h).toBeCloseTo(66.667, 2);
    expect(region.y).toBeCloseTo(16.667, 2);
    expect(region.x).toBe(0);
  });
  it('contain pillarboxes a narrower region horizontally', () => {
    const { region } = cropLayout(NO_CROP, 'contain', 200, 400, 300, 300);
    expect(region.h).toBe(100);
    expect(region.w).toBeCloseTo(50);
    expect(region.x).toBeCloseTo(25);
  });
  it('cover overflows the box on the shorter axis (clipped by the box)', () => {
    const wide = cropLayout(NO_CROP, 'cover', 600, 400, 300, 300).region;
    expect(wide.h).toBe(100);
    expect(wide.w).toBeCloseTo(150);
    expect(wide.x).toBeCloseTo(-25);
    const narrow = cropLayout(NO_CROP, 'cover', 200, 400, 300, 300).region;
    expect(narrow.w).toBe(100);
    expect(narrow.h).toBeCloseTo(200);
    expect(narrow.y).toBeCloseTo(-50);
  });
  it('the fit uses the cropped aspect, not the source aspect', () => {
    // 100x100 source cropped to 50% wide: aspect 0.5, so contain in a square box is pillarboxed to 50% wide
    const { region } = cropLayout({ l: 25, t: 0, r: 25, b: 0 }, 'contain', 100, 100, 100, 100);
    expect(region.w).toBeCloseTo(50);
  });
  it('falls back to fill when a size is unknown', () => {
    for (const args of [[0, 100, 10, 10], [100, 100, NaN, 10], [100, -1, 10, 10], [100, 100, 10, 0]] as Array<[number, number, number, number]>) {
      expect(cropLayout(crop, 'cover', ...args).region).toEqual({ x: 0, y: 0, w: 100, h: 100 });
    }
  });
  it('sanitises a bad crop instead of dividing by zero', () => {
    const { image } = cropLayout({ l: 100, t: 0, r: 100, b: 0 }, 'fill', 10, 10, 10, 10);
    expect(Number.isFinite(image.w) && Number.isFinite(image.x)).toBe(true);
  });
  it('property: the image always covers the region exactly when the crop is undone', () => {
    for (let l = 0; l < 90; l += 17) for (let r = 0; r < 90 - l; r += 23) {
      const { image } = cropLayout({ l, t: 3, r, b: 7 }, 'fill', 50, 80, 60, 40);
      const vw = (100 - l - r) / 100;
      // the visible slice of the image, scaled to the region, is [l, 100-r] of the image width
      expect(image.x + (image.w * l) / 100).toBeCloseTo(0, 6);
      expect(image.w * vw).toBeCloseTo(100, 6);
    }
  });
});

describe('createImageElement crop', () => {
  const input = { containerId: 's', index: 'a', x: 0, y: 0, w: 10, h: 10, src: 'data:image/png;base64,AA==', naturalW: 4, naturalH: 4 } as Parameters<typeof createImageElement>[0];
  it('stores a sanitised crop and omits an empty one', () => {
    expect(createImageElement({ ...input, crop: { l: 10, t: 0, r: 0, b: 0 } }).crop).toEqual({ l: 10, t: 0, r: 0, b: 0 });
    expect('crop' in createImageElement({ ...input, crop: NO_CROP })).toBe(false);
    expect('crop' in createImageElement(input)).toBe(false);
  });
});
