import {
  DEFAULT_PADDING,
  MAX_COLUMNS,
  MIN_FONT_SCALE,
  autofitMode,
  autofitScales,
  fitShrinkScale,
  hasAutofit,
  insetsToCss,
  isScaled,
  lineReductionFor,
  mergeBodyStyle,
  resolveColumns,
  resolveInsets,
  textVert,
  writingModeCss,
} from './body';
import type { BodyStyle } from '../scene/types';

describe('resolveInsets', () => {
  it('defaults to 8px on every side', () => {
    expect(resolveInsets(undefined)).toEqual({ l: 8, t: 8, r: 8, b: 8 });
    expect(DEFAULT_PADDING).toBe(8);
  });
  it('uses the uniform padding when there are no insets', () => {
    expect(resolveInsets({ padding: 3 })).toEqual({ l: 3, t: 3, r: 3, b: 3 });
    expect(resolveInsets({ padding: 0 })).toEqual({ l: 0, t: 0, r: 0, b: 0 });
  });
  it('per-side insets win and a missing side falls back to the padding', () => {
    expect(resolveInsets({ padding: 5, insets: { l: 1, t: 2, r: 3, b: 4 } })).toEqual({ l: 1, t: 2, r: 3, b: 4 });
    expect(resolveInsets({ padding: 5, insets: { l: 1 } as never })).toEqual({ l: 1, t: 5, r: 5, b: 5 });
  });
  it('coerces junk and negatives', () => {
    expect(resolveInsets({ padding: NaN as number, insets: { l: -4, t: 'x', r: Infinity, b: 6 } as never })).toEqual({ l: 0, t: 8, r: 8, b: 6 });
    expect(resolveInsets({ padding: 'url(x)' as never })).toEqual({ l: 8, t: 8, r: 8, b: 8 });
  });
  it('builds the CSS shorthand in top right bottom left order', () => {
    expect(insetsToCss({ l: 1, t: 2, r: 3, b: 4 })).toBe('2px 3px 4px 1px');
  });
});

describe('autofit', () => {
  it('mode defaults to none and ignores unknown values', () => {
    expect(autofitMode(undefined)).toBe('none');
    expect(autofitMode({})).toBe('none');
    expect(autofitMode({ autofit: { mode: 'weird' as never } })).toBe('none');
    expect(autofitMode({ autofit: { mode: 'shrink' } })).toBe('shrink');
    expect(autofitMode({ autofit: { mode: 'resize' } })).toBe('resize');
    expect(autofitMode({ autofit: { mode: 'none' } })).toBe('none');
  });
  it('hasAutofit is true for any declared autofit, including none', () => {
    expect(hasAutofit(undefined)).toBe(false);
    expect(hasAutofit({})).toBe(false);
    expect(hasAutofit({ autofit: { mode: 'none' } })).toBe(true);
  });
  it('scales apply only to shrink and are clamped', () => {
    expect(autofitScales(undefined)).toEqual({ font: 1, line: 1 });
    expect(autofitScales({ autofit: { mode: 'resize', fontScale: 0.5 } })).toEqual({ font: 1, line: 1 });
    expect(autofitScales({ autofit: { mode: 'shrink' } })).toEqual({ font: 1, line: 1 });
    expect(autofitScales({ autofit: { mode: 'shrink', fontScale: 0.625, lnSpcReduction: 0.2 } })).toEqual({ font: 0.625, line: 0.8 });
    expect(autofitScales({ autofit: { mode: 'shrink', fontScale: 7, lnSpcReduction: 9 } })).toEqual({ font: 1, line: 0.5 });
    expect(autofitScales({ autofit: { mode: 'shrink', fontScale: -1, lnSpcReduction: -1 } })).toEqual({ font: 0.1, line: 1 });
    expect(autofitScales({ autofit: { mode: 'shrink', fontScale: NaN, lnSpcReduction: 'x' as never } })).toEqual({ font: 1, line: 1 });
  });
  it('isScaled reports whether the stored scales change anything', () => {
    expect(isScaled(undefined)).toBe(false);
    expect(isScaled({ autofit: { mode: 'shrink', fontScale: 1 } })).toBe(false);
    expect(isScaled({ autofit: { mode: 'shrink', fontScale: 0.9 } })).toBe(true);
    expect(isScaled({ autofit: { mode: 'shrink', lnSpcReduction: 0.1 } })).toBe(true);
  });
});

describe('vertical text and columns', () => {
  it('textVert defaults to horz', () => {
    expect(textVert(undefined)).toBe('horz');
    expect(textVert({ vert: 'bad' as never })).toBe('horz');
    expect(textVert({ vert: 'vert' })).toBe('vert');
    expect(textVert({ vert: 'eaVert' })).toBe('eaVert');
  });
  it('maps directions to writing modes', () => {
    expect(writingModeCss('horz')).toEqual({});
    expect(writingModeCss('vert')).toEqual({ writingMode: 'vertical-rl', textOrientation: 'sideways' });
    expect(writingModeCss('eaVert')).toEqual({ writingMode: 'vertical-rl', textOrientation: 'mixed' });
  });
  it('resolveColumns clamps the count and the gap', () => {
    expect(resolveColumns(undefined)).toEqual({ count: 1, gap: 0 });
    expect(resolveColumns({ columns: { count: 3, gap: 12 } })).toEqual({ count: 3, gap: 12 });
    expect(resolveColumns({ columns: { count: 2.6 } })).toEqual({ count: 3, gap: 0 });
    expect(resolveColumns({ columns: { count: 0, gap: -5 } })).toEqual({ count: 1, gap: 0 });
    expect(resolveColumns({ columns: { count: 99 } })).toEqual({ count: MAX_COLUMNS, gap: 0 });
    expect(resolveColumns({ columns: { count: NaN, gap: NaN } })).toEqual({ count: 1, gap: 0 });
  });
});

describe('fitShrinkScale', () => {
  it('pairs deeper scales with a line-spacing reduction', () => {
    expect(lineReductionFor(1)).toBe(0);
    expect(lineReductionFor(0.95)).toBe(0);
    expect(lineReductionFor(0.9)).toBe(0.1);
    expect(lineReductionFor(0.8)).toBe(0.1);
    expect(lineReductionFor(0.79)).toBe(0.2);
  });
  it('keeps scale 1 when the text already fits', () => {
    const calls: number[] = [];
    expect(fitShrinkScale((c) => (calls.push(c.fontScale), true))).toEqual({ fontScale: 1, lnSpcReduction: 0 });
    expect(calls).toEqual([1]);
  });
  it('finds the largest fitting scale on the 2.5% grid (oracle table)', () => {
    // Content height is proportional to the font scale, so the box fits when scale <= limit.
    const cases: Array<[number, number]> = [
      [0.99, 0.975],
      [0.9, 0.9],
      [0.8, 0.8],
      [0.625, 0.625],
      [0.61, 0.6],
      [0.3, 0.3],
      [0.26, 0.25],
    ];
    for (const [limit, expected] of cases) {
      const r = fitShrinkScale((c) => c.fontScale <= limit + 1e-9);
      expect(r.fontScale).toBeCloseTo(expected, 6);
      expect(r.lnSpcReduction).toBe(lineReductionFor(r.fontScale));
    }
  });
  it('returns the floor when nothing fits, and is monotone-search exact', () => {
    expect(fitShrinkScale(() => false)).toEqual({ fontScale: MIN_FONT_SCALE, lnSpcReduction: 0.2 });
    // Property: for every limit on the grid the search returns exactly the best fitting step.
    for (let step = 0; step <= 30; step++) {
      const limit = 1 - step * 0.025;
      const r = fitShrinkScale((c) => c.fontScale <= limit + 1e-9);
      expect(r.fontScale).toBeCloseTo(Math.max(limit, MIN_FONT_SCALE), 6);
    }
  });
  it('uses O(log n) probes', () => {
    let n = 0;
    fitShrinkScale((c) => (n++, c.fontScale <= 0.5));
    expect(n).toBeLessThanOrEqual(8);
  });
});

it('BodyStyle type accepts the new attributes', () => {
  const b: BodyStyle = { insets: { l: 1, t: 1, r: 1, b: 1 }, anchorCtr: true, autofit: { mode: 'shrink', fontScale: 0.9 }, vert: 'eaVert', columns: { count: 2, gap: 4 } };
  expect(b.autofit?.mode).toBe('shrink');
});

describe('mergeBodyStyle', () => {
  it('merges, replaces and removes attributes without mutating', () => {
    const base: BodyStyle = { padding: 4, vert: 'vert', autofit: { mode: 'shrink', fontScale: 0.5 } };
    const merged = mergeBodyStyle(base, { vert: undefined, autofit: { mode: 'resize' }, anchorCtr: true });
    expect(merged).toEqual({ padding: 4, autofit: { mode: 'resize' }, anchorCtr: true });
    expect(base.vert).toBe('vert');
    expect(mergeBodyStyle(undefined, { padding: 1 })).toEqual({ padding: 1 });
  });
});
