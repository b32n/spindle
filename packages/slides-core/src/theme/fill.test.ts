import type { Fill, ThemeData } from '../index';
import { getBuiltinTheme } from './builtin';
import { colorChannels, resolveColor } from './resolve';
import {
  GRADIENT_PRESETS,
  cssGradientAngle,
  fillToCss,
  fillToPaint,
  isSafePictureSrc,
  linearGradientVector,
  patternBlend,
  patternCoverage,
  paintToCss,
  resolveFill,
} from './fill';

const theme: ThemeData = {
  ...getBuiltinTheme(),
  colors: { ...getBuiltinTheme().colors, accent1: '#102030', accent2: '#ffffff' },
};
const PNG = 'data:image/png;base64,iVBORw0KGgo=';
const red = { kind: 'rgb', hex: '#ff0000' } as const;
const blue = { kind: 'rgb', hex: '#0000ff' } as const;

describe('colorChannels / resolveColor', () => {
  it('returns channels and alpha, clamping alpha', () => {
    expect(colorChannels(red, theme)).toEqual([255, 0, 0, 1]);
    expect(colorChannels({ ...red, alpha: 2 }, theme)[3]).toBe(1);
    expect(colorChannels({ ...red, alpha: -1 }, theme)[3]).toBe(0);
    expect(colorChannels({ kind: 'theme', slot: 'accent1' }, theme)).toEqual([16, 32, 48, 1]);
    expect(colorChannels({ kind: 'theme', slot: 'accent1', lumMod: 2, lumOff: 0.1 }, theme)).toEqual([58, 90, 122, 1]);
  });
  it('keeps the CSS string forms', () => {
    expect(resolveColor(red, theme)).toBe('rgb(255, 0, 0)');
    expect(resolveColor({ ...red, alpha: 0.5 }, theme)).toBe('rgba(255, 0, 0, 0.5)');
    expect(resolveColor({ ...red, alpha: 1 }, theme)).toBe('rgb(255, 0, 0)');
  });
});

describe('isSafePictureSrc', () => {
  it.each([
    [PNG, true],
    ['data:image/jpeg;base64,/9j/4AAQ', true],
    ['data:image/svg+xml;charset=utf-8,%3Csvg%3E%3C/svg%3E', true],
    ['https://example.com/a.png', false],
    ['http://example.com/a.png', false],
    ['//example.com/a.png', false],
    ['javascript:alert(1)', false],
    ['data:text/html;base64,PGI+', false],
    ['data:image/png;base64,AAAA")};x{', false],
    ['data:image/png;base64,AA AA', false],
    ['', false],
  ])('%s -> %s', (src, ok) => expect(isSafePictureSrc(src)).toBe(ok));
});

describe('pattern approximation', () => {
  it.each([
    ['pct5', 0.05],
    ['pct50', 0.5],
    ['pct90', 0.9],
    ['dkHorz', 0.75],
    ['ltVert', 0.25],
    ['wdUpDiag', 0.6],
    ['narHorz', 0.4],
    ['narVert', 0.4],
    ['smCheck', 0.5],
    ['unknownPreset', 0.5],
  ])('coverage of %s is %s', (preset, c) => expect(patternCoverage(preset)).toBe(c));

  it('mixes foreground and background by coverage', () => {
    const half = patternBlend({ kind: 'pattern', preset: 'pct50', fg: red, bg: blue }, theme);
    expect(half).toEqual({ kind: 'rgb', hex: '#800080' });
    const dark = patternBlend({ kind: 'pattern', preset: 'dkHorz', fg: red, bg: blue }, theme);
    expect(dark).toEqual({ kind: 'rgb', hex: '#bf0040' });
  });
  it('mixes alpha and omits it when opaque', () => {
    const c = patternBlend({ kind: 'pattern', preset: 'pct50', fg: { ...red, alpha: 0.5 }, bg: { ...blue, alpha: 0 } }, theme);
    expect(c).toEqual({ kind: 'rgb', hex: '#800080', alpha: 0.25 });
  });
  it('resolves theme colours too', () => {
    const c = patternBlend({ kind: 'pattern', preset: 'pct0', fg: red, bg: { kind: 'theme', slot: 'accent1' } }, theme);
    expect(c).toEqual({ kind: 'rgb', hex: '#102030' });
  });
});

describe('fillToPaint', () => {
  it('none -> null', () => expect(fillToPaint({ kind: 'none' }, theme)).toBeNull());
  it('solid and pattern -> a colour', () => {
    expect(fillToPaint({ kind: 'solid', color: red }, theme)).toEqual({ kind: 'color', color: 'rgb(255, 0, 0)' });
    expect(fillToPaint({ kind: 'pattern', preset: 'pct50', fg: red, bg: blue }, theme)).toEqual({
      kind: 'color',
      color: 'rgb(128, 0, 128)',
    });
  });
  it('gradient resolves stops, clamps offsets and defaults the angle', () => {
    const p = fillToPaint(
      { kind: 'gradient', type: 'linear', stops: [{ pos: -1, color: red }, { pos: 2, color: { ...blue, alpha: 0.5 } }, { pos: NaN, color: red }] },
      theme,
    );
    expect(p).toEqual({
      kind: 'gradient',
      type: 'linear',
      angle: 90,
      stops: [
        { offset: 0, color: 'rgb(255, 0, 0)' },
        { offset: 1, color: 'rgba(0, 0, 255, 0.5)' },
        { offset: 0, color: 'rgb(255, 0, 0)' },
      ],
    });
  });
  it('a gradient with one stop is a colour, with none is no fill', () => {
    expect(fillToPaint({ kind: 'gradient', type: 'radial', stops: [{ pos: 0.3, color: red }] }, theme)).toEqual({
      kind: 'color',
      color: 'rgb(255, 0, 0)',
    });
    expect(fillToPaint({ kind: 'gradient', type: 'radial', stops: [] }, theme)).toBeNull();
  });
  it('non-finite angle and tile sizes fall back to the defaults', () => {
    const stops = [{ pos: 0, color: red }, { pos: 1, color: blue }];
    for (const angle of [NaN, Infinity, -Infinity]) {
      expect((fillToPaint({ kind: 'gradient', type: 'linear', angle, stops }, theme) as { angle: number }).angle).toBe(90);
    }
    expect(fillToPaint({ kind: 'gradient', type: 'linear', angle: 30, stops }, theme)).toMatchObject({ angle: 30 });
    for (const bad of [NaN, Infinity]) {
      expect(fillToPaint({ kind: 'picture', src: PNG, mode: 'tile', tile: { w: bad, h: bad } }, theme)).toMatchObject({
        tile: { w: 64, h: 64 },
      });
    }
    expect(fillToPaint({ kind: 'picture', src: PNG, mode: 'tile', tile: { w: 12, h: 5 } }, theme)).toMatchObject({ tile: { w: 12, h: 5 } });
  });
  it('picture: safe data URL only, defaults and clamps', () => {
    expect(fillToPaint({ kind: 'picture', src: 'https://x/y.png', mode: 'stretch' }, theme)).toBeNull();
    expect(fillToPaint({ kind: 'picture', src: PNG, mode: 'stretch' }, theme)).toEqual({
      kind: 'picture', src: PNG, mode: 'stretch', tile: { w: 64, h: 64 }, alpha: 1,
    });
    expect(fillToPaint({ kind: 'picture', src: PNG, mode: 'tile', tile: { w: 0, h: 20 }, alpha: 3 }, theme)).toEqual({
      kind: 'picture', src: PNG, mode: 'tile', tile: { w: 1, h: 20 }, alpha: 1,
    });
    expect((fillToPaint({ kind: 'picture', src: PNG, mode: 'stretch', alpha: 0.4 }, theme) as { alpha: number }).alpha).toBe(0.4);
  });
});

describe('CSS output', () => {
  it('maps DrawingML angles to CSS angles', () => {
    expect([0, 90, 180, 270, 360, -90, 450].map(cssGradientAngle)).toEqual([90, 180, 270, 0, 90, 0, 180]);
  });
  it('linear and radial gradients', () => {
    const lin = fillToCss({ kind: 'gradient', type: 'linear', angle: 0, stops: [{ pos: 0, color: red }, { pos: 0.333333, color: blue }] }, theme);
    expect(lin).toEqual({ background: 'linear-gradient(90deg, rgb(255, 0, 0) 0%, rgb(0, 0, 255) 33.33%)', alpha: 1 });
    for (const type of ['radial', 'path'] as const) {
      const css = fillToCss({ kind: 'gradient', type, stops: [{ pos: 0, color: red }, { pos: 1, color: blue }] }, theme)!;
      expect(css.background).toBe('radial-gradient(ellipse farthest-corner at 50% 50%, rgb(255, 0, 0) 0%, rgb(0, 0, 255) 100%)');
    }
  });
  it('colours, pictures and none', () => {
    expect(fillToCss({ kind: 'solid', color: red }, theme)).toEqual({ background: 'rgb(255, 0, 0)', alpha: 1 });
    expect(fillToCss({ kind: 'none' }, theme)).toBeNull();
    expect(fillToCss({ kind: 'picture', src: PNG, mode: 'stretch', alpha: 0.5 }, theme)).toEqual({
      backgroundImage: `url("${PNG}")`, backgroundSize: '100% 100%', backgroundRepeat: 'no-repeat', alpha: 0.5,
    });
    expect(paintToCss({ kind: 'picture', src: PNG, mode: 'tile', tile: { w: 8, h: 9 }, alpha: 1 })).toEqual({
      backgroundImage: `url("${PNG}")`, backgroundSize: '8px 9px', backgroundRepeat: 'repeat', alpha: 1,
    });
  });
});

describe('linearGradientVector', () => {
  const near = (a: number, b: number) => expect(a).toBeCloseTo(b, 6);
  it('left to right spans the width at mid height', () => {
    const v = linearGradientVector(0, 200, 100);
    [v.x1, v.y1, v.x2, v.y2].forEach((n, i) => near(n, [0, 50, 200, 50][i]));
  });
  it('top to bottom spans the height', () => {
    const v = linearGradientVector(90, 200, 100);
    [v.x1, v.y1, v.x2, v.y2].forEach((n, i) => near(n, [100, 0, 100, 100][i]));
  });
  it('diagonal reaches the corners (centre symmetric, longer than either side)', () => {
    const v = linearGradientVector(45, 200, 100);
    near((v.x1 + v.x2) / 2, 100);
    near((v.y1 + v.y2) / 2, 50);
    expect(Math.hypot(v.x2 - v.x1, v.y2 - v.y1)).toBeCloseTo(Math.abs(200 * Math.SQRT1_2) + Math.abs(100 * Math.SQRT1_2), 6);
    expect(v.x2).toBeGreaterThan(v.x1);
    expect(v.y2).toBeGreaterThan(v.y1);
  });
});

describe('colorChannels fallbacks', () => {
  it('a theme slot missing from the theme is black', () => {
    const sparse = { ...theme, colors: {} } as unknown as ThemeData;
    expect(colorChannels({ kind: 'theme', slot: 'accent1' }, sparse)).toEqual([0, 0, 0, 1]);
    expect(colorChannels({ kind: 'rgb', hex: '#abc' }, sparse)).toEqual([170, 187, 204, 1]);
  });
});

describe('resolveFill', () => {
  it('solid, pattern blend, first gradient stop; null otherwise', () => {
    expect(resolveFill({ kind: 'solid', color: red }, theme)).toBe('rgb(255, 0, 0)');
    expect(resolveFill({ kind: 'pattern', preset: 'pct50', fg: red, bg: blue }, theme)).toBe('rgb(128, 0, 128)');
    expect(resolveFill({ kind: 'gradient', type: 'linear', stops: [{ pos: 0, color: blue }, { pos: 1, color: red }] }, theme)).toBe('rgb(0, 0, 255)');
    expect(resolveFill({ kind: 'gradient', type: 'linear', stops: [] }, theme)).toBeNull();
    expect(resolveFill({ kind: 'picture', src: PNG, mode: 'stretch' }, theme)).toBeNull();
    expect(resolveFill({ kind: 'none' }, theme)).toBeNull();
  });
});

describe('GRADIENT_PRESETS', () => {
  it('have unique ids and paint with the built-in theme', () => {
    expect(new Set(GRADIENT_PRESETS.map((p) => p.id)).size).toBe(GRADIENT_PRESETS.length);
    for (const p of GRADIENT_PRESETS) {
      const fill: Fill = p.fill;
      expect(fillToPaint(fill, theme)?.kind).toBe('gradient');
      expect(p.fill.type === 'linear').toBe(p.fill.angle !== undefined);
    }
  });
});
