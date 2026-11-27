// Fill painting — turn a `Fill` into concrete CSS / SVG paint data at render time. Pure functions only; the React
// renderers map the result onto elements. Picture fills accept data URLs only (the library never fetches).

import type { Fill, GradientFill, PatternFill, PictureFill, Color, ThemeColorSlot } from '../scene/types';
import type { ThemeData } from './types';
import { colorChannels, resolveColor } from './resolve';

/** Data-URL image types a picture fill may carry (anything else is dropped, so a fill cannot reach the network). */
const SAFE_IMAGE_SRC = /^data:image\/(png|jpe?g|gif|webp|bmp|svg\+xml)(;[a-z0-9=.+-]+)*,[A-Za-z0-9+/=%._~!$&'()*;:@-]*$/i;

export function isSafePictureSrc(src: string): boolean {
  return SAFE_IMAGE_SRC.test(src);
}

const finiteOr = (n: number | undefined, fallback: number): number => (n !== undefined && Number.isFinite(n) ? n : fallback);
const clamp01 = (n: number): number => (Number.isFinite(n) ? Math.max(0, Math.min(1, n)) : 0);

/** How much of a pattern preset is the foreground colour (DrawingML presets: pctN, dark/light/wide variants, rest 50%). */
export function patternCoverage(preset: string): number {
  const pct = /^pct(\d{1,2})$/.exec(preset);
  if (pct) return Number(pct[1]) / 100;
  if (/^dk/.test(preset)) return 0.75;
  if (/^lt/.test(preset)) return 0.25;
  if (/^wd/.test(preset)) return 0.6;
  if (/^n(arH|arV)/.test(preset)) return 0.4;
  return 0.5;
}

/** The solid colour a pattern fill is approximated by: foreground and background mixed by the preset's coverage. */
export function patternBlend(fill: PatternFill, theme: ThemeData): Color {
  const c = patternCoverage(fill.preset);
  const [fr, fg, fb, fa] = colorChannels(fill.fg, theme);
  const [br, bg, bb, ba] = colorChannels(fill.bg, theme);
  const mix = (f: number, b: number): number => Math.round(f * c + b * (1 - c));
  const hex = (n: number): string => n.toString(16).padStart(2, '0');
  const alpha = fa * c + ba * (1 - c);
  return {
    kind: 'rgb',
    hex: `#${hex(mix(fr, br))}${hex(mix(fg, bg))}${hex(mix(fb, bb))}`,
    ...(alpha < 1 ? { alpha } : {}),
  };
}

/**
 * Resolve a fill to a single CSS color, or `null` for no fill. Gradients fall back to their first stop and pattern
 * fills to their solid blend; picture fills have no single color (`null`). Renderers use `fillToPaint` instead.
 */
export function resolveFill(fill: Fill, theme: ThemeData): string | null {
  switch (fill.kind) {
    case 'solid':
      return resolveColor(fill.color, theme);
    case 'gradient':
      return fill.stops.length > 0 ? resolveColor(fill.stops[0].color, theme) : null;
    case 'pattern':
      return resolveColor(patternBlend(fill, theme), theme);
    default:
      return null;
  }
}

/** One resolved gradient stop. */
export interface PaintStop {
  offset: number;
  color: string;
}

export type FillPaint =
  | { kind: 'color'; color: string }
  | { kind: 'gradient'; type: GradientFill['type']; angle: number; stops: PaintStop[] }
  | { kind: 'picture'; src: string; mode: PictureFill['mode']; tile: { w: number; h: number }; alpha: number };

/**
 * Resolve a fill to paint data, or `null` for "no fill" (also for an unusable fill: a gradient without stops, a picture
 * whose source is not a safe data URL).
 */
export function fillToPaint(fill: Fill, theme: ThemeData): FillPaint | null {
  switch (fill.kind) {
    case 'none':
      return null;
    case 'solid':
      return { kind: 'color', color: resolveColor(fill.color, theme) };
    case 'pattern':
      return { kind: 'color', color: resolveColor(patternBlend(fill, theme), theme) };
    case 'gradient': {
      if (fill.stops.length === 0) return null;
      const stops = fill.stops.map((s) => ({ offset: clamp01(s.pos), color: resolveColor(s.color, theme) }));
      if (stops.length === 1) return { kind: 'color', color: stops[0].color };
      return { kind: 'gradient', type: fill.type, angle: finiteOr(fill.angle, 90), stops };
    }
    case 'picture':
      if (!isSafePictureSrc(fill.src)) return null;
      return {
        kind: 'picture',
        src: fill.src,
        mode: fill.mode,
        tile: { w: Math.max(1, finiteOr(fill.tile?.w, 64)), h: Math.max(1, finiteOr(fill.tile?.h, 64)) },
        alpha: fill.alpha === undefined ? 1 : clamp01(fill.alpha),
      };
  }
}

/** CSS declarations for a paint; `alpha` < 1 means the caller must draw it on its own layer with that opacity. */
export interface FillCss {
  /** Colours and gradients use the `background` shorthand; pictures use `backgroundImage` plus longhands (never mixed). */
  background?: string;
  backgroundImage?: string;
  backgroundSize?: string;
  backgroundRepeat?: string;
  alpha: number;
}

/** DrawingML angle (clockwise from left-to-right) to a CSS `linear-gradient` angle (clockwise from bottom-to-top). */
export function cssGradientAngle(angle: number): number {
  return (((angle + 90) % 360) + 360) % 360;
}

export function paintToCss(paint: FillPaint): FillCss {
  if (paint.kind === 'color') return { background: paint.color, alpha: 1 };
  if (paint.kind === 'gradient') {
    const stops = paint.stops.map((s) => `${s.color} ${Math.round(s.offset * 10000) / 100}%`).join(', ');
    const background =
      paint.type === 'linear'
        ? `linear-gradient(${cssGradientAngle(paint.angle)}deg, ${stops})`
        : `radial-gradient(ellipse farthest-corner at 50% 50%, ${stops})`;
    return { background, alpha: 1 };
  }
  const url = `url("${paint.src}")`;
  return paint.mode === 'tile'
    ? { backgroundImage: url, backgroundSize: `${paint.tile.w}px ${paint.tile.h}px`, backgroundRepeat: 'repeat', alpha: paint.alpha }
    : { backgroundImage: url, backgroundSize: '100% 100%', backgroundRepeat: 'no-repeat', alpha: paint.alpha };
}

/** CSS for a fill, or `null` for no fill. */
export function fillToCss(fill: Fill, theme: ThemeData): FillCss | null {
  const paint = fillToPaint(fill, theme);
  return paint ? paintToCss(paint) : null;
}

/**
 * The two ends of the SVG gradient line for a linear gradient over a `w` x `h` box (user-space units), matching the CSS
 * gradient-line rule: the line is centred, runs at `angle` and is long enough that the corners get the end colours.
 */
export function linearGradientVector(angle: number, w: number, h: number): { x1: number; y1: number; x2: number; y2: number } {
  const rad = (cssGradientAngle(angle) * Math.PI) / 180;
  const dx = Math.sin(rad);
  const dy = -Math.cos(rad);
  const half = (Math.abs(w * dx) + Math.abs(h * dy)) / 2;
  const cx = w / 2;
  const cy = h / 2;
  const r = (n: number): number => Math.round(n * 1e4) / 1e4 + 0;
  return { x1: r(cx - dx * half), y1: r(cy - dy * half), x2: r(cx + dx * half), y2: r(cy + dy * half) };
}

/** Gradient presets offered by the fill picker (theme-relative, so a theme switch recolours them). */
export const GRADIENT_PRESETS: ReadonlyArray<{ id: string; name: string; fill: GradientFill }> = [
  { id: 'vertical', name: 'Top to bottom', fill: presetGradient('linear', 90, 'accent1', 'accent2') },
  { id: 'horizontal', name: 'Left to right', fill: presetGradient('linear', 0, 'accent1', 'accent2') },
  { id: 'diagonal', name: 'Diagonal', fill: presetGradient('linear', 45, 'accent3', 'accent4') },
  { id: 'radial', name: 'Radial', fill: presetGradient('radial', 0, 'accent1', 'accent5') },
  { id: 'path', name: 'From centre', fill: presetGradient('path', 0, 'accent6', 'accent2') },
];

function presetGradient(type: GradientFill['type'], angle: number, from: ThemeColorSlot, to: ThemeColorSlot): GradientFill {
  return {
    kind: 'gradient',
    type,
    ...(type === 'linear' ? { angle } : {}),
    stops: [
      { pos: 0, color: { kind: 'theme', slot: from } },
      { pos: 1, color: { kind: 'theme', slot: to } },
    ],
  };
}
