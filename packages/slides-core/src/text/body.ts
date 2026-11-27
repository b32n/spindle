// Text body layout — pure helpers for the body-level attributes of PPTX `a:bodyPr` (insets, autofit, vertical text,
// columns) shared by the static renderer, the live editor and the shrink-to-fit search. Every value can come from an
// imported file or a peer, so each helper coerces to a finite, in-range number or a known keyword before it reaches CSS.

import type { Autofit, BodyInsets, BodyStyle, TextVert } from '../scene/types';

/** Inner padding of a text body when nothing says otherwise (px). */
export const DEFAULT_PADDING = 8;
/** The smallest font scale shrink-to-fit will pick (PowerPoint stops at about a quarter). */
export const MIN_FONT_SCALE = 0.25;
/** Granularity of the shrink-to-fit search (PowerPoint steps in 2.5 percent). */
export const FONT_SCALE_STEP = 0.025;
export const MAX_COLUMNS = 16;

const finite = (v: unknown): v is number => typeof v === 'number' && Number.isFinite(v);
const clamp = (v: number, lo: number, hi: number): number => Math.min(hi, Math.max(lo, v));
const inset = (v: unknown, fallback: number): number => (finite(v) ? Math.max(0, v) : fallback);

/** The effective inner padding per side: `insets` (a missing side falls back to `padding`), else uniform `padding`, else 8. */
export function resolveInsets(body: BodyStyle | undefined): BodyInsets {
  const pad = inset(body?.padding, DEFAULT_PADDING);
  const i = body?.insets as Partial<Record<keyof BodyInsets, unknown>> | undefined;
  return { l: inset(i?.l, pad), t: inset(i?.t, pad), r: inset(i?.r, pad), b: inset(i?.b, pad) };
}

/** The CSS `padding` shorthand for the insets. */
export function insetsToCss(insets: BodyInsets): string {
  return `${insets.t}px ${insets.r}px ${insets.b}px ${insets.l}px`;
}

/** The autofit mode, defaulting to `none` for an absent or unknown value. */
export function autofitMode(body: BodyStyle | undefined): Autofit['mode'] {
  const m = body?.autofit?.mode;
  return m === 'shrink' || m === 'resize' ? m : 'none';
}

/** True when the box declares its own overflow behaviour; such a body is not clipped (an unset body keeps clipping). */
export function hasAutofit(body: BodyStyle | undefined): boolean {
  return body?.autofit !== undefined;
}

/** The stored scales of a `shrink` body: `font` multiplies every font size, `line` multiplies the line spacing. 1/1 otherwise. */
export function autofitScales(body: BodyStyle | undefined): { font: number; line: number } {
  if (autofitMode(body) !== 'shrink') return { font: 1, line: 1 };
  const a = body!.autofit!;
  return {
    font: finite(a.fontScale) ? clamp(a.fontScale, 0.1, 1) : 1,
    line: 1 - (finite(a.lnSpcReduction) ? clamp(a.lnSpcReduction, 0, 0.5) : 0),
  };
}

/** True when the stored scales change anything (so the renderer must route sizes through the scale). */
export function isScaled(body: BodyStyle | undefined): boolean {
  const s = autofitScales(body);
  return s.font !== 1 || s.line !== 1;
}

/** The text direction, `horz` for an absent or unknown value. */
export function textVert(body: BodyStyle | undefined): TextVert {
  const v = body?.vert;
  return v === 'vert' || v === 'eaVert' ? v : 'horz';
}

/** CSS for a text direction: a vertical body flows top to bottom, lines stacking from the right (as `a:bodyPr vert`). */
export function writingModeCss(vert: TextVert): { writingMode?: 'vertical-rl'; textOrientation?: 'sideways' | 'mixed' } {
  if (vert === 'vert') return { writingMode: 'vertical-rl', textOrientation: 'sideways' };
  if (vert === 'eaVert') return { writingMode: 'vertical-rl', textOrientation: 'mixed' };
  return {};
}

/** The column setup of a body: a count of 1 to 16 and the gap in px. `count` 1 means a single column. */
export function resolveColumns(body: BodyStyle | undefined): { count: number; gap: number } {
  const c = body?.columns;
  const count = finite(c?.count) ? clamp(Math.round(c!.count), 1, MAX_COLUMNS) : 1;
  return { count, gap: finite(c?.gap) ? Math.max(0, c!.gap!) : 0 };
}

/** The line-spacing reduction that goes with a font scale (PowerPoint pairs 10 or 20 percent with the deeper scales). */
export function lineReductionFor(fontScale: number): number {
  if (fontScale >= 0.95) return 0;
  return fontScale >= 0.8 ? 0.1 : 0.2;
}

export interface ShrinkResult {
  fontScale: number;
  lnSpcReduction: number;
}

/**
 * Shrink-to-fit: the largest font scale (in 2.5 percent steps, never below 25 percent) for which `fits` reports the
 * text inside the box. `fits` receives each candidate and must be monotone (a smaller scale never fits worse), which
 * makes a binary search over the step grid exact. When even the smallest scale overflows, the smallest is returned.
 */
export function fitShrinkScale(fits: (candidate: ShrinkResult) => boolean): ShrinkResult {
  const at = (step: number): ShrinkResult => {
    const fontScale = Math.round((1 - step * FONT_SCALE_STEP) * 1000) / 1000;
    return { fontScale, lnSpcReduction: lineReductionFor(fontScale) };
  };
  const last = Math.round((1 - MIN_FONT_SCALE) / FONT_SCALE_STEP);
  if (fits(at(0))) return at(0);
  let lo = 0; // known not to fit
  let hi = last; // assumed to fit (or the floor)
  while (hi - lo > 1) {
    const mid = (lo + hi) >> 1;
    if (fits(at(mid))) hi = mid;
    else lo = mid;
  }
  return at(hi);
}

/** A body style with `patch` merged in; a patch value of `undefined` removes that attribute. */
export function mergeBodyStyle(body: BodyStyle | undefined, patch: Partial<BodyStyle>): BodyStyle {
  const out: Record<string, unknown> = { ...body, ...patch };
  for (const k of Object.keys(out)) if (out[k] === undefined) delete out[k];
  return out as BodyStyle;
}
