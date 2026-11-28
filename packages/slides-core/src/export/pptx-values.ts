// Value mappings between the PPTX interchange (`@b32nio/spindle-shared/export/pptx`) and the slides model: units, colours,
// fills and outlines. Pure functions, no I/O. The interchange is in px (1 px = 9525 EMU) and pt for type; the editor sets type
// in px on the same 1280 px slide, so a point is 4/3 px: 18 pt text is 24 px, as it measures against the same slide in PowerPoint.

import type { PptxColor, PptxFill, PptxLine } from '@b32nio/spindle-shared/export/pptx';
import type { Color, Fill, Stroke } from '../scene/types';

const PX_PER_PT = 4 / 3;
const round4 = (n: number): number => Math.round(n * 1e4) / 1e4;

/** Points (type sizes, spacing in pt) to px; four decimals keep `11 pt` stable through a round trip (`14.6667` px). */
export const ptToPx = (pt: number): number => round4(pt * PX_PER_PT);
export const pxToPt = (px: number): number => round4(px / PX_PER_PT);

/** A finite number, or `undefined` for anything else (model values come from imported or collaborator JSON). */
export const finite = (v: unknown): number | undefined => (typeof v === 'number' && Number.isFinite(v) ? v : undefined);

export const colorFromPptx = (c: PptxColor): Color => ({ ...c });
export const colorToPptx = (c: Color): PptxColor => ({ ...c });

const BLACK: Color = { kind: 'rgb', hex: '#000000' };
const WHITE: Color = { kind: 'rgb', hex: '#FFFFFF' };

/** A fill; a picture fill is `undefined` (the interchange names the image part without its bytes), which the caller counts. */
export function fillFromPptx(fill: PptxFill | undefined): Fill | undefined {
  switch (fill?.kind) {
    case undefined:
      return undefined;
    case 'none':
      return { kind: 'none' };
    case 'solid':
      return { kind: 'solid', color: colorFromPptx(fill.color) };
    case 'gradient': {
      const stops = fill.stops.map((s) => ({ pos: s.pos, color: colorFromPptx(s.color) }));
      if (fill.shape === 'linear') return { kind: 'gradient', type: 'linear', angle: fill.angle ?? 0, stops };
      return { kind: 'gradient', type: fill.shape === 'radial' ? 'radial' : 'path', stops };
    }
    case 'pattern':
      return { kind: 'pattern', preset: fill.prst, fg: fill.fg ? colorFromPptx(fill.fg) : BLACK, bg: fill.bg ? colorFromPptx(fill.bg) : WHITE };
    case 'picture':
      return undefined;
  }
}

export function fillToPptx(fill: Fill | undefined): PptxFill | undefined {
  switch (fill?.kind) {
    case undefined:
      return undefined;
    case 'none':
      return { kind: 'none' };
    case 'solid':
      return { kind: 'solid', color: colorToPptx(fill.color) };
    case 'gradient': {
      const stops = fill.stops.map((s) => ({ pos: s.pos, color: colorToPptx(s.color) }));
      if (fill.type === 'linear') return { kind: 'gradient', shape: 'linear', angle: fill.angle ?? 90, stops };
      return { kind: 'gradient', shape: fill.type === 'radial' ? 'radial' : 'rect', stops };
    }
    case 'pattern':
      return { kind: 'pattern', prst: fill.preset, fg: colorToPptx(fill.fg), bg: colorToPptx(fill.bg) };
    case 'picture':
      return { kind: 'picture', ...(fill.mode === 'tile' ? { tile: true } : {}), ...(fill.alpha !== undefined ? { alpha: fill.alpha } : {}) };
  }
}

/** The colour an outline's fill stands for: a solid colour, a gradient's first stop or a pattern's foreground. */
function lineColor(fill: PptxFill | undefined): Color | undefined {
  if (fill?.kind === 'solid') return colorFromPptx(fill.color);
  if (fill?.kind === 'gradient' && fill.stops.length > 0) return colorFromPptx(fill.stops[0].color);
  if (fill?.kind === 'pattern' && fill.fg) return colorFromPptx(fill.fg);
  return undefined;
}

const DOTS = new Set(['dot', 'sysDot', 'sysDash']);
const DASHES = new Set(['dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot', 'sysDashDot', 'sysDashDotDot', 'custom']);

/** An outline as a `Stroke`; `undefined` when the line has no colour to draw (no fill, or an explicit no-fill). */
export function strokeFromPptx(line: PptxLine | undefined): Stroke | undefined {
  const color = lineColor(line?.fill);
  if (!line || !color) return undefined;
  const dash = line.dash !== undefined && DOTS.has(line.dash) ? 'dot' : line.dash !== undefined && DASHES.has(line.dash) ? 'dash' : undefined;
  return { color, width: line.width ?? 1, ...(dash ? { dash } : {}) };
}

export function strokeToPptx(stroke: Stroke | undefined): PptxLine | undefined {
  if (!stroke) return undefined;
  const dash = stroke.dash === 'dash' ? 'dash' : stroke.dash === 'dot' ? 'sysDot' : undefined;
  return { fill: { kind: 'solid', color: colorToPptx(stroke.color) }, width: stroke.width, ...(dash ? { dash } : {}) };
}
