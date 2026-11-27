// Slide size presets and the policy for resizing a deck's content. Sizes are px at 96 DPI (12192000 EMU = 1280 px), the
// same unit as every element frame. Changing the size scales each element's frame per axis, so an element keeps its
// relative position and proportion on the slide (the plan's "relative positions by policy"); font sizes are not scaled.

import type { SlideElement } from './scene/types';
import type { LayoutData } from './theme/types';

export type SlideSizePresetId = '16:9' | '16:10' | '4:3' | 'a4' | 'a4-portrait' | 'portrait';

export interface SlideSizePreset {
  id: SlideSizePresetId;
  label: string;
  w: number;
  h: number;
}

export const SLIDE_SIZE_PRESETS: readonly SlideSizePreset[] = [
  { id: '16:9', label: 'Widescreen (16:9)', w: 1280, h: 720 },
  { id: '16:10', label: 'Widescreen (16:10)', w: 1280, h: 800 },
  { id: '4:3', label: 'Standard (4:3)', w: 960, h: 720 },
  { id: 'a4', label: 'A4 landscape', w: 1123, h: 794 },
  { id: 'a4-portrait', label: 'A4 portrait', w: 794, h: 1123 },
  { id: 'portrait', label: 'Portrait (9:16)', w: 720, h: 1280 },
] as const;

/** Smallest and largest side accepted for a custom size (PowerPoint allows 1 to 56 in; this is 100 px to 5376 px). */
export const MIN_SLIDE_SIDE = 100;
export const MAX_SLIDE_SIDE = 5376;

/** The preset a size equals (within a pixel), or `'custom'`. */
export function detectSlideSizePreset(size: { w: number; h: number }): SlideSizePresetId | 'custom' {
  const hit = SLIDE_SIZE_PRESETS.find((p) => Math.abs(p.w - size.w) < 1 && Math.abs(p.h - size.h) < 1);
  return hit ? hit.id : 'custom';
}

/** Round a custom size to whole pixels inside the accepted range; a non-finite side falls back to `fallback`. */
export function clampSlideSize(size: { w: number; h: number }, fallback: { w: number; h: number }): { w: number; h: number } {
  const side = (v: number, f: number) => {
    const n = Number.isFinite(v) ? v : f;
    return Math.min(MAX_SLIDE_SIDE, Math.max(MIN_SLIDE_SIDE, Math.round(n)));
  };
  return { w: side(size.w, fallback.w), h: side(size.h, fallback.h) };
}

/** The patch that scales one element by `sx` horizontally and `sy` vertically (frame, line free ends, table row minimums). */
export function scaleElementPatch(el: SlideElement, sx: number, sy: number): Partial<SlideElement> {
  const patch: Record<string, unknown> = { x: el.x * sx, y: el.y * sy, w: el.w * sx, h: el.h * sy };
  if (el.type === 'line') {
    if (el.startPoint) patch.startPoint = { x: el.startPoint.x * sx, y: el.startPoint.y * sy };
    if (el.endPoint) patch.endPoint = { x: el.endPoint.x * sx, y: el.endPoint.y * sy };
  } else if (el.type === 'table' && el.rowHeights) {
    patch.rowHeights = el.rowHeights.map((r) => r * sy);
  }
  return patch as Partial<SlideElement>;
}

/** A layout with its placeholder frames and static elements scaled like slide content. */
export function scaleLayout(layout: LayoutData, sx: number, sy: number): LayoutData {
  const next: LayoutData = {
    ...layout,
    placeholders: layout.placeholders.map((p) => ({
      ...p,
      frame: { ...p.frame, x: p.frame.x * sx, y: p.frame.y * sy, w: p.frame.w * sx, h: p.frame.h * sy },
    })),
  };
  if (layout.elements) next.elements = layout.elements.map((el) => ({ ...el, ...scaleElementPatch(el, sx, sy) }) as SlideElement);
  return next;
}
