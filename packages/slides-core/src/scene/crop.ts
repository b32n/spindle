// Image crop geometry — pure helpers shared by the editor and the (future) PPTX mapping. A crop is the percent of the
// source trimmed from each edge (`a:srcRect`); the renderer clips the picture to what is left and the image `fit`
// then applies to that cropped region.

import type { ImageCrop } from './types';

/** The least percent of either source axis a crop leaves visible (a crop can never remove the whole image). */
export const MIN_CROP_VISIBLE = 1;

export const NO_CROP: Readonly<ImageCrop> = Object.freeze({ l: 0, t: 0, r: 0, b: 0 });

const clampPct = (v: unknown): number => (typeof v === 'number' && Number.isFinite(v) ? Math.min(100, Math.max(0, v)) : 0);

/** Clamp one axis: each edge to 0-100, then the second edge so at least MIN_CROP_VISIBLE percent stays. */
function axis(a: unknown, b: unknown): [number, number] {
  const first = Math.min(clampPct(a), 100 - MIN_CROP_VISIBLE);
  return [first, Math.min(clampPct(b), 100 - MIN_CROP_VISIBLE - first)];
}

/** Sanitise a crop (any input, e.g. from a file or a peer): finite, in range, never emptying the image. `undefined` when it crops nothing. */
export function normalizeCrop(crop: unknown): ImageCrop | undefined {
  if (!crop || typeof crop !== 'object') return undefined;
  const c = crop as Partial<Record<keyof ImageCrop, unknown>>;
  const [l, r] = axis(c.l, c.r);
  const [t, b] = axis(c.t, c.b);
  return l || r || t || b ? { l, t, r, b } : undefined;
}

export function isCropped(crop: unknown): boolean {
  return normalizeCrop(crop) !== undefined;
}

/** Set one edge of a crop, limited so the opposite edge keeps MIN_CROP_VISIBLE percent visible; returns an all-zero crop when nothing is cropped. */
export function setCropEdge(crop: ImageCrop | undefined, edge: keyof ImageCrop, value: number): ImageCrop {
  const cur = { ...NO_CROP, ...normalizeCrop(crop) };
  const opposite = { l: 'r', r: 'l', t: 'b', b: 't' }[edge] as keyof ImageCrop;
  const max = 100 - MIN_CROP_VISIBLE - cur[opposite];
  cur[edge] = Math.min(clampPct(value), max);
  return cur;
}

/** A rectangle in percent (0-100) of its parent box. */
export interface PctRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface CropLayout {
  /** The cropped region within the element box (the clip), per `fit`. May exceed the box for 'cover'. */
  region: PctRect;
  /** The whole source image, positioned within `region` (it overflows the region by the cropped edges). */
  image: PctRect;
}

/**
 * Where to put the clip and the image for a cropped picture. `fit` applies to the visible region: 'fill' stretches it
 * to the box, 'contain' letterboxes it, 'cover' fills the box keeping the region's aspect (the overflow is clipped by the box).
 * Falls back to 'fill' when a size is unknown (not positive and finite).
 */
export function cropLayout(
  crop: ImageCrop,
  fit: 'fill' | 'contain' | 'cover' | undefined,
  naturalW: number,
  naturalH: number,
  boxW: number,
  boxH: number
): CropLayout {
  const c = { ...NO_CROP, ...normalizeCrop(crop) };
  const vw = (100 - c.l - c.r) / 100;
  const vh = (100 - c.t - c.b) / 100;
  const ok = [naturalW, naturalH, boxW, boxH].every((n) => Number.isFinite(n) && n > 0);
  let rw = 100;
  let rh = 100;
  if (ok && (fit === 'contain' || fit === 'cover')) {
    const region = (naturalW * vw) / (naturalH * vh);
    const wider = region > boxW / boxH;
    // contain: the limiting axis is the narrower one; cover: the opposite
    if (wider === (fit === 'contain')) rh = (100 * boxW) / boxH / region;
    else rw = (100 * region * boxH) / boxW;
  }
  return {
    region: { x: (100 - rw) / 2, y: (100 - rh) / 2, w: rw, h: rh },
    image: { x: -c.l / vw, y: -c.t / vh, w: 100 / vw, h: 100 / vh },
  };
}
