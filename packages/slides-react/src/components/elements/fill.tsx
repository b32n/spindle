// Fill rendering for the element views: CSS for HTML boxes (text boxes, table cells, slide backgrounds) and SVG paint
// servers for shapes. The paint data itself comes from slides-core (`fillToPaint`); this file only maps it to React.

import React from 'react';
import {
  fillToPaint,
  linearGradientVector,
  paintToCss,
  type Fill,
  type FillPaint,
  type ThemeData,
} from '@b32nio/spindle-slides-core';

/** CSS for a fill box. `layer` is set when the fill needs its own translucent layer (a picture with alpha < 1). */
export interface HtmlFill {
  style: React.CSSProperties;
  layer: React.ReactElement | null;
}

const NO_FILL: HtmlFill = { style: {}, layer: null };

/**
 * Style (and, for translucent pictures, a back layer) painting `fill` into an HTML box. A box with a `layer` must be
 * `position: relative; isolation: isolate` so the layer sits behind the box's content.
 */
export function htmlFill(fill: Fill | undefined, theme: ThemeData): HtmlFill {
  const paint = fill ? fillToPaint(fill, theme) : null;
  if (!paint) return NO_FILL;
  const css = paintToCss(paint);
  const { alpha, ...decl } = css;
  if (alpha >= 1) return { style: decl, layer: null };
  const layer = (
    <div
      aria-hidden
      data-fill-layer=""
      style={{ position: 'absolute', inset: 0, zIndex: -1, pointerEvents: 'none', opacity: alpha, ...decl }}
    />
  );
  return { style: { position: 'relative', isolation: 'isolate' }, layer };
}

/**
 * Selector-safe paint-server id for an element. `instance` (React `useId`) keeps ids unique when the same slide is
 * rendered more than once (stage, filmstrip, present mode), so `url(#id)` never binds to another copy.
 */
export function paintId(elementId: string, instance: string): string {
  return `fill-${elementId}-${instance}`.replace(/[^A-Za-z0-9_-]/g, '_');
}

/**
 * The SVG `<defs>` for a shape fill and the `fill` attribute value that uses it. Plain colours need no defs.
 */
export function svgFill(
  fill: Fill,
  theme: ThemeData,
  id: string,
  w: number,
  h: number,
): { defs: React.ReactElement | null; fill: string } {
  const paint = fillToPaint(fill, theme);
  if (!paint) return { defs: null, fill: 'none' };
  if (paint.kind === 'color') return { defs: null, fill: paint.color };
  return { defs: <defs>{paintServer(paint, id, w, h)}</defs>, fill: `url(#${id})` };
}

function paintServer(paint: Exclude<FillPaint, { kind: 'color' }>, id: string, w: number, h: number): React.ReactElement {
  if (paint.kind === 'gradient') {
    const stops = paint.stops.map((s, i) => <stop key={i} offset={s.offset} stopColor={s.color} />);
    if (paint.type === 'linear') {
      const v = linearGradientVector(paint.angle, w, h);
      return (
        <linearGradient id={id} gradientUnits="userSpaceOnUse" x1={v.x1} y1={v.y1} x2={v.x2} y2={v.y2}>
          {stops}
        </linearGradient>
      );
    }
    // Radial and path gradients: an ellipse from the centre reaching the corners (CSS `farthest-corner`). Unlike the
    // linear case (user space, full shape box) this uses bounding-box units, so on non-rect shapes it follows the
    // geometry's bbox.
    return (
      <radialGradient id={id} cx="0.5" cy="0.5" r="0.7071">
        {stops}
      </radialGradient>
    );
  }
  const tw = paint.mode === 'tile' ? paint.tile.w : w;
  const th = paint.mode === 'tile' ? paint.tile.h : h;
  return (
    <pattern id={id} patternUnits="userSpaceOnUse" width={tw} height={th}>
      <image href={paint.src} width={tw} height={th} preserveAspectRatio="none" opacity={paint.alpha} />
    </pattern>
  );
}
