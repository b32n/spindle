// Small helpers turning scene styling into concrete SVG/CSS attributes.

import { resolveColor, type Stroke, type ThemeData } from '@b32nio/spindle-slides-core';

export interface StrokeAttrs {
  stroke: string;
  strokeWidth: number;
  strokeDasharray?: string;
}

export function strokeAttrs(stroke: Stroke | undefined, theme: ThemeData): StrokeAttrs | null {
  if (!stroke || stroke.width <= 0) return null;
  const attrs: StrokeAttrs = {
    stroke: resolveColor(stroke.color, theme),
    strokeWidth: stroke.width,
  };
  if (stroke.dash === 'dash') attrs.strokeDasharray = `${stroke.width * 3} ${stroke.width * 2}`;
  else if (stroke.dash === 'dot') attrs.strokeDasharray = `${stroke.width} ${stroke.width * 1.5}`;
  return attrs;
}
