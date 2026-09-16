// Screen ↔ slide coordinate mapping. The interactive surface is the scaled
// slide box, whose bounding rect's top-left is slide (0, 0) and whose CSS
// scale relates screen px to slide px.

import type { Point } from '@b32nio/spindle-slides-core';

export interface StageMetrics {
  /** Bounding client rect of the scaled slide surface. */
  rect: { left: number; top: number };
  /** CSS scale factor applied to the slide. */
  scale: number;
}

export function screenToSlide(clientX: number, clientY: number, m: StageMetrics): Point {
  return {
    x: (clientX - m.rect.left) / m.scale,
    y: (clientY - m.rect.top) / m.scale,
  };
}

/** Distance in slide units corresponding to `px` screen pixels. */
export function screenDistanceToSlide(px: number, scale: number): number {
  return px / scale;
}
