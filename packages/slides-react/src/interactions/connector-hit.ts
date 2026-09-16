// Hit-testing helpers for the connector-drawing UI: find the shape under the
// cursor (to show its connection dots) and the nearest anchor to snap a
// dragged endpoint to. Connectors can start from / attach to any non-line
// element; lines themselves are never connection targets.

import { anchorPoints, pointInFrame } from '@b32nio/spindle-slides-core';
import type { AnchorId, Frame, Point, DeckImpl } from '@b32nio/spindle-slides-core';

function frameOf(el: { x: number; y: number; w: number; h: number; rotation: number }): Frame {
  return { x: el.x, y: el.y, w: el.w, h: el.h, rotation: el.rotation };
}

/** Topmost connectable element whose frame (inflated by `margin`) contains p. */
export function elementUnderPoint(deck: DeckImpl, slideId: string, p: Point, margin = 0): string | null {
  const ids = deck.getElementIdsForSlide(slideId); // back-to-front
  for (let i = ids.length - 1; i >= 0; i--) {
    const el = deck.getElement(ids[i]);
    if (!el || el.type === 'line') continue;
    const f = frameOf(el);
    const inflated: Frame = { x: f.x - margin, y: f.y - margin, w: f.w + 2 * margin, h: f.h + 2 * margin, rotation: f.rotation };
    if (pointInFrame(p, inflated)) return el.id;
  }
  return null;
}

/** Nearest anchor (across all connectable elements bar `excludeId`) within maxDist. */
export function nearestAnchor(
  deck: DeckImpl,
  slideId: string,
  p: Point,
  excludeId: string | null,
  maxDist: number
): { elementId: string; anchor: AnchorId; point: Point } | null {
  let best: { elementId: string; anchor: AnchorId; point: Point } | null = null;
  let bestD = maxDist;
  for (const id of deck.getElementIdsForSlide(slideId)) {
    if (id === excludeId) continue;
    const el = deck.getElement(id);
    if (!el || el.type === 'line') continue;
    for (const { anchor, point } of anchorPoints(frameOf(el))) {
      const d = Math.hypot(point.x - p.x, point.y - p.y);
      if (d <= bestD) {
        bestD = d;
        best = { elementId: id, anchor, point };
      }
    }
  }
  return best;
}
