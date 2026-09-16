// SlideView renders a slide at native slide coordinates (e.g. 1280×720). The
// caller scales it with CSS — ScaledSlide wraps it in a correctly-sized box so
// it participates in normal layout at any zoom. The same static renderers back
// the editor stage, the filmstrip thumbnails, present mode, and PDF export.

import React from 'react';
import { resolveFill } from '@b32nio/spindle-slides-core';
import { useDeck, useSlide, useSlideElementIds, useTheme } from '../hooks';
import { ElementView } from './elements/ElementView';

export function SlideView({
  slideId,
  interactive = false,
}: {
  slideId: string;
  interactive?: boolean;
}): React.ReactElement {
  const deck = useDeck();
  const slide = useSlide(slideId);
  const theme = useTheme();
  const elementIds = useSlideElementIds(slideId);
  const { w, h } = deck.getSlideSize();

  const background =
    (slide?.background && resolveFill(slide.background, theme)) ??
    resolveFill({ kind: 'solid', color: { kind: 'theme', slot: 'lt1' } }, theme) ??
    '#ffffff';

  return (
    <div
      data-slide-id={slideId}
      style={{
        position: 'relative',
        width: w,
        height: h,
        background,
        overflow: 'hidden',
      }}
    >
      {elementIds.map((id) => (
        <ElementView key={id} elementId={id} interactive={interactive} />
      ))}
    </div>
  );
}

/**
 * SlideView placed inside a box scaled by `scale`, so it occupies
 * (w*scale)×(h*scale) in the surrounding layout.
 */
export function ScaledSlide({ slideId, scale }: { slideId: string; scale: number }): React.ReactElement {
  const deck = useDeck();
  const { w, h } = deck.getSlideSize();
  return (
    <div style={{ width: w * scale, height: h * scale, flex: 'none' }}>
      <div style={{ transform: `scale(${scale})`, transformOrigin: 'top left', width: w, height: h }}>
        <SlideView slideId={slideId} />
      </div>
    </div>
  );
}
