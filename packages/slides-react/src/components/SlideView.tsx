// SlideView renders a slide at native slide coordinates (e.g. 1280×720). The
// caller scales it with CSS — ScaledSlide wraps it in a correctly-sized box so
// it participates in normal layout at any zoom. The same static renderers back
// the editor stage, the filmstrip thumbnails, present mode, and PDF export.

import React from 'react';
import { resolveFill } from '@b32nio/spindle-slides-core';
import { htmlFill } from './elements/fill';
import { useDeck, useSlide, useSlideElementIds, useSlideIds, useTheme } from '../hooks';
import { ElementView, StaticElement } from './elements/ElementView';
import { FieldContextValue } from './elements/StaticRichText';

/**
 * Gradients and colours own the whole `background`; pictures set only longhands, so the theme colour goes under them
 * (a transparent PNG still shows it). With nothing paintable the theme colour alone is the background.
 */
function slideBackgroundStyle(style: React.CSSProperties, fallback: string): React.CSSProperties {
  if (style.background) return style;
  if (style.backgroundImage) return { backgroundColor: fallback, ...style };
  return { background: fallback };
}

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
  // Slide-number fields show the slide's place in the deck; the date is read once per render of this slide.
  const slideNumber = useSlideIds().indexOf(slideId) + 1;
  const fields = React.useMemo(() => ({ slideNumber: slideNumber > 0 ? slideNumber : null, now: new Date() }), [slideNumber]);

  // Static content of the slide's layout (logos, rules) sits under the slide's own elements.
  const layoutElements = slide?.layoutRef ? deck.getLayout(slide.layoutRef)?.elements : undefined;

  const bg = htmlFill(slide?.background, theme);
  // No usable background (absent, none, unsafe picture): the theme's light colour, as before.
  const fallback = resolveFill({ kind: 'solid', color: { kind: 'theme', slot: 'lt1' } }, theme) ?? '#ffffff';

  return (
    <div
      data-slide-id={slideId}
      style={{
        position: 'relative',
        width: w,
        height: h,
        ...slideBackgroundStyle(bg.style, fallback),
        ...(bg.layer ? { isolation: 'isolate' } : {}),
        overflow: 'hidden',
      }}
    >
      {bg.layer}
      <FieldContextValue.Provider value={fields}>
        {layoutElements?.map((el) => (
          <StaticElement key={el.id} el={el} />
        ))}
        {elementIds.map((id) => (
          <ElementView key={id} elementId={id} interactive={interactive} />
        ))}
      </FieldContextValue.Provider>
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
