// SlideStage — the central editing canvas. Renders the active slide, centered,
// at a given zoom (fit-to-container or an explicit factor). Ctrl/Cmd + wheel
// (and trackpad pinch, which Chrome reports as a ctrl-wheel) zoom the slide
// toward the cursor instead of the whole page.

import React, { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { useDeck, useActiveSlideId } from '../hooks';
import { useDeckContext } from '../context/DeckContext';
import { SlideView } from './SlideView';
import { InteractiveSlide } from './InteractiveSlide';

const MIN_ZOOM = 0.1;
const MAX_ZOOM = 8;

export interface SlideStageProps {
  /** Explicit zoom factor; when omitted the slide scales to fit the container. */
  zoom?: number;
  /** Enable selection, gestures, and overlays. */
  interactive?: boolean;
  /** Called when the user zooms via the wheel; lets the host leave "fit" mode. */
  onZoomChange?: (zoom: number) => void;
}

export function SlideStage({ zoom, interactive = false, onZoomChange }: SlideStageProps): React.ReactElement {
  const deck = useDeck();
  const { editing } = useDeckContext();
  const activeSlideId = useActiveSlideId();
  const { w, h } = deck.getSlideSize();
  const containerRef = useRef<HTMLDivElement>(null);
  const contentRef = useRef<HTMLDivElement>(null);
  const [fitScale, setFitScale] = useState(1);

  // Clicking the grey area around the slide (the container itself, not the
  // slide surface) clears the selection and leaves any text edit.
  const onGreyPointerDown = (e: React.PointerEvent) => {
    if (!interactive || e.target !== containerRef.current) return;
    editing.setEditingId(null);
    deck.setSelection({ slideId: activeSlideId, elementIds: [] });
  };

  useLayoutEffect(() => {
    if (zoom !== undefined) return;
    const el = containerRef.current;
    if (!el) return;
    const measure = () => {
      const pad = 48;
      const scale = Math.min((el.clientWidth - pad) / w, (el.clientHeight - pad) / h);
      setFitScale(Math.max(0.05, scale));
    };
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, [zoom, w, h]);

  const scale = zoom ?? fitScale;

  // Keep the live scale in a ref so the (non-passive) wheel listener doesn't
  // need re-subscribing on every zoom change.
  const scaleRef = useRef(scale);
  useLayoutEffect(() => {
    scaleRef.current = scale;
  }, [scale]);
  // Where to re-anchor the view after a wheel zoom re-renders (cursor stable).
  const anchor = useRef<{ fx: number; fy: number; cx: number; cy: number } | null>(null);

  useEffect(() => {
    const el = containerRef.current;
    if (!el || !onZoomChange) return;
    const onWheel = (e: WheelEvent) => {
      // Only intercept the zoom gesture (Ctrl/Cmd + wheel, or trackpad pinch);
      // a plain wheel keeps scrolling the canvas.
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      // Clamp the per-event delta so one big mouse notch (~120) doesn't
      // overshoot, then scale exponentially — snappy on both mouse and pinch.
      const dy = Math.sign(e.deltaY) * Math.min(Math.abs(e.deltaY), 50);
      const next = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, scaleRef.current * Math.exp(-dy * 0.01)));
      if (next === scaleRef.current) return;
      const content = contentRef.current;
      if (content) {
        const cr = content.getBoundingClientRect();
        anchor.current = { fx: (e.clientX - cr.left) / cr.width, fy: (e.clientY - cr.top) / cr.height, cx: e.clientX, cy: e.clientY };
      }
      // Advance the ref now so a fast wheel burst (multiple events before React
      // re-renders) accumulates instead of collapsing to a single step.
      scaleRef.current = next;
      onZoomChange(next);
    };
    el.addEventListener('wheel', onWheel, { passive: false });
    return () => el.removeEventListener('wheel', onWheel);
  }, [onZoomChange]);

  // After a wheel zoom re-renders at the new scale, scroll so the point that was
  // under the cursor stays under the cursor.
  useLayoutEffect(() => {
    const a = anchor.current;
    const el = containerRef.current;
    const content = contentRef.current;
    if (!a || !el || !content) return;
    anchor.current = null;
    const cr = content.getBoundingClientRect();
    el.scrollLeft += cr.left + a.fx * cr.width - a.cx;
    el.scrollTop += cr.top + a.fy * cr.height - a.cy;
  }, [scale]);

  return (
    <div
      ref={containerRef}
      data-slide-stage=""
      onPointerDown={onGreyPointerDown}
      style={{
        flex: '1 1 auto',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        overflow: 'auto',
        background: 'transparent',
      }}
    >
      <div ref={contentRef} style={{ flex: 'none', margin: 'auto' }}>
        {activeSlideId && interactive ? (
          <InteractiveSlide key={activeSlideId} slideId={activeSlideId} scale={scale} />
        ) : (
          <div style={{ width: w * scale, height: h * scale, flex: 'none' }}>
            <div style={{ transform: `scale(${scale})`, transformOrigin: 'top left', width: w, height: h, boxShadow: '0 4px 24px rgba(0,0,0,0.16)' }}>
              {activeSlideId ? <SlideView slideId={activeSlideId} /> : null}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
