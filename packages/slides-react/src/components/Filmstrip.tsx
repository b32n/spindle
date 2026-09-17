// Filmstrip — the vertical slide navigator. Click to activate, drag to
// reorder (HTML5 DnD → deck.moveSlide), arrow keys to move between slides, and
// Delete to remove the focused slide. An add-slide button sits at the bottom.

import React, { useRef, useState } from 'react';
import { Plus } from 'lucide-react';
import { useDeck, useSlideIds, useActiveSlideId } from '../hooks';
import { ScaledSlide } from './SlideView';

const THUMB_WIDTH = 168;

export function Filmstrip({ readOnly = false }: { readOnly?: boolean } = {}): React.ReactElement {
  const deck = useDeck();
  const slideIds = useSlideIds();
  const activeSlideId = useActiveSlideId();
  const { w } = deck.getSlideSize();
  const scale = THUMB_WIDTH / w;
  const containerRef = useRef<HTMLDivElement>(null);
  const [dragId, setDragId] = useState<string | null>(null);
  const [overId, setOverId] = useState<string | null>(null);

  const onDrop = (targetId: string) => {
    if (readOnly) return;
    if (!dragId || dragId === targetId) return;
    // Drop places the dragged slide immediately after the target.
    deck.moveSlide(dragId, { afterSlideId: targetId });
    setDragId(null);
    setOverId(null);
  };

  const focusThumb = (id: string) => {
    (containerRef.current?.querySelector(`[data-slide-thumb="${id}"]`) as HTMLElement | null)?.focus();
  };

  const deleteSlide = (id: string) => {
    if (readOnly) return;
    if (slideIds.length <= 1) return;
    const i = slideIds.indexOf(id);
    const neighbor = slideIds[i + 1] ?? slideIds[i - 1];
    deck.deleteSlide(id);
    if (neighbor) { deck.setActiveSlide(neighbor); focusThumb(neighbor); }
  };

  const onThumbKeyDown = (e: React.KeyboardEvent, id: string, i: number) => {
    if (e.key === 'Delete' || e.key === 'Backspace') {
      if (readOnly) return;
      e.preventDefault();
      e.stopPropagation();
      deleteSlide(id);
    } else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      // Handle navigation here (and stop it) so it doesn't also bubble to the
      // editor's global handler and move twice.
      e.preventDefault();
      e.stopPropagation();
      const next = slideIds[i + (e.key === 'ArrowDown' ? 1 : -1)];
      if (next) { deck.setActiveSlide(next); focusThumb(next); }
    }
  };

  return (
    <div
      ref={containerRef}
      style={{
        display: 'flex',
        flexDirection: 'column',
        gap: 10,
        padding: 12,
        overflowY: 'auto',
        background: 'rgba(255,255,255,0.7)',
        backdropFilter: 'blur(20px) saturate(180%)',
        WebkitBackdropFilter: 'blur(20px) saturate(180%)',
        border: '1px solid rgba(0,0,0,0.05)',
        borderRadius: 14,
        boxShadow: '0 1px 3px rgba(0,0,0,0.04), 0 8px 24px -8px rgba(0,0,0,0.10)',
        minWidth: THUMB_WIDTH + 48,
        boxSizing: 'border-box',
      }}
    >
      {slideIds.map((id, i) => {
        const active = id === activeSlideId;
        return (
          <div
            key={id}
            draggable={!readOnly}
            onDragStart={() => { if (!readOnly) setDragId(id); }}
            onDragOver={(e) => { if (readOnly) return; e.preventDefault(); setOverId(id); }}
            onDrop={() => onDrop(id)}
            onDragEnd={() => { setDragId(null); setOverId(null); }}
            style={{ display: 'flex', gap: 8, alignItems: 'flex-start', opacity: dragId === id ? 0.4 : 1, borderTop: overId === id && dragId ? '2px solid #2d7ff9' : '2px solid transparent' }}
          >
            <span style={{ fontSize: 12, color: '#8a93a2', width: 16, textAlign: 'right', paddingTop: 4 }}>{i + 1}</span>
            <button
              data-slide-thumb={id}
              onClick={(e) => { deck.setActiveSlide(id); e.currentTarget.focus(); }}
              onKeyDown={(e) => onThumbKeyDown(e, id, i)}
              style={{ padding: 0, border: active ? '2px solid #2d7ff9' : '1px solid #d5d9e0', borderRadius: 4, overflow: 'hidden', cursor: 'pointer', background: '#fff', lineHeight: 0, boxShadow: active ? '0 0 0 2px rgba(45,127,249,0.2)' : 'none' }}
              aria-label={`Go to slide ${i + 1}`}
              aria-current={active}
            >
              <ScaledSlide slideId={id} scale={scale} />
            </button>
          </div>
        );
      })}
      {!readOnly && (
        <button
          onClick={() => { const s = deck.addSlide({ afterSlideId: activeSlideId, layoutId: 'titleContent' }); deck.setActiveSlide(s.id); }}
          style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, marginLeft: 24, height: 40, border: '1px dashed #c4cad3', borderRadius: 6, background: '#fff', color: '#5b6673', cursor: 'pointer', fontSize: 13 }}
        >
          <Plus size={15} /> New slide
        </button>
      )}
    </div>
  );
}
