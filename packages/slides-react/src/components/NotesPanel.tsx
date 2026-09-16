// NotesPanel — speaker notes for the active slide. v1 edits notes as plain
// text (a textarea) and stores them as a RichTextDoc; rich notes formatting is
// a follow-up. Commits on change.

import React from 'react';
import { richTextFromPlainText, richTextToPlainText } from '@b32nio/spindle-slides-core';
import { useDeck, useActiveSlideId, useSlide } from '../hooks';
import { useDeckContext } from '../context/DeckContext';

export function NotesPanel(): React.ReactElement {
  const deck = useDeck();
  const { editing } = useDeckContext();
  const slideId = useActiveSlideId();
  const slide = useSlide(slideId);
  const value = slide?.notes ? richTextToPlainText(slide.notes) : '';

  // Interacting with notes is "outside" the canvas — clear the element selection
  // and leave any text edit.
  const deselect = () => {
    editing.setEditingId(null);
    deck.setSelection({ slideId, elementIds: [] });
  };

  return (
    <div onPointerDown={deselect} style={{ flex: 'none', background: 'rgba(255,255,255,0.7)', backdropFilter: 'blur(20px) saturate(180%)', WebkitBackdropFilter: 'blur(20px) saturate(180%)', border: '1px solid rgba(0,0,0,0.05)', borderRadius: 14, boxShadow: '0 1px 3px rgba(0,0,0,0.04), 0 8px 24px -8px rgba(0,0,0,0.10)', padding: '10px 16px' }}>
      <div style={{ fontSize: 11, textTransform: 'uppercase', letterSpacing: 0.5, color: '#8a93a2', marginBottom: 4 }}>
        Speaker notes
      </div>
      <textarea
        value={value}
        onChange={(e) => deck.setSlideNotes(slideId, richTextFromPlainText(e.target.value))}
        placeholder="Add notes for the presenter…"
        style={{
          width: '100%',
          minHeight: 60,
          resize: 'vertical',
          border: '1px solid #e2e4e8',
          borderRadius: 6,
          padding: 8,
          fontFamily: 'inherit',
          fontSize: 13,
          color: '#2b3440',
          outline: 'none',
          boxSizing: 'border-box',
        }}
      />
    </div>
  );
}
