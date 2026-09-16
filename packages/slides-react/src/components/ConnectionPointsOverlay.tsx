// ConnectionPointsOverlay — the connector-drawing affordances, drawn in slide
// coordinates on top of the stage. On hover it shows a shape's 8 connection
// dots (edge midpoints + corners); dragging from a dot draws a connector,
// previewed here live with the snapped target anchor highlighted. Resize/rotate
// handles are untouched — these dots are a separate, hover-only affordance.

import React, { useSyncExternalStore } from 'react';
import { anchorPoints, type AnchorId, type Frame } from '@b32nio/spindle-slides-core';
import { useDeckContext } from '../context/DeckContext';

const ACCENT = '#2d7ff9';

function frameOf(el: { x: number; y: number; w: number; h: number; rotation: number }): Frame {
  return { x: el.x, y: el.y, w: el.w, h: el.h, rotation: el.rotation };
}

function Dot({ x, y, r, elementId, anchor, active }: { x: number; y: number; r: number; elementId: string; anchor: AnchorId; active: boolean }): React.ReactElement {
  return (
    <div
      data-connect-element={elementId}
      data-connect-anchor={anchor}
      style={{
        position: 'absolute',
        left: x - r,
        top: y - r,
        width: r * 2,
        height: r * 2,
        borderRadius: '50%',
        background: active ? ACCENT : '#fff',
        border: `${r * 0.35}px solid ${ACCENT}`,
        boxSizing: 'border-box',
        pointerEvents: 'auto',
        cursor: 'crosshair',
      }}
    />
  );
}

export function ConnectionPointsOverlay({ scale }: { scale: number }): React.ReactElement | null {
  const { deck, connectors } = useDeckContext();
  const state = useSyncExternalStore(connectors.subscribe, connectors.get);
  const { w, h } = deck.getSlideSize();
  const r = 5 / scale;

  const draft = state.draft;
  const edit = state.edit;
  // While drawing → the draft's candidate target; while editing a tip → its snap
  // target; otherwise → the hovered shape.
  const snap = draft?.snap ?? edit?.snap ?? null;
  const dotsForId = draft ? draft.overElementId : edit ? edit.snap?.elementId ?? edit.overElementId : state.hoverId;
  const dotsEl = dotsForId ? deck.getElement(dotsForId) : null;
  const snapAnchor = snap && snap.elementId === dotsForId ? snap.anchor : null;

  if (!draft && !dotsEl) return null;

  return (
    <div style={{ position: 'absolute', left: 0, top: 0, pointerEvents: 'none' }}>
      {/* Live preview of the connector being drawn. */}
      {draft && (
        <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} style={{ position: 'absolute', left: 0, top: 0, overflow: 'visible', pointerEvents: 'none' }}>
          <defs>
            <marker id="connect-preview-head" viewBox="0 0 10 10" refX={9} refY={5} markerWidth={7} markerHeight={7} orient="auto-start-reverse">
              <path d="M0 0 L10 5 L0 10 Z" fill={ACCENT} />
            </marker>
          </defs>
          <line
            x1={draft.from.x} y1={draft.from.y} x2={draft.to.x} y2={draft.to.y}
            stroke={ACCENT} strokeWidth={2 / scale} strokeLinecap="round"
            markerEnd="url(#connect-preview-head)"
          />
        </svg>
      )}
      {/* Connection dots for the hovered / candidate-target element. */}
      {dotsEl && dotsEl.type !== 'line' &&
        anchorPoints(frameOf(dotsEl)).map(({ anchor, point }) => (
          <Dot key={anchor} x={point.x} y={point.y} r={r} elementId={dotsEl.id} anchor={anchor} active={anchor === snapAnchor} />
        ))}
    </div>
  );
}
