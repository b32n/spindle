// TableResizeOverlay — thin grip strips over a selected table's internal column
// and row boundaries. Dragging a column boundary resizes that column pair;
// dragging a row boundary sets the row above it a minimum height (content can
// still grow it taller). InteractiveSlide owns the gestures (data-col-resize /
// data-row-resize). Row boundaries sit at the *measured* row positions, since
// rows are content-driven. Only shown for an unrotated, single-selected table.

import React, { useEffect, useReducer, useSyncExternalStore } from 'react';
import type { TableElement } from '@b32nio/spindle-slides-core';
import { useDeckContext } from '../context/DeckContext';
import { useSelection } from '../hooks';

export function TableResizeOverlay({ scale }: { scale: number }): React.ReactElement | null {
  const { deck, tableMetrics, transient } = useDeckContext();
  const selection = useSelection();
  const transientState = useSyncExternalStore(transient.subscribe, transient.get);
  const [, force] = useReducer((n) => n + 1, 0);

  useEffect(() => {
    const offs = [
      deck.on('elementChange', force as () => void),
      deck.on('deckChange', force as () => void),
      tableMetrics.subscribe(force as () => void),
    ];
    return () => offs.forEach((o) => o());
  }, [deck, tableMetrics]);

  const id = selection.elementIds.length === 1 ? selection.elementIds[0] : null;
  const el = id ? deck.getElement(id) : null;
  if (!el || el.type !== 'table' || el.rotation !== 0) return null;
  const t = el as TableElement;
  // Track the live frame during a move so the grips don't lag behind the table.
  const live = transientState.liveFrames?.get(t.id);
  const box = { x: live?.x ?? t.x, y: live?.y ?? t.y, w: live?.w ?? t.w, h: live?.h ?? t.h };

  const grip = 8 / scale;
  let ax = 0;
  const colBounds = t.colFractions.slice(0, -1).map((f, i) => ((ax += f), { i, x: ax * box.w }));
  // Row boundaries at the measured row bottoms (rows-1 internal boundaries);
  // grip i sizes the row above it. Fall back to even rows before measurement.
  const measured = tableMetrics.get(t.id);
  const rowEdges = measured && measured.length === t.rows + 1
    ? measured.map((frac) => frac * box.h)
    : Array.from({ length: t.rows + 1 }, (_, r) => (r / t.rows) * box.h);

  return (
    <div style={{ position: 'absolute', left: 0, top: 0, transform: `translate(${box.x}px, ${box.y}px)`, width: box.w, height: box.h, pointerEvents: 'none' }}>
      {colBounds.map((b) => (
        <div key={`c${b.i}`} data-col-resize={b.i} style={{ position: 'absolute', left: b.x - grip / 2, top: 0, width: grip, height: box.h, cursor: 'col-resize', pointerEvents: 'auto' }} />
      ))}
      {Array.from({ length: t.rows - 1 }, (_, i) => (
        <div key={`r${i}`} data-row-resize={i} style={{ position: 'absolute', top: rowEdges[i + 1] - grip / 2, left: 0, height: grip, width: box.w, cursor: 'row-resize', pointerEvents: 'auto' }} />
      ))}
    </div>
  );
}
