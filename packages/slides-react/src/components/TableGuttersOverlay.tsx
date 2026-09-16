// TableGuttersOverlay — spreadsheet-style header bars just outside a selected
// table's top and left edges, plus a move grip at the top-left corner. Clicking
// a column/row header selects that whole column/row; dragging the corner grip
// moves the table (a clear, discoverable handle — the table body selects cells),
// and a plain click on it selects the whole grid. InteractiveSlide owns the
// pointer handling (data-col-select / data-row-select / data-table-move). The
// bars follow the live frame during a move, and only show for an unrotated,
// single-selected table.

import React, { useEffect, useReducer, useSyncExternalStore } from 'react';
import { Move } from 'lucide-react';
import type { TableElement } from '@b32nio/spindle-slides-core';
import { useDeckContext } from '../context/DeckContext';
import { useSelection } from '../hooks';
import { selectionRect } from '../interactions/table-selection-store';

const BAR = 'rgba(226,232,240,0.95)';
const BAR_ACTIVE = 'rgba(45,127,249,0.28)';
const LINE = 'rgba(45,127,249,0.7)';
const ACCENT = '#2d7ff9';

export function TableGuttersOverlay({ scale }: { scale: number }): React.ReactElement | null {
  const { deck, tableSel, tableMetrics, transient } = useDeckContext();
  const selection = useSelection();
  const sel = useSyncExternalStore(tableSel.subscribe, tableSel.getState);
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
  // Follow the live frame during a move/resize gesture so the bars track the
  // table instead of snapping into place only after the drop.
  const live = transientState.liveFrames?.get(t.id);
  const box = { x: live?.x ?? t.x, y: live?.y ?? t.y, w: live?.w ?? t.w, h: live?.h ?? t.h };

  const rect = sel && sel.tableId === t.id ? selectionRect(sel) : null;
  const size = 15 / scale; // constant screen size for the header bars
  const g = 1 / scale;

  // Column x-edges come from the fractions (fixed layout makes them exact).
  const colX: number[] = [];
  let ax = 0;
  for (const f of t.colFractions) { colX.push(ax); ax += f * box.w; }
  // Row y-edges (rows+1 values) come from the *measured* content-driven rows;
  // fall back to even rows before the first measurement lands.
  const measured = tableMetrics.get(t.id);
  const rowEdges: number[] = measured && measured.length === t.rows + 1
    ? measured.map((frac) => frac * box.h)
    : Array.from({ length: t.rows + 1 }, (_, r) => (r / t.rows) * box.h);

  const barBase: React.CSSProperties = {
    position: 'absolute',
    background: BAR,
    boxSizing: 'border-box',
    pointerEvents: 'auto',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  };

  return (
    <div style={{ position: 'absolute', left: 0, top: 0, transform: `translate(${box.x}px, ${box.y}px)`, width: box.w, height: box.h, pointerEvents: 'none' }}>
      {/* Column headers along the top edge. */}
      {t.colFractions.map((f, c) => {
        const active = !!rect && c >= rect.c0 && c <= rect.c1;
        return (
          <div
            key={`c${c}`}
            data-col-select={c}
            title="Select column"
            style={{ ...barBase, left: colX[c], top: -size - g, width: f * box.w, height: size, background: active ? BAR_ACTIVE : BAR, border: `${g}px solid ${LINE}`, cursor: 'pointer' }}
          />
        );
      })}
      {/* Row headers along the left edge, sized to the real rendered rows. */}
      {Array.from({ length: t.rows }, (_, r) => {
        const active = !!rect && r >= rect.r0 && r <= rect.r1;
        return (
          <div
            key={`r${r}`}
            data-row-select={r}
            title="Select row"
            style={{ ...barBase, top: rowEdges[r], left: -size - g, height: rowEdges[r + 1] - rowEdges[r], width: size, background: active ? BAR_ACTIVE : BAR, border: `${g}px solid ${LINE}`, cursor: 'pointer' }}
          />
        );
      })}
      {/* Top-left corner: drag to move the whole table, click to select all. */}
      <div
        data-table-move="1"
        data-all-select="1"
        title="Drag to move · click to select all cells"
        style={{ ...barBase, left: -size - g, top: -size - g, width: size, height: size, background: ACCENT, border: `${g}px solid ${ACCENT}`, borderRadius: 3 / scale, cursor: 'move' }}
      >
        <Move size={size * 0.7} color="#fff" strokeWidth={2.5} />
      </div>
    </div>
  );
}
