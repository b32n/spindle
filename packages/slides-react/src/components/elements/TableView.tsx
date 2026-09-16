// TableView — renders a table element as a fixed-layout HTML <table>. Column
// widths come from the element's fractions; row heights are content-driven (a
// row grows to fit its text), so the element's height auto-syncs to the table
// and the overlays read the measured row positions (see TableMetricsStore).
// Each cell shows its rich text via StaticRichText, or a live editor for the
// cell being edited. Grid lines use the element's border stroke.

import React, { useEffect, useLayoutEffect, useRef, useSyncExternalStore } from 'react';
import { resolveColor, resolveFill, type Color, type TableElement, type ThemeData } from '@b32nio/spindle-slides-core';
import { useDeckContext } from '../../context/DeckContext';
import { inSelection } from '../../interactions/table-selection-store';
import { StaticRichText } from './StaticRichText';
import { RichTextEditor } from '../RichTextEditor';

// Translucent tint painted over a selected cell (sits above its own fill, below
// the text) via a huge inset shadow the cell's overflow:hidden clips flat.
const SELECT_TINT = 'inset 0 0 0 9999px rgba(45,127,249,0.20)';

// Cell text metrics. StaticRichText inherits font-size from the <td>, so we set
// it explicitly to the SAME values the live editor uses — otherwise text would
// resize/shift the moment you double-click into a cell.
const CELL_FONT = 16;
const CELL_PAD = 6;

export function TableView({ el, theme, interactive = false }: { el: TableElement; theme: ThemeData; interactive?: boolean }): React.ReactElement {
  const { deck, editing, tableSel, tableMetrics, transient } = useDeckContext();
  const editState = useSyncExternalStore(editing.subscribe, editing.getState);
  const sel = useSyncExternalStore(tableSel.subscribe, tableSel.getState);
  const editingCell = interactive && editState.id === el.id ? editState.cell : null;
  const cellSel = interactive && sel && sel.tableId === el.id ? sel : null;

  // Rows size to their content, so the element's height follows the rendered
  // table and the overlays get the *real* row positions (fractions of the table
  // height) instead of guessing from the record. Only the full-size interactive
  // render drives this — thumbnails would measure a scaled/clipped box.
  const tableRef = useRef<HTMLTableElement>(null);
  const syncRef = useRef<() => void>(() => {});
  syncRef.current = () => {
    const node = tableRef.current;
    if (!interactive || !node) return;
    const rect = node.getBoundingClientRect();
    const rows = node.tBodies[0] ? Array.from(node.tBodies[0].rows) : [];
    const tops = rect.height > 0 ? rows.map((tr) => (tr.getBoundingClientRect().top - rect.top) / rect.height) : [];
    tops.push(1);
    tableMetrics.set(el.id, tops);
    // Don't fight a live gesture (resize/move owns the frame until it commits).
    if (!transient.get().liveFrames?.has(el.id)) deck.syncElementHeight(el.id, node.offsetHeight);
  };
  // Re-measure on content reflow (ResizeObserver) and after every render — the
  // latter re-syncs the height when the frame itself changed (e.g. a resize).
  useEffect(() => {
    if (!interactive || typeof ResizeObserver === 'undefined') return;
    const node = tableRef.current;
    if (!node) return;
    const ro = new ResizeObserver(() => syncRef.current());
    ro.observe(node);
    return () => { ro.disconnect(); tableMetrics.clear(el.id); };
  }, [tableMetrics, el.id, interactive]);
  useLayoutEffect(() => { syncRef.current(); });

  const border = el.border
    ? `${Math.max(1, el.border.width)}px solid ${resolveColor(el.border.color, theme)}`
    : `1px solid ${resolveColor({ kind: 'theme', slot: 'dk2' } as Color, theme)}`;

  return (
    <table
      ref={tableRef}
      data-table-id={el.id}
      style={{ width: '100%', borderCollapse: 'collapse', tableLayout: 'fixed', fontFamily: 'inherit' }}
    >
      <colgroup>
        {el.colFractions.map((f, i) => (
          <col key={i} style={{ width: `${f * 100}%` }} />
        ))}
      </colgroup>
      <tbody>
        {el.cells.map((row, r) => (
          // A row's manual minimum height (px); content can still push it taller.
          <tr key={r} style={{ height: el.rowHeights?.[r] ? el.rowHeights[r] : undefined }}>
            {row.map((cell, c) => {
              const bg = cell.fill ? resolveFill(cell.fill, theme) : null;
              const isEditing = !!editingCell && editingCell[0] === r && editingCell[1] === c;
              const isSelected = !!cellSel && !isEditing && inSelection(cellSel, r, c);
              // Fill in cell defaults so the static <td> and the live editor
              // agree on padding, size, and vertical alignment (no jump on edit).
              const cellStyle = {
                padding: cell.bodyStyle?.padding ?? CELL_PAD,
                fontSize: cell.bodyStyle?.fontSize ?? CELL_FONT,
                vAlign: cell.bodyStyle?.vAlign ?? 'top',
                ...cell.bodyStyle,
              };
              return (
                <td
                  key={c}
                  data-cell={`${r},${c}`}
                  style={{
                    border,
                    verticalAlign: cellStyle.vAlign,
                    fontSize: cellStyle.fontSize,
                    background: bg ?? undefined,
                    boxShadow: isSelected ? SELECT_TINT : undefined,
                    padding: isEditing ? 0 : cellStyle.padding,
                    overflow: 'hidden',
                    boxSizing: 'border-box',
                  }}
                >
                  {isEditing ? (
                    <RichTextEditor elementId={el.id} theme={theme} cell={[r, c]} bodyStyle={cellStyle} />
                  ) : (
                    <StaticRichText doc={cell.richText} theme={theme} />
                  )}
                </td>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
