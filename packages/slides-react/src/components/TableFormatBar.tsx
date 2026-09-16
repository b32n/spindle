// TableFormatBar — inline table controls in the main Toolbar row when a single
// table is selected: insert/delete rows and columns relative to the current
// cell selection, fill the selected cells, and set the grid border color +
// width. (The same insert/delete live on the cell right-click menu too.)
// Column/row resize is a drag on the cell boundaries.

import React, { useSyncExternalStore } from 'react';
import {
  BetweenHorizontalStart, BetweenHorizontalEnd, BetweenVerticalStart, BetweenVerticalEnd,
  Trash2, PaintBucket,
} from 'lucide-react';
import { resolveColor, type Color, type Stroke, type TableElement } from '@b32nio/spindle-slides-core';
import { useDeck, useTheme, useSelection, useElement } from '../hooks';
import { useDeckContext } from '../context/DeckContext';
import { cellsInSelection, selectionRect } from '../interactions/table-selection-store';
import { ToolbarButton, ToolbarDivider } from './toolbarUI';

const field: React.CSSProperties = {
  height: 30, border: '1px solid rgba(0,0,0,0.1)', borderRadius: 8,
  background: 'rgba(255,255,255,0.8)', color: '#334155', padding: '0 6px',
  fontSize: 13, fontFamily: '"Inter", sans-serif', cursor: 'pointer',
};
const WIDTHS = [1, 2, 3, 4];

function colorToHex(color: Color, theme: ReturnType<typeof useTheme>): string {
  const s = resolveColor(color, theme);
  return /^#[0-9a-fA-F]{6}$/.test(s) ? s : '#888888';
}

export function TableFormatBar(): React.ReactElement | null {
  const deck = useDeck();
  const theme = useTheme();
  const selection = useSelection();
  const { tableSel } = useDeckContext();
  const sel = useSyncExternalStore(tableSel.subscribe, tableSel.getState);
  const id = selection.elementIds.length === 1 ? selection.elementIds[0] : null;
  const el = useElement(id ?? '') as TableElement | undefined;
  if (!id || !el || el.type !== 'table') return null;

  const border = el.border ?? { color: { kind: 'theme', slot: 'dk2' } as Color, width: 1 };
  const setBorder = (patch: Partial<Stroke>) => deck.updateElement(id, { border: { ...border, ...patch } });

  // Insert/delete are relative to the selected cells; with no selection they
  // act on the last row/column (append-ish), so the buttons always do something.
  const rect = sel && sel.tableId === id ? selectionRect(sel) : null;
  const r0 = rect ? rect.r0 : el.rows - 1;
  const r1 = rect ? rect.r1 : el.rows - 1;
  const c0 = rect ? rect.c0 : el.cols - 1;
  const c1 = rect ? rect.c1 : el.cols - 1;
  const canDeleteRows = r1 - r0 + 1 < el.rows;
  const canDeleteCols = c1 - c0 + 1 < el.cols;

  // Fill applies to the current cell-range selection (a row / column / block);
  // without one there's nothing to fill, so the control is disabled.
  const cells = sel && sel.tableId === id ? cellsInSelection(sel) : [];
  const canFill = cells.length > 0;
  const firstCell = canFill ? el.cells[cells[0][0]]?.[cells[0][1]] : undefined;
  const fillHex = firstCell?.fill?.kind === 'solid' && /^#[0-9a-fA-F]{6}$/.test(resolveColor(firstCell.fill.color, theme))
    ? resolveColor(firstCell.fill.color, theme)
    : '#ffffff';
  const setFill = (fill: TableElement['cells'][number][number]['fill']) => deck.updateTableCells(id, cells, { fill });

  return (
    <>
      <ToolbarDivider />
      <ToolbarButton title="Insert row above" onClick={() => deck.insertTableRow(id, r0)}>
        <BetweenHorizontalStart size={15} />
      </ToolbarButton>
      <ToolbarButton title="Insert row below" onClick={() => deck.insertTableRow(id, r1 + 1)}>
        <BetweenHorizontalEnd size={15} />
      </ToolbarButton>
      <ToolbarButton title="Delete row(s)" disabled={!canDeleteRows} onClick={() => deck.removeTableRows(id, r0, r1)}>
        <Trash2 size={14} />
      </ToolbarButton>
      <ToolbarDivider />
      <ToolbarButton title="Insert column left" onClick={() => deck.insertTableColumn(id, c0)}>
        <BetweenVerticalStart size={15} />
      </ToolbarButton>
      <ToolbarButton title="Insert column right" onClick={() => deck.insertTableColumn(id, c1 + 1)}>
        <BetweenVerticalEnd size={15} />
      </ToolbarButton>
      <ToolbarButton title="Delete column(s)" disabled={!canDeleteCols} onClick={() => deck.removeTableColumns(id, c0, c1)}>
        <Trash2 size={14} />
      </ToolbarButton>
      <ToolbarDivider />
      <PaintBucket size={14} style={{ color: canFill ? '#64748b' : '#cbd5e1' }} aria-hidden />
      <input
        type="color"
        title={canFill ? 'Cell fill (selected cells)' : 'Select cells to fill'}
        disabled={!canFill}
        value={fillHex}
        onMouseDown={(e) => e.preventDefault()}
        onChange={(e) => setFill({ kind: 'solid', color: { kind: 'rgb', hex: e.target.value } })}
        style={{ ...field, width: 34, padding: 3, opacity: canFill ? 1 : 0.4, cursor: canFill ? 'pointer' : 'default' }}
      />
      <ToolbarButton title="Clear fill" disabled={!canFill} onClick={() => setFill(undefined)}>
        <span style={{ fontSize: 13, fontWeight: 700 }}>⦸</span>
      </ToolbarButton>
      <ToolbarDivider />
      <input
        type="color"
        title="Grid line color"
        value={colorToHex(border.color, theme)}
        onMouseDown={(e) => e.preventDefault()}
        onChange={(e) => setBorder({ color: { kind: 'rgb', hex: e.target.value } })}
        style={{ ...field, width: 34, padding: 3 }}
      />
      <select
        title="Grid line width"
        value={String(border.width)}
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => setBorder({ width: Number(e.target.value) })}
        style={field}
      >
        {WIDTHS.map((w) => (
          <option key={w} value={w}>{w}px</option>
        ))}
      </select>
    </>
  );
}
