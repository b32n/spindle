// TableCellMenu — the right-click menu for a table, shared by cell right-clicks
// and the row/column header gutters. Insert/delete and fill act on the current
// cell selection when the right-clicked cell is inside it (so right-clicking a
// selected row fills/deletes the whole row); otherwise they act on the single
// [row, col] that was clicked.

import React, { useSyncExternalStore } from 'react';
import { resolveColor, type Color, type TableElement } from '@b32nio/spindle-slides-core';
import { useDeck, useTheme, useElement } from '../hooks';
import { useDeckContext } from '../context/DeckContext';
import { selectionRect, inSelection, cellsInSelection } from '../interactions/table-selection-store';
import { Menu, type MenuItem } from './Menu';

export interface TableCellMenuProps {
  tableId: string;
  row: number;
  col: number;
  x: number;
  y: number;
  onClose: () => void;
}

function fillHex(color: Color | undefined, theme: ReturnType<typeof useTheme>): string {
  if (!color) return '#ffffff';
  const s = resolveColor(color, theme);
  return /^#[0-9a-fA-F]{6}$/.test(s) ? s : '#ffffff';
}

export function TableCellMenu({ tableId, row, col, x, y, onClose }: TableCellMenuProps): React.ReactElement | null {
  const deck = useDeck();
  const theme = useTheme();
  const { tableSel } = useDeckContext();
  const sel = useSyncExternalStore(tableSel.subscribe, tableSel.getState);
  const el = useElement(tableId) as TableElement | undefined;
  if (!el || el.type !== 'table') return null;

  // Operate on the selection when the clicked cell is inside it, else the cell.
  const rect = sel && sel.tableId === tableId && inSelection(sel, row, col) ? selectionRect(sel) : null;
  const r0 = rect ? rect.r0 : row, r1 = rect ? rect.r1 : row;
  const c0 = rect ? rect.c0 : col, c1 = rect ? rect.c1 : col;
  const rowSpan = r1 - r0 + 1, colSpan = c1 - c0 + 1;
  const targetCells: Array<[number, number]> = rect ? cellsInSelection(sel!) : [[row, col]];
  const cell = el.cells[row]?.[col];
  const setFill = (fill: TableElement['cells'][number][number]['fill']) => deck.updateTableCells(tableId, targetCells, { fill });

  const items: Array<MenuItem | 'sep'> = [
    { label: 'Insert row above', run: () => deck.insertTableRow(tableId, r0) },
    { label: 'Insert row below', run: () => deck.insertTableRow(tableId, r1 + 1) },
    { label: 'Insert column left', run: () => deck.insertTableColumn(tableId, c0) },
    { label: 'Insert column right', run: () => deck.insertTableColumn(tableId, c1 + 1) },
    'sep',
    { label: rowSpan > 1 ? `Delete ${rowSpan} rows` : 'Delete row', run: () => deck.removeTableRows(tableId, r0, r1), disabled: rowSpan >= el.rows },
    { label: colSpan > 1 ? `Delete ${colSpan} columns` : 'Delete column', run: () => deck.removeTableColumns(tableId, c0, c1), disabled: colSpan >= el.cols },
  ];

  const fillLabel = rect ? (rowSpan > 1 || colSpan > 1 ? 'Fill selection' : 'Cell fill') : 'Cell fill';
  const footer = (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '6px 14px', fontSize: 13, color: '#2b3440' }}>
      <span>{fillLabel}</span>
      <input
        type="color"
        title={fillLabel}
        value={fillHex(cell?.fill?.kind === 'solid' ? cell.fill.color : undefined, theme)}
        onChange={(e) => setFill({ kind: 'solid', color: { kind: 'rgb', hex: e.target.value } })}
        style={{ width: 30, height: 26, border: '1px solid #d5d9e0', borderRadius: 6, padding: 2, cursor: 'pointer' }}
      />
      <button
        onPointerDown={(e) => { e.stopPropagation(); setFill(undefined); onClose(); }}
        style={{ marginLeft: 'auto', border: '1px solid #d5d9e0', background: '#fff', borderRadius: 5, padding: '3px 8px', fontSize: 12, color: '#3e4c59', cursor: 'pointer' }}
      >
        Clear
      </button>
    </div>
  );

  return <Menu x={x} y={y} items={items} onClose={onClose} footer={footer} />;
}
