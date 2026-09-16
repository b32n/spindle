// Remote-selection overlay for the canvas grid.
//
// Reads workbook.getCollabHandle().awareness, watches for peer cell
// selections (set under awareness.cellSelection by the local mirror in
// this same component), and paints absolute-positioned colored rectangles
// + name labels above the canvas. Position math mirrors the canvas
// renderer's: walk getRowHeight / getColWidth for indices, subtract
// scroll, add header offsets.
//
// One responsibility kept here on purpose: writing the LOCAL selection
// state to awareness. That keeps all the collab-selection glue in one
// file instead of threading another effect through CanvasGrid.

import React, { useEffect, useRef, useState } from 'react';
import type { Sheet, Range, Selection, WorkbookImpl } from '@b32nio/spindle-sheets-core';
import type { Awareness } from 'y-protocols/awareness';

interface RemotePeer {
  clientId: number;
  user: { name: string; color: string; userId: string };
  cellSelection: {
    sheetId: string;
    activeCell: { row: number; col: number } | null;
    ranges: Range[];
  };
}

export interface RemoteSelectionOverlayProps {
  workbook: WorkbookImpl;
  sheet: Sheet;
  /** Local active cell + ranges (mirrored to awareness on change). */
  activeCell: { row: number; col: number } | null;
  selection: Selection;
  scrollTop: number;
  scrollLeft: number;
  headerWidth: number;
  headerHeight: number;
  /** Total canvas dimensions for clipping checks. */
  width: number;
  height: number;
}

export function RemoteSelectionOverlay({
  workbook,
  sheet,
  activeCell,
  selection,
  scrollTop,
  scrollLeft,
  headerWidth,
  headerHeight,
  width,
  height,
}: RemoteSelectionOverlayProps): React.ReactElement | null {
  const [peers, setPeers] = useState<RemotePeer[]>([]);
  const awarenessRef = useRef<Awareness | null>(null);

  // Track the live Awareness instance from the collab handle.
  useEffect(() => {
    const handle = workbook.getCollabHandle();
    awarenessRef.current = handle?.awareness ?? null;
    if (!handle) {
      setPeers([]);
      return;
    }
    const refresh = (): void => {
      const localId = handle.ydoc.clientID;
      const out: RemotePeer[] = [];
      handle.awareness.getStates().forEach((state, clientId) => {
        if (clientId === localId) return;
        const user = state.user as RemotePeer['user'] | undefined;
        const cellSelection = state.cellSelection as RemotePeer['cellSelection'] | undefined;
        if (!user || !cellSelection) return;
        out.push({ clientId, user, cellSelection });
      });
      setPeers(out);
    };
    handle.awareness.on('change', refresh);
    refresh();
    return () => {
      handle.awareness.off('change', refresh);
    };
  }, [workbook]);

  // Mirror our local selection into our own awareness state so peers see us.
  useEffect(() => {
    const aw = awarenessRef.current;
    if (!aw) return;
    aw.setLocalStateField('cellSelection', {
      sheetId: sheet.id,
      activeCell,
      ranges: selection.ranges,
    });
  }, [sheet, activeCell, selection, peers /* re-mirror after a peer joins */]);

  // No collab attached or nobody else around — render nothing.
  if (peers.length === 0) return null;

  const visibles = peers.filter((p) => p.cellSelection.sheetId === sheet.id);
  if (visibles.length === 0) return null;

  return (
    <div
      style={{
        position: 'absolute',
        top: 0,
        left: 0,
        width,
        height,
        pointerEvents: 'none',
        overflow: 'hidden',
        zIndex: 5,
      }}
    >
      {/* Inner clip area: rectangles + outlines never draw over the
          row-number column or column-letter row when scrolled. */}
      <div
        style={{
          position: 'absolute',
          top: headerHeight,
          left: headerWidth,
          width: Math.max(0, width - headerWidth),
          height: Math.max(0, height - headerHeight),
          overflow: 'hidden',
        }}
      >
        {visibles.map((peer) =>
          renderPeerRects(peer, sheet, scrollTop, scrollLeft, headerWidth, headerHeight),
        )}
      </div>
      {/* Name pills sit OUTSIDE the clip area so they can hang above the
          frozen / header row without getting cropped. */}
      {visibles.map((peer) =>
        renderPeerLabel(peer, sheet, scrollTop, scrollLeft, headerWidth, headerHeight),
      )}
    </div>
  );
}

/**
 * Selection fills + active-cell outline. Coordinates are translated to be
 * relative to the inner clip area (which starts at headerWidth, headerHeight).
 */
function renderPeerRects(
  peer: RemotePeer,
  sheet: Sheet,
  scrollTop: number,
  scrollLeft: number,
  headerWidth: number,
  headerHeight: number,
): React.ReactNode {
  const color = peer.user.color;
  const els: React.ReactNode[] = [];

  // Translucent fill for each range
  for (let i = 0; i < peer.cellSelection.ranges.length; i++) {
    const r = peer.cellSelection.ranges[i];
    const rect = cellRangeRect(sheet, r, scrollTop, scrollLeft, headerWidth, headerHeight);
    if (!rect) continue;
    els.push(
      <div
        key={`r-${peer.clientId}-${i}`}
        style={{
          position: 'absolute',
          left: rect.x - headerWidth,
          top: rect.y - headerHeight,
          width: rect.w,
          height: rect.h,
          backgroundColor: color,
          opacity: 0.18,
        }}
      />,
    );
  }

  if (peer.cellSelection.activeCell) {
    const ac = peer.cellSelection.activeCell;
    const rect = cellRect(sheet, ac.row, ac.col, scrollTop, scrollLeft, headerWidth, headerHeight);
    if (rect) {
      els.push(
        <div
          key={`a-${peer.clientId}`}
          style={{
            position: 'absolute',
            left: rect.x - headerWidth,
            top: rect.y - headerHeight,
            width: rect.w,
            height: rect.h,
            border: `2px solid ${color}`,
            boxSizing: 'border-box',
          }}
        />,
      );
    }
  }

  return els;
}

/**
 * Name pill above the active cell. Lives outside the clip area so it can
 * draw on top of the header row (which is the natural visual treatment for
 * remote-user labels). Hidden when the active cell is fully scrolled out
 * of the data area.
 */
function renderPeerLabel(
  peer: RemotePeer,
  sheet: Sheet,
  scrollTop: number,
  scrollLeft: number,
  headerWidth: number,
  headerHeight: number,
): React.ReactNode {
  if (!peer.cellSelection.activeCell) return null;
  const ac = peer.cellSelection.activeCell;
  const rect = cellRect(sheet, ac.row, ac.col, scrollTop, scrollLeft, headerWidth, headerHeight);
  if (!rect) return null;
  // Hide if the cell's right edge is to the left of the data area or its
  // bottom is above the data area — the cell is fully scrolled away.
  if (rect.x + rect.w <= headerWidth) return null;
  if (rect.y + rect.h <= headerHeight) return null;
  return (
    <div
      key={`l-${peer.clientId}`}
      style={{
        position: 'absolute',
        left: Math.max(headerWidth, rect.x),
        top: rect.y - 18,
        backgroundColor: peer.user.color,
        color: 'white',
        padding: '1px 6px',
        borderRadius: 3,
        fontSize: 11,
        fontWeight: 500,
        fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
        whiteSpace: 'nowrap',
        lineHeight: 1.4,
      }}
    >
      {peer.user.name}
    </div>
  );
}

// ============================================================================
// Coordinate helpers
// ============================================================================

interface CellRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

function cellRect(
  sheet: Sheet,
  row: number,
  col: number,
  scrollTop: number,
  scrollLeft: number,
  headerWidth: number,
  headerHeight: number,
): CellRect | null {
  if (row < 0 || col < 0) return null;
  const frozenRows = sheet.getFrozenRows();
  const frozenCols = sheet.getFrozenCols();

  // Rows in the frozen band stay pinned at headerHeight regardless of
  // scrollTop. Rows below the band scroll normally and offset past the
  // frozen height.
  let y: number;
  if (row < frozenRows) {
    y = headerHeight;
    for (let r = 0; r < row; r++) y += sheet.getRowHeight(r);
  } else {
    let frozenSum = 0;
    for (let r = 0; r < frozenRows; r++) frozenSum += sheet.getRowHeight(r);
    y = headerHeight + frozenSum - scrollTop;
    for (let r = frozenRows; r < row; r++) y += sheet.getRowHeight(r);
  }

  let x: number;
  if (col < frozenCols) {
    x = headerWidth;
    for (let c = 0; c < col; c++) x += sheet.getColWidth(c);
  } else {
    let frozenColSum = 0;
    for (let c = 0; c < frozenCols; c++) frozenColSum += sheet.getColWidth(c);
    x = headerWidth + frozenColSum - scrollLeft;
    for (let c = frozenCols; c < col; c++) x += sheet.getColWidth(c);
  }

  return {
    x,
    y,
    w: sheet.getColWidth(col),
    h: sheet.getRowHeight(row),
  };
}

function cellRangeRect(
  sheet: Sheet,
  range: Range,
  scrollTop: number,
  scrollLeft: number,
  headerWidth: number,
  headerHeight: number,
): CellRect | null {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  const a = cellRect(sheet, startRow, startCol, scrollTop, scrollLeft, headerWidth, headerHeight);
  if (!a) return null;
  let w = 0;
  for (let c = startCol; c <= endCol; c++) w += sheet.getColWidth(c);
  let h = 0;
  for (let r = startRow; r <= endRow; r++) h += sheet.getRowHeight(r);
  return { x: a.x, y: a.y, w, h };
}
