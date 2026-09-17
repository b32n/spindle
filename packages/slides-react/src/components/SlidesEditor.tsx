// SlidesEditor — the root editor layout: toolbar, filmstrip, and the
// interactive stage. Owns keyboard shortcuts (attached to a focused wrapper,
// not window) and the right-click context menu.

import React, { useEffect, useRef, useState } from 'react';
import { useDeck, useKeyboardShortcuts, useCommentsOpen, useFilmstripOpen } from '../hooks';
import { usePasteImport } from '../hooks/usePasteImport';
import { useDeckContext } from '../context/DeckContext';
import { Toolbar, type Zoom } from './Toolbar';
import { Filmstrip } from './Filmstrip';
import { SlideStage } from './SlideStage';
import { NotesPanel } from './NotesPanel';
import { ContextMenu } from './ContextMenu';
import { SlideContextMenu } from './SlideContextMenu';
import { TableCellMenu } from './TableCellMenu';
import { PresentMode } from './PresentMode';
import { CommentsPanel } from './CommentsPanel';

export interface SlidesEditorProps {
  style?: React.CSSProperties;
  /** Read-only viewer (no toolbar, gestures, or shortcuts). */
  readOnly?: boolean;
  /**
   * Extra buttons pinned at the end of the toolbar, before Present. App-level
   * concerns like PDF/PNG export live here — export is intentionally kept out
   * of this package; the host wires it up from the public render API.
   */
  headerActions?: React.ReactNode;
  /**
   * Extra controls appended to the toolbar's scrollable region (after the
   * format bars, before Present). Same host-injection story as
   * headerActions — app-specific toolbar buttons.
   */
  toolbarExtras?: React.ReactNode;
}

const MOBILE_MAX_W = 640;

// Keyframes for the mobile filmstrip drawer (injected once).
function ensureDrawerStyles(): void {
  if (typeof document === 'undefined' || document.getElementById('spindle-slides-drawer-css')) return;
  const el = document.createElement('style');
  el.id = 'spindle-slides-drawer-css';
  el.textContent =
    '@keyframes spindle-drawer-fade{from{opacity:0}to{opacity:1}}' +
    '@keyframes spindle-drawer-in{from{transform:translateX(-100%)}to{transform:translateX(0)}}';
  document.head.appendChild(el);
}

export function SlidesEditor({ style, readOnly = false, headerActions, toolbarExtras }: SlidesEditorProps): React.ReactElement {
  const deck = useDeck();
  const [zoom, setZoom] = useState<Zoom>('fit');
  const [menu, setMenu] = useState<
    | { x: number; y: number; kind: 'element' }
    | { x: number; y: number; kind: 'slide'; slideId: string }
    | { x: number; y: number; kind: 'tableCell'; tableId: string; row: number; col: number }
    | null
  >(null);
  const [presenting, setPresenting] = useState(false);
  const { ui, tableSel } = useDeckContext();
  const showComments = useCommentsOpen();
  const filmstripOpen = useFilmstripOpen();
  const { onKeyDown } = useKeyboardShortcuts();
  const { onPaste } = usePasteImport();
  const rootRef = useRef<HTMLDivElement>(null);
  const [isMobile, setIsMobile] = useState(
    () => typeof window !== 'undefined' && !!window.matchMedia?.(`(max-width: ${MOBILE_MAX_W}px)`).matches
  );

  // Focus the editor so keyboard shortcuts work without an explicit click.
  useEffect(() => {
    if (!readOnly) rootRef.current?.focus();
  }, [readOnly]);

  useEffect(() => { ensureDrawerStyles(); }, []);

  // The filmstrip eats horizontal room that's scarce on phones, so its default
  // follows the breakpoint: shown on desktop, hidden on mobile (where it opens
  // as a slide-over drawer). Crossing the breakpoint resets to that default; an
  // explicit toggle in between still wins until the next crossing.
  useEffect(() => {
    if (typeof window === 'undefined' || !window.matchMedia) return;
    const mq = window.matchMedia(`(max-width: ${MOBILE_MAX_W}px)`);
    const apply = () => {
      setIsMobile(mq.matches);
      ui.setFilmstripOpen(!mq.matches);
    };
    apply();
    mq.addEventListener('change', apply);
    return () => mq.removeEventListener('change', apply);
  }, [ui]);

  return (
    <div
      ref={rootRef}
      tabIndex={readOnly ? undefined : 0}
      onKeyDown={readOnly ? undefined : onKeyDown}
      onPaste={readOnly ? undefined : onPaste}
      onContextMenu={(e) => {
        if (readOnly) return;
        const target = e.target as HTMLElement;
        // A slide thumbnail in the filmstrip → slide menu.
        const thumb = target.closest('[data-slide-thumb]');
        if (thumb) {
          e.preventDefault();
          setMenu({ x: e.clientX, y: e.clientY, kind: 'slide', slideId: thumb.getAttribute('data-slide-thumb')! });
          return;
        }
        if (target.closest('[data-slide-stage]')) {
          e.preventDefault();
          // A table cell, or a row/column header gutter, → the table menu
          // (insert/delete rows·cols, fill). Otherwise the generic element menu.
          // Elsewhere (toolbar/header/notes) → native.
          const cellEl = target.closest('[data-cell]');
          const elEl = target.closest('[data-element-id]');
          const cellTableId = elEl?.getAttribute('data-element-id');
          const rowSelEl = target.closest('[data-row-select]');
          const colSelEl = target.closest('[data-col-select]');
          const gripEl = target.closest('[data-table-move]');
          // Gutters live in an overlay (no data-cell ancestor); they only show
          // for the single selected table, so read it from the selection.
          const gutterTableId = deck.getSelection().elementIds[0];
          const gutterTable = gutterTableId ? deck.getElement(gutterTableId) : undefined;

          if (cellEl && cellTableId && deck.getElement(cellTableId)?.type === 'table') {
            const [r, c] = cellEl.getAttribute('data-cell')!.split(',').map(Number);
            setMenu({ x: e.clientX, y: e.clientY, kind: 'tableCell', tableId: cellTableId, row: r, col: c });
          } else if ((rowSelEl || colSelEl || gripEl) && gutterTable?.type === 'table') {
            // Select the whole row/column first so the menu's fill/delete target
            // it; the corner grip targets the top-left cell.
            let r = 0, c = 0;
            if (rowSelEl) { r = Number(rowSelEl.getAttribute('data-row-select')); tableSel.set(gutterTableId!, [r, 0], [r, gutterTable.cols - 1]); }
            else if (colSelEl) { c = Number(colSelEl.getAttribute('data-col-select')); tableSel.set(gutterTableId!, [0, c], [gutterTable.rows - 1, c]); }
            setMenu({ x: e.clientX, y: e.clientY, kind: 'tableCell', tableId: gutterTableId!, row: r, col: c });
          } else {
            setMenu({ x: e.clientX, y: e.clientY, kind: 'element' });
          }
        }
      }}
      style={{
        display: 'flex',
        flexDirection: 'column',
        height: '100%',
        position: 'relative',
        outline: 'none',
        fontFamily: 'Inter, system-ui, sans-serif',
        color: '#1f2933',
        ...style,
      }}
    >
      {!readOnly && (
        <Toolbar
          extras={toolbarExtras}
          headerActions={headerActions}
          zoom={zoom}
          onZoomChange={setZoom}
          onPresent={() => setPresenting(true)}
        />
      )}
      <div style={{ display: 'flex', flex: '1 1 auto', minHeight: 0, gap: 12, padding: '4px 12px 12px', background: 'linear-gradient(180deg, #f1f5f9 0%, #eaeef4 100%)' }}>
        {/* Desktop: filmstrip sits inline in the layout, taking real width. */}
        {filmstripOpen && !isMobile && <Filmstrip />}
        <div style={{ display: 'flex', flexDirection: 'column', flex: '1 1 auto', minWidth: 0, gap: 12 }}>
          <SlideStage zoom={zoom === 'fit' ? undefined : zoom} interactive={!readOnly} onZoomChange={setZoom} />
          {!readOnly && <NotesPanel />}
        </div>
        {!readOnly && showComments && <CommentsPanel onClose={() => ui.setCommentsOpen(false)} />}
      </div>
      {/* Mobile: filmstrip opens as a slide-over drawer over the stage. Tapping
          the backdrop — or a thumbnail (the click bubbles) — closes it. */}
      {filmstripOpen && isMobile && (
        <div
          onClick={() => ui.setFilmstripOpen(false)}
          style={{ position: 'absolute', inset: 0, zIndex: 60, background: 'rgba(15,23,42,0.35)', animation: 'spindle-drawer-fade .18s ease' }}
        >
          <div
            style={{
              position: 'absolute', top: 0, left: 0, bottom: 0, width: 'min(78vw, 240px)',
              background: '#eef1f5', boxShadow: '2px 0 16px rgba(15,23,42,0.25)', overflowY: 'auto',
              padding: 8, animation: 'spindle-drawer-in .2s cubic-bezier(.2,.7,.3,1)',
            }}
          >
            <Filmstrip />
          </div>
        </div>
      )}
      {menu?.kind === 'element' && <ContextMenu x={menu.x} y={menu.y} onClose={() => setMenu(null)} />}
      {menu?.kind === 'slide' && <SlideContextMenu slideId={menu.slideId} x={menu.x} y={menu.y} onClose={() => setMenu(null)} />}
      {menu?.kind === 'tableCell' && <TableCellMenu tableId={menu.tableId} row={menu.row} col={menu.col} x={menu.x} y={menu.y} onClose={() => setMenu(null)} />}
      {presenting && <PresentMode onExit={() => setPresenting(false)} />}
    </div>
  );
}
