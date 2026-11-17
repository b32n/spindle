// ResponsiveToolbar — lays its children out in a single full-width row. When
// they don't all fit, the trailing ones collapse into an overflow surface that
// is PORTALED to <body> (so it can't be clipped or covered by the editor's own
// stacking contexts). On desktop that surface is a dropdown under the "⋯"; on
// mobile (≤640px) the whole toolbar collapses to a "Tools" button that opens a
// bottom sheet holding ALL the tools.
//
// The overflow surface is a structured vertical list: each control on its own
// row with its label (from the child's title/tooltip). The whole row is
// clickable — a tap on the label forwards to the control — so labels aren't
// dead. The surface measures its own rows, so controls that render nothing
// (an inactive contextual format bar) collapse away and reappear live.

import React, { useCallback, useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from 'react';
import { createPortal } from 'react-dom';
import { ToolbarMenuContext, type ToolbarMenuApi } from './toolbar-menu';

export interface ResponsiveToolbarProps {
  /** Toolbar items, in priority order (earliest stay inline the longest). */
  children: React.ReactNode;
  /** Gap between items, px (should match the toolbar's own gap). */
  gap?: number;
  /** Icon/content for the overflow trigger (defaults to a "⋯" glyph). */
  moreIcon?: React.ReactNode;
  moreLabel?: string;
  style?: React.CSSProperties;
  className?: string;
}

// Rows narrower than this (a divider ≈13px, or an empty/inactive control = 0)
// are hidden in the overflow surface.
const MIN_ROW_W = 18;
const MOBILE_MAX_W = 640;
const Z = 2147483000; // above any editor overlay

function subscribeMobile(onChange: () => void): () => void {
  if (typeof window === 'undefined' || !window.matchMedia) return () => {};
  const mq = window.matchMedia(`(max-width: ${MOBILE_MAX_W}px)`);
  mq.addEventListener('change', onChange);
  return () => mq.removeEventListener('change', onChange);
}

function isMobileNow(): boolean {
  return typeof window !== 'undefined' && !!window.matchMedia && window.matchMedia(`(max-width: ${MOBILE_MAX_W}px)`).matches;
}

const moreButtonStyle: React.CSSProperties = {
  display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 7,
  height: 32, padding: '0 8px', flex: '0 0 auto', border: 'none', borderRadius: 8,
  background: 'transparent', color: '#64748b', cursor: 'pointer', fontSize: 18,
  lineHeight: 1, fontFamily: 'inherit', fontWeight: 500,
};

const panelBase: React.CSSProperties = {
  display: 'flex', flexDirection: 'column', gap: 1,
  background: '#fff', boxSizing: 'border-box', fontFamily: 'inherit',
  overflowY: 'auto', overflowX: 'hidden',
};

function flattenChildren(children: React.ReactNode): React.ReactNode[] {
  const out: React.ReactNode[] = [];
  React.Children.forEach(children, (child) => {
    if (React.isValidElement(child) && child.type === React.Fragment) {
      out.push(...flattenChildren((child.props as { children?: React.ReactNode }).children));
    } else if (child !== null && child !== undefined && typeof child !== 'boolean') {
      out.push(child);
    }
  });
  return out;
}

function getLabel(item: React.ReactNode): string | null {
  if (!React.isValidElement(item)) return null;
  const p = item.props as Record<string, unknown>;
  const cand = p.tooltip ?? p.title ?? p['aria-label'];
  return typeof cand === 'string' && cand.trim() ? cand : null;
}

function ensureStyles(): void {
  if (typeof document === 'undefined' || document.getElementById('sp-rt-styles')) return;
  const s = document.createElement('style');
  s.id = 'sp-rt-styles';
  s.textContent =
    '.sp-rt-menuitem{transition:background .12s ease}' +
    '.sp-rt-menuitem:hover{background:rgba(99,102,241,0.10)}' +
    '.sp-rt-more:hover{background:rgba(99,102,241,0.08);color:#6366f1}' +
    '@keyframes sp-rt-pop{from{opacity:0;transform:translateY(-6px)}to{opacity:1;transform:translateY(0)}}' +
    '@keyframes sp-rt-sheet{from{transform:translateY(100%)}to{transform:translateY(0)}}' +
    '@keyframes sp-rt-fade{from{opacity:0}to{opacity:1}}' +
    '@keyframes sp-rt-fadeout{from{opacity:1}to{opacity:0}}';
  document.head.appendChild(s);
}

export function ResponsiveToolbar({
  children, gap = 2, moreIcon, moreLabel = 'More tools', style, className,
}: ResponsiveToolbarProps): React.ReactElement {
  const items = flattenChildren(children);
  const n = items.length;

  const containerRef = useRef<HTMLDivElement>(null);
  const itemRefs = useRef<Array<HTMLDivElement | null>>([]);
  const moreRef = useRef<HTMLButtonElement>(null);
  const widths = useRef<number[]>([]);
  const moreW = useRef(40);
  const rowRefs = useRef<Record<number, HTMLElement | null>>({});
  const [rowW, setRowW] = useState<Record<number, number>>({});

  const isMobile = useSyncExternalStore(subscribeMobile, isMobileNow, () => false);
  const [visible, setVisible] = useState(() => (isMobileNow() ? 0 : n));
  const [open, setOpen] = useState(false);
  const [anchor, setAnchor] = useState<{ top: number; right: number } | null>(null);
  const [modal, setModal] = useState<{ title: string; content: React.ReactNode } | null>(null);
  const [modalClosing, setModalClosing] = useState(false);

  // Fade the modal out, then unmount it (keeps the surface open).
  const dismissModal = useCallback(() => {
    setModalClosing(true);
    window.setTimeout(() => { setModal(null); setModalClosing(false); }, 140);
  }, []);

  const menuApi: ToolbarMenuApi = {
    inMenu: true,
    openModal: (title, content) => { setModalClosing(false); setModal({ title, content }); },
    closeModal: dismissModal,
    closeMenu: () => { setModal(null); setModalClosing(false); setOpen(false); },
  };

  useEffect(ensureStyles, []);

  // Desktop: measure inline items and compute how many fit. Mobile: everything
  // lives in the sheet, so nothing stays inline.
  const measure = useCallback(() => {
    const container = containerRef.current;
    if (!container) return;
    if (isMobile) { setVisible((v) => (v === 0 ? v : 0)); return; }
    for (let i = 0; i < itemRefs.current.length; i++) {
      const el = itemRefs.current[i];
      if (el) widths.current[i] = el.offsetWidth;
    }
    if (moreRef.current) moreW.current = moreRef.current.offsetWidth;
    const avail = container.clientWidth;
    let total = 0;
    for (let i = 0; i < n; i++) total += (widths.current[i] || 0) + (i > 0 ? gap : 0);
    let count: number;
    if (total <= avail + 0.5) {
      count = n;
    } else {
      const budget = avail - moreW.current - gap;
      let used = 0;
      count = 0;
      for (let i = 0; i < n; i++) {
        const add = (widths.current[i] || 0) + (count > 0 ? gap : 0);
        if (used + add <= budget) { used += add; count += 1; } else break;
      }
    }
    setVisible((prev) => (prev === count ? prev : count));
  }, [n, gap, isMobile]);

  // Layout measurement has to land before paint (otherwise items flash/overflow),
  // so the resulting setState must be synchronous here; it is guarded by the
  // equality check in `measure`, so it cannot loop.
  // eslint-disable-next-line react-hooks/set-state-in-effect
  useLayoutEffect(() => { measure(); });

  useEffect(() => {
    const container = containerRef.current;
    if (!container || typeof ResizeObserver === 'undefined') return;
    const ro = new ResizeObserver(() => measure());
    ro.observe(container);
    return () => ro.disconnect();
  }, [measure]);

  // On mobile everything overflows; on desktop only what didn't fit.
  const overflow = isMobile ? n > 0 : visible < n;
  const firstOverflow = isMobile ? 0 : visible;

  // Measure the surface's own rows while open, so empty (inactive) controls and
  // dividers collapse — and reappear when they become active.
  useLayoutEffect(() => {
    if (!open) return;
    const next = { ...rowW };
    let changed = false;
    for (let i = firstOverflow; i < n; i++) {
      const el = rowRefs.current[i];
      if (!el) continue;
      // Measure the natural content, not the (possibly flex-grown) wrapper: an
      // empty control has no child, a divider has one thin child.
      const w = el.childElementCount === 0 ? 0 : (el.firstElementChild as HTMLElement).offsetWidth;
      if (next[i] !== w) { next[i] = w; changed = true; }
    }
    // Synchronous so hidden rows never paint; the `changed` guard stops the loop.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    if (changed) setRowW(next);
    // Runs every render so live content changes (a format bar becoming active
    // while the surface is open) re-measure; the `changed` guard stops the loop.
     
  });

  // Anchor the desktop dropdown to the "⋯" (follows scroll/resize).
  useLayoutEffect(() => {
    if (!open || isMobile) return;
    const update = () => {
      const b = moreRef.current;
      if (!b) return;
      const r = b.getBoundingClientRect();
      setAnchor({ top: r.bottom + 6, right: Math.max(8, window.innerWidth - r.right) });
    };
    update();
    window.addEventListener('resize', update);
    window.addEventListener('scroll', update, true);
    return () => { window.removeEventListener('resize', update); window.removeEventListener('scroll', update, true); };
  }, [open, isMobile]);

  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') setOpen(false); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [open]);

  // A whole row is one click target: a tap anywhere forwards to the control
  // inside (so the label isn't dead), then closes the surface.
  const onRowClick = (e: React.MouseEvent) => {
    const control = (e.currentTarget as HTMLElement).querySelector('[data-rt-control]');
    const target = e.target as Node;
    if (control && !control.contains(target)) {
      const el = control.querySelector('button, [role="button"], a, input, select') as HTMLElement | null;
      if (el) { el.click(); }
    }
    setOpen(false);
  };

  const rows: React.ReactNode[] = [];
  for (let i = firstOverflow; i < n; i++) {
    const item = items[i];
    const w = rowW[i];
    const hidden = w !== undefined && w < MIN_ROW_W; // empty control or divider
    const label = getLabel(item);
    // Only a labelled single-action control is clickable as a whole row (tap the
    // label → run it → close). Unlabelled clusters (DeckControls) and dropdown-
    // openers (ShapePicker) keep their own controls live and leave the surface
    // open so their sub-menus can appear.
    const clickable = !hidden && !!label;
    rows.push(
      <div
        key={i}
        role={clickable ? 'menuitem' : undefined}
        className={clickable ? 'sp-rt-menuitem' : undefined}
        onClick={clickable ? onRowClick : undefined}
        style={hidden
          ? { height: 0, minHeight: 0, padding: 0, margin: 0, overflow: 'hidden' }
          : label
            ? { display: 'flex', alignItems: 'center', gap: 12, minHeight: 40, padding: '3px 12px 3px 8px', borderRadius: 8, cursor: 'pointer' }
            : { width: '100%' }}
      >
        <span
          data-rt-control
          ref={(el) => { rowRefs.current[i] = el; }}
          style={label
            ? { display: 'inline-flex', alignItems: 'center', flex: '0 0 auto' }
            // Menu-mode / cluster controls own their layout (they render their
            // own MenuRows via useToolbarMenu, or wrap their buttons).
            : { display: 'block', width: '100%' }}
        >
          {item}
        </span>
        {!hidden && label && <span style={{ fontSize: 14, color: '#334155', whiteSpace: 'nowrap', fontFamily: 'inherit' }}>{label}</span>}
      </div>
    );
  }

  const dropdown = anchor && (
    <>
      <div onClick={() => setOpen(false)} style={{ position: 'fixed', inset: 0, zIndex: Z }} />
      <div
        role="menu"
        style={{
          ...panelBase, position: 'fixed', top: anchor.top, right: anchor.right, zIndex: Z + 1,
          minWidth: 220, maxWidth: 340, maxHeight: '70vh', padding: 6, borderRadius: 12,
          border: '1px solid rgba(15,23,42,0.08)',
          boxShadow: '0 16px 40px -12px rgba(0,0,0,0.38), 0 2px 8px rgba(0,0,0,0.08)',
          animation: 'sp-rt-pop .12s ease',
        }}
      >
        <ToolbarMenuContext.Provider value={menuApi}>{rows}</ToolbarMenuContext.Provider>
      </div>
    </>
  );

  const sheet = (
    <>
      <div onClick={() => setOpen(false)} style={{ position: 'fixed', inset: 0, zIndex: Z, background: 'rgba(15,23,42,0.35)', animation: 'sp-rt-fade .15s ease' }} />
      <div
        role="menu"
        style={{
          ...panelBase, position: 'fixed', left: 0, right: 0, bottom: 0, zIndex: Z + 1,
          maxHeight: '80vh', padding: '6px 8px calc(16px + env(safe-area-inset-bottom))',
          borderTopLeftRadius: 18, borderTopRightRadius: 18,
          boxShadow: '0 -12px 40px -8px rgba(0,0,0,0.35)',
          animation: 'sp-rt-sheet .22s cubic-bezier(0.32,0.72,0,1)',
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', padding: '10px 6px 8px', position: 'sticky', top: 0, background: '#fff', zIndex: 1 }}>
          <div style={{ position: 'absolute', top: 4, left: '50%', transform: 'translateX(-50%)', width: 36, height: 4, borderRadius: 2, background: 'rgba(15,23,42,0.18)' }} />
          <span style={{ fontSize: 13, fontWeight: 600, color: '#64748b', letterSpacing: 0.2 }}>Tools</span>
          <button type="button" aria-label="Close" onClick={() => setOpen(false)} style={{ marginLeft: 'auto', border: 'none', background: 'transparent', color: '#64748b', fontSize: 22, lineHeight: 1, cursor: 'pointer', padding: 4 }}>×</button>
        </div>
        <ToolbarMenuContext.Provider value={menuApi}>{rows}</ToolbarMenuContext.Provider>
      </div>
    </>
  );

  // A control's options open here — portaled above the surface, so a shape grid
  // / colour swatches / theme list render cleanly regardless of the editor.
  const modalEl = modal && (
    <>
      <div onClick={dismissModal} style={{ position: 'fixed', inset: 0, zIndex: Z + 2, background: 'rgba(15,23,42,0.45)', animation: `${modalClosing ? 'sp-rt-fadeout' : 'sp-rt-fade'} .15s ease forwards` }} />
      <div
        role="dialog"
        aria-label={modal.title}
        style={{
          position: 'fixed', left: '50%', top: '50%', transform: 'translate(-50%, -50%)', zIndex: Z + 3,
          width: 'min(94vw, 440px)', maxHeight: '82vh', overflowY: 'auto', boxSizing: 'border-box',
          background: '#fff', borderRadius: 16, padding: 16, fontFamily: 'inherit',
          boxShadow: '0 24px 60px -16px rgba(0,0,0,0.45), 0 0 0 1px rgba(0,0,0,0.06)',
          // Opacity-only fade — a transform animation would fight the centering
          // translate(-50%,-50%) and make the modal flash off-centre first.
          animation: `${modalClosing ? 'sp-rt-fadeout' : 'sp-rt-fade'} .15s ease forwards`,
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12 }}>
          <span style={{ fontSize: 15, fontWeight: 600, color: '#1f2937' }}>{modal.title}</span>
          {/* Closes only the modal — back to the surface, which stays open. */}
          <button type="button" aria-label="Close" onClick={dismissModal} style={{ marginLeft: 'auto', border: 'none', background: 'transparent', color: '#64748b', fontSize: 22, lineHeight: 1, cursor: 'pointer', padding: 4 }}>×</button>
        </div>
        {modal.content}
      </div>
    </>
  );

  return (
    <div
      ref={containerRef}
      className={className}
      style={{ display: 'flex', alignItems: 'center', gap, width: '100%', minWidth: 0, ...style }}
    >
      {!isMobile && items.slice(0, visible).map((item, i) => (
        <div key={i} ref={(el) => { itemRefs.current[i] = el; }} style={{ display: 'inline-flex', alignItems: 'center', flex: '0 0 auto' }}>
          {item}
        </div>
      ))}

      {overflow && (
        <div style={{ marginLeft: isMobile ? undefined : 'auto', flex: '0 0 auto', display: 'inline-flex' }}>
          <button
            ref={moreRef}
            type="button"
            className="sp-rt-more"
            aria-label={moreLabel}
            title={moreLabel}
            aria-haspopup="menu"
            aria-expanded={open}
            onClick={() => setOpen((o) => !o)}
            style={moreButtonStyle}
          >
            {isMobile
              ? <><span aria-hidden style={{ fontSize: 16, lineHeight: 1 }}>☰</span><span style={{ fontSize: 14 }}>Tools</span></>
              : (moreIcon ?? '⋯')}
          </button>
        </div>
      )}

      {open && overflow && typeof document !== 'undefined' && createPortal(isMobile ? sheet : dropdown, document.body)}
      {modalEl && typeof document !== 'undefined' && createPortal(modalEl, document.body)}
    </div>
  );
}
