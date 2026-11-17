// A dropdown panel rendered in a portal to <body>, anchored under a button.
// Portalling escapes the toolbar's overflow clipping and stacking context, so
// menus always paint above the editor content.

import React, { useLayoutEffect, useState } from 'react';
import { createPortal } from 'react-dom';

export function Popover({
  anchor,
  onClose,
  children,
}: {
  anchor: HTMLElement | null;
  onClose: () => void;
  children: React.ReactNode;
}): React.ReactElement | null {
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null);

  // Reading the anchor's rect is a layout measurement; it must land before paint
  // (no flash at 0,0), so the sync setState is intentional. Runs only on anchor change.
  useLayoutEffect(() => {
    if (!anchor) return;
    const r = anchor.getBoundingClientRect();
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setPos({ left: r.left, top: r.bottom + 4 });
  }, [anchor]);

  if (!pos) return null;
  return createPortal(
    <>
      <div onPointerDown={onClose} style={{ position: 'fixed', inset: 0, zIndex: 2000 }} />
      <div
        style={{
          position: 'fixed',
          left: pos.left,
          top: pos.top,
          zIndex: 2001,
          background: '#fff',
          border: '1px solid #d5d9e0',
          borderRadius: 8,
          boxShadow: '0 8px 28px rgba(0,0,0,0.16)',
          padding: 6,
        }}
      >
        {children}
      </div>
    </>,
    document.body
  );
}
