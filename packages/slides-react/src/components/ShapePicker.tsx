// ShapePicker — a toolbar dropdown exposing every built-in shape preset (the
// toolbar previously offered only rectangle + ellipse). Each preset renders as
// a mini inline-SVG using the same geometry as the canvas; selecting one
// inserts it centred on the active slide.

import React, { useRef, useState } from 'react';
import { Shapes as ShapesIcon, ChevronDown } from 'lucide-react';
import type { ShapePreset, NewElementSpec } from '@b32nio/spindle-slides-core';
import { useToolbarMenu, MenuRow } from '@b32nio/spindle-shared/react';
import { shapeGeom } from './elements/shapes';
import { useDeck } from '../hooks';
import { Popover } from './Popover';
import { TB } from './toolbarUI';

const PRESETS: ShapePreset[] = [
  'rect', 'roundRect', 'ellipse', 'triangle', 'rightTriangle', 'diamond',
  'pentagon', 'hexagon', 'octagon', 'star5', 'arrowRight', 'arrowLeft',
  'chevron', 'parallelogram', 'trapezoid', 'plus', 'heart', 'cloud',
];

const btn: React.CSSProperties = { ...TB.dropdownButton, gap: 2, padding: '0 6px' };

function ShapeIcon({ preset }: { preset: ShapePreset }): React.ReactElement {
  const g = shapeGeom(preset, 22, 22);
  return (
    <svg width={22} height={22} viewBox="-1 -1 24 24" style={{ display: 'block' }}>
      {g.type === 'ellipse' ? (
        <ellipse cx={g.cx} cy={g.cy} rx={g.rx} ry={g.ry} fill="#5b6673" />
      ) : (
        <path d={g.d} fill="#5b6673" />
      )}
    </svg>
  );
}

export function ShapePicker(): React.ReactElement {
  const deck = useDeck();
  const menu = useToolbarMenu();
  const [open, setOpen] = useState(false);
  const anchorRef = useRef<HTMLButtonElement>(null);
  const { w, h } = deck.getSlideSize();

  const insert = (preset: ShapePreset) => {
    const size = { w: 220, h: 180 };
    const slideId = deck.getActiveSlideId();
    const el = deck.addElement(slideId, {
      type: 'shape',
      shape: preset,
      x: (w - size.w) / 2,
      y: (h - size.h) / 2,
    } as NewElementSpec);
    deck.setSelection({ slideId, elementIds: [el.id] });
    setOpen(false);
  };

  const grid = (cell: number, onPick: (p: ShapePreset) => void) => (
    <div style={{ display: 'grid', gridTemplateColumns: `repeat(auto-fill, ${cell}px)`, gap: 6, justifyContent: 'center' }}>
      {PRESETS.map((p) => (
        <button
          key={p}
          title={p}
          onClick={() => onPick(p)}
          style={{ width: cell, height: cell, display: 'flex', alignItems: 'center', justifyContent: 'center', border: '1px solid transparent', borderRadius: 6, background: '#fff', cursor: 'pointer' }}
          onMouseEnter={(e) => ((e.currentTarget as HTMLElement).style.background = '#f1f4f8')}
          onMouseLeave={(e) => ((e.currentTarget as HTMLElement).style.background = '#fff')}
        >
          <ShapeIcon preset={p} />
        </button>
      ))}
    </div>
  );

  // Overflow-menu mode: a row that opens the shape grid in a modal.
  if (menu.inMenu) {
    return (
      <MenuRow
        icon={<ShapesIcon size={17} />}
        label="Shape"
        onClick={() => menu.openModal('Insert shape', grid(44, (p) => { insert(p); menu.closeMenu(); }))}
      />
    );
  }

  return (
    <>
      <button ref={anchorRef} title="Shapes" style={btn} onClick={() => setOpen((o) => !o)}>
        <ShapesIcon size={16} />
        <ChevronDown size={12} />
      </button>
      {open && (
        <Popover anchor={anchorRef.current} onClose={() => setOpen(false)}>
          {grid(32, insert)}
        </Popover>
      )}
    </>
  );
}
