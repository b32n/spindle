// SelectionOverlay draws the selection box, resize handles, and rotate handle
// on top of the slide. It lives inside the scaled slide container (slide
// coordinates), so it rotates with a single element naturally; handle sizes are
// counter-scaled by 1/scale to stay a constant screen size at any zoom.
//
// Handles carry data-handle / data-rotate attributes; the interactive stage
// reads them on pointerdown to start the right gesture.

import React, { useEffect, useReducer, useSyncExternalStore } from 'react';
import { unionAABB, resolveEndpoints, type Frame, type LineElement, type Point, type Rect } from '@b32nio/spindle-slides-core';
import { useDeckContext } from '../context/DeckContext';
import { useSelection } from '../hooks';

const ACCENT = '#2d7ff9';
const CORNER_CURSOR: Record<string, string> = {
  nw: 'nwse-resize', se: 'nwse-resize', ne: 'nesw-resize', sw: 'nesw-resize',
  n: 'ns-resize', s: 'ns-resize', e: 'ew-resize', w: 'ew-resize',
};

// A circular-arrow rotate cursor (white body + dark outline so it reads on any
// background), hotspot centred.
const ROTATE_CURSOR =
  'url("data:image/svg+xml,' +
  encodeURIComponent(
    '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke-linecap="round" stroke-linejoin="round">' +
      '<path d="M5 12a7 7 0 1 1 2 4.9" stroke="#fff" stroke-width="4.5"/><path d="M5 17.5v-4h4" stroke="#fff" stroke-width="4.5"/>' +
      '<path d="M5 12a7 7 0 1 1 2 4.9" stroke="#111" stroke-width="2"/><path d="M5 17.5v-4h4" stroke="#111" stroke-width="2"/>' +
      '</svg>'
  ) +
  '") 12 12, grab';

function frameOf(el: { x: number; y: number; w: number; h: number; rotation: number }): Frame {
  return { x: el.x, y: el.y, w: el.w, h: el.h, rotation: el.rotation };
}

function Handle({ name, x, y, hs }: { name: string; x: number; y: number; hs: number }): React.ReactElement {
  return (
    <div
      data-handle={name}
      style={{
        position: 'absolute',
        left: x - hs / 2,
        top: y - hs / 2,
        width: hs,
        height: hs,
        background: '#fff',
        border: `${hs * 0.14}px solid ${ACCENT}`,
        borderRadius: hs * 0.2,
        boxSizing: 'border-box',
        cursor: CORNER_CURSOR[name] ?? 'pointer',
        pointerEvents: 'auto',
      }}
    />
  );
}

function RotateHandle({ cx, top, hs }: { cx: number; top: number; hs: number }): React.ReactElement {
  return (
    <div
      data-rotate="true"
      style={{
        position: 'absolute',
        left: cx - hs / 2,
        top: top - hs * 2,
        width: hs,
        height: hs,
        background: '#fff',
        border: `${hs * 0.14}px solid ${ACCENT}`,
        borderRadius: '50%',
        boxSizing: 'border-box',
        cursor: ROTATE_CURSOR,
        pointerEvents: 'auto',
      }}
    />
  );
}

/** Thin grab strips along a selected table's four edges. Its body is busy
 *  selecting cells, so dragging the border is how you move the whole table. */
function TableMoveBands({ w, h, bt }: { w: number; h: number; bt: number }): React.ReactElement {
  const band: React.CSSProperties = { position: 'absolute', background: 'transparent', cursor: 'move', pointerEvents: 'auto' };
  return (
    <>
      <div data-table-move="1" style={{ ...band, left: 0, top: -bt / 2, width: w, height: bt }} />
      <div data-table-move="1" style={{ ...band, left: 0, top: h - bt / 2, width: w, height: bt }} />
      <div data-table-move="1" style={{ ...band, left: -bt / 2, top: 0, width: bt, height: h }} />
      <div data-table-move="1" style={{ ...band, left: w - bt / 2, top: 0, width: bt, height: h }} />
    </>
  );
}

/** A round handle at a line tip; drag it to resize+rotate the line at once. */
function EndpointHandle({ end, x, y, hs }: { end: 'start' | 'end'; x: number; y: number; hs: number }): React.ReactElement {
  return (
    <div
      data-endpoint={end}
      style={{
        position: 'absolute',
        left: x - hs / 2,
        top: y - hs / 2,
        width: hs,
        height: hs,
        background: '#fff',
        border: `${hs * 0.16}px solid ${ACCENT}`,
        borderRadius: '50%',
        boxSizing: 'border-box',
        cursor: 'crosshair',
        pointerEvents: 'auto',
      }}
    />
  );
}

export function SelectionOverlay({ scale }: { scale: number }): React.ReactElement | null {
  const { deck, transient, connectors } = useDeckContext();
  const selection = useSelection();
  const [, force] = useReducer((n) => n + 1, 0);

  // Re-render when a selected element or the deck changes (post-commit frames).
  useEffect(() => {
    const offs = [deck.on('elementChange', force as () => void), deck.on('deckChange', force as () => void)];
    return () => offs.forEach((o) => o());
  }, [deck]);

  // Subscribe to the transient store so the box follows the element live during
  // a gesture (move/resize/rotate write frames here per pointermove).
  const transientState = useSyncExternalStore(transient.subscribe, transient.get);
  const live = transientState.liveFrames;
  const edit = useSyncExternalStore(connectors.subscribe, connectors.get).edit;

  const ids = selection.elementIds.filter((id) => deck.getElement(id));
  if (ids.length === 0) return null;

  const hs = 10 / scale;
  const border = 1.5 / scale;

  const getFrame = (id: string): Frame => live?.get(id) ?? frameOf(deck.getElement(id)!);

  // A single selected line shows tip handles instead of a box — a line is two
  // endpoints, and dragging a tip resizes and rotates it together.
  if (ids.length === 1) {
    const only = deck.getElement(ids[0])!;
    if (only.type === 'line') {
      const line = only as LineElement;
      const getEndpointFrame = (elId: string): Frame | undefined => {
        const l = live?.get(elId);
        if (l) return l;
        const e = deck.getElement(elId);
        return e ? frameOf(e) : undefined;
      };
      const resolved = resolveEndpoints(line, getEndpointFrame);
      const editing = edit?.elementId === line.id ? edit : null;
      const start: Point = editing?.end === 'start' ? editing.point : resolved.start;
      const end: Point = editing?.end === 'end' ? editing.point : resolved.end;
      const eh = hs * 1.1;
      return (
        <div style={{ position: 'absolute', left: 0, top: 0, pointerEvents: 'none' }}>
          <EndpointHandle end="start" x={start.x} y={start.y} hs={eh} />
          <EndpointHandle end="end" x={end.x} y={end.y} hs={eh} />
        </div>
      );
    }
    const f = getFrame(ids[0]);
    return (
      <div style={{ position: 'absolute', left: 0, top: 0, pointerEvents: 'none' }}>
        <div
          style={{
            position: 'absolute',
            left: 0,
            top: 0,
            width: f.w,
            height: f.h,
            transform: `translate(${f.x}px, ${f.y}px) rotate(${f.rotation}deg)`,
            transformOrigin: 'center center',
          }}
        >
          <div style={{ position: 'absolute', inset: 0, border: `${border}px solid ${ACCENT}`, boxSizing: 'border-box' }} />
          {only.type === 'table' && <TableMoveBands w={f.w} h={f.h} bt={hs} />}
          <RotateHandle cx={f.w / 2} top={0} hs={hs} />
          {(['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const).map((name) => {
            const px = name.includes('w') ? 0 : name.includes('e') ? f.w : f.w / 2;
            const py = name.includes('n') ? 0 : name.includes('s') ? f.h : f.h / 2;
            return <Handle key={name} name={name} x={px} y={py} hs={hs} />;
          })}
        </div>
      </div>
    );
  }

  // Multi-select: axis-aligned union box with corner handles + rotate.
  const bounds: Rect = unionAABB(ids.map(getFrame)) ?? { x: 0, y: 0, w: 0, h: 0 };
  return (
    <div style={{ position: 'absolute', left: 0, top: 0, pointerEvents: 'none' }}>
      {/* Light outline on each member. */}
      {ids.map((id) => {
        const f = getFrame(id);
        return (
          <div
            key={id}
            style={{
              position: 'absolute',
              left: 0,
              top: 0,
              width: f.w,
              height: f.h,
              transform: `translate(${f.x}px, ${f.y}px) rotate(${f.rotation}deg)`,
              transformOrigin: 'center center',
              border: `${border}px solid rgba(45,127,249,0.5)`,
              boxSizing: 'border-box',
            }}
          />
        );
      })}
      <div style={{ position: 'absolute', left: bounds.x, top: bounds.y, width: bounds.w, height: bounds.h }}>
        <div style={{ position: 'absolute', inset: 0, border: `${border}px solid ${ACCENT}`, boxSizing: 'border-box' }} />
        <RotateHandle cx={bounds.w / 2} top={0} hs={hs} />
        {(['nw', 'ne', 'se', 'sw'] as const).map((name) => {
          const px = name.includes('w') ? 0 : bounds.w;
          const py = name.includes('n') ? 0 : bounds.h;
          return <Handle key={name} name={name} x={px} y={py} hs={hs} />;
        })}
      </div>
    </div>
  );
}
