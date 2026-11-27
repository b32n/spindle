// ConnectorView — an ElementView specialised for a *bound* line (a connector).
// Its endpoints track other elements' anchors, so its box is derived, not
// stored: we resolve it from the targets' committed frames, overridden by any
// live frame in the transient store so the connector follows a shape mid-drag.
// Only connectors subscribe to the transient store, so an unrelated drag never
// re-renders ordinary elements.

import React, { useCallback } from 'react';
import { useSyncExternalStore } from 'react';
import { anchorDirection, connectionFrame, connectorBox, resolveEndpoints, type ConnectionFrame, type Direction, type LineElement } from '@b32nio/spindle-slides-core';
import { useElement, useTheme } from '../../hooks';
import { useDeckContext } from '../../context/DeckContext';
import { LineView } from './LineView';

export function ConnectorView({ elementId, interactive }: { elementId: string; interactive: boolean }): React.ReactElement | null {
  const el = useElement(elementId) as LineElement | undefined;
  const theme = useTheme();
  const { deck, transient, nodes, connectors } = useDeckContext();
  const live = useSyncExternalStore(transient.subscribe, transient.get).liveFrames;
  const edit = useSyncExternalStore(connectors.subscribe, connectors.get).edit;

  const ref = useCallback(
    (node: HTMLDivElement | null) => {
      if (!interactive) return;
      if (node) nodes.register(elementId, node);
      else nodes.unregister(elementId);
    },
    [nodes, elementId, interactive]
  );

  if (!el) return null;

  const getFrame = (id: string): ConnectionFrame | undefined => {
    const e = deck.getElement(id);
    const l = live?.get(id);
    return e ? { ...connectionFrame(e), ...l } : l;
  };
  const dirOf = (bind: LineElement['startBind']): Direction | undefined => {
    const f = bind && getFrame(bind.elementId);
    return f ? anchorDirection(f, bind!.anchor) : undefined;
  };

  const resolved = resolveEndpoints(el, getFrame);
  // While dragging one of this line's tips, override that end live.
  const start = edit?.elementId === el.id && edit.end === 'start' ? edit.point : resolved.start;
  const end = edit?.elementId === el.id && edit.end === 'end' ? edit.point : resolved.end;
  const box = connectorBox(start, end);
  const effective: LineElement = { ...el, x: box.x, y: box.y, w: box.w, h: box.h, flipH: box.flipH, flipV: box.flipV, rotation: 0 };
  // Local endpoints keep the arrowhead on the *bound* end regardless of where
  // the shapes sit relative to each other (the box can't encode direction).
  const endpoints = { x1: start.x - box.x, y1: start.y - box.y, x2: end.x - box.x, y2: end.y - box.y };

  const style: React.CSSProperties = {
    position: 'absolute',
    left: 0,
    top: 0,
    width: box.w,
    height: box.h,
    transform: `translate(${box.x}px, ${box.y}px)`,
    opacity: el.opacity ?? 1,
    pointerEvents: interactive && !el.locked ? 'auto' : 'none',
    cursor: interactive ? 'move' : 'default',
  };

  return (
    <div ref={ref} data-element-id={el.id} style={style}>
      <LineView el={effective} theme={theme} endpoints={endpoints} dirs={{ start: edit?.elementId === el.id && edit.end === 'start' ? undefined : dirOf(el.startBind), end: edit?.elementId === el.id && edit.end === 'end' ? undefined : dirOf(el.endBind) }} />
    </div>
  );
}
