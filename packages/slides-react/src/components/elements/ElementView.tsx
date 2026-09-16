// ElementView — the absolutely-positioned wrapper for one element. It
// subscribes to just that element (and the theme) via keyed hooks, so a single
// element edit re-renders this node alone. Rotation is a CSS transform about
// the box centre; the un-rotated box's top-left sits at (x, y).
//
// When interactive, the wrapper registers itself in the NodeRegistry so
// gestures can write transforms straight to the DOM, and it accepts pointer
// events (the stage delegates hit-testing via `closest('[data-element-id]')`).

import React, { useCallback } from 'react';
import { useElement, useTheme } from '../../hooks';
import { useDeckContext } from '../../context/DeckContext';
import type { SlideElement } from '@b32nio/spindle-slides-core';
import { TextView } from './TextView';
import { ShapeView } from './ShapeView';
import { ImageView } from './ImageView';
import { LineView } from './LineView';
import { ConnectorView } from './ConnectorView';
import { TableView } from './TableView';

function renderInner(el: SlideElement, theme: ReturnType<typeof useTheme>, interactive: boolean): React.ReactElement {
  switch (el.type) {
    case 'text':
      return <TextView el={el} theme={theme} interactive={interactive} />;
    case 'shape':
      return <ShapeView el={el} theme={theme} interactive={interactive} />;
    case 'image':
      return <ImageView el={el} />;
    case 'line':
      return <LineView el={el} theme={theme} />;
    case 'table':
      return <TableView el={el} theme={theme} interactive={interactive} />;
  }
}

export function ElementView({
  elementId,
  interactive = false,
}: {
  elementId: string;
  interactive?: boolean;
}): React.ReactElement | null {
  const el = useElement(elementId);
  const theme = useTheme();
  const { nodes } = useDeckContext();

  const ref = useCallback(
    (node: HTMLDivElement | null) => {
      if (!interactive) return;
      if (node) nodes.register(elementId, node);
      else nodes.unregister(elementId);
    },
    [nodes, elementId, interactive]
  );

  if (!el) return null;

  // A line with a binding or an explicit endpoint is a connector: its geometry
  // is derived (from linked shapes and/or stored endpoints), so a dedicated view
  // resolves + tracks it (subscribing to the transient store for live drags).
  // Ordinary elements never pay that cost.
  if (el.type === 'line' && (el.startBind || el.endBind || el.startPoint || el.endPoint)) {
    return <ConnectorView elementId={elementId} interactive={interactive} />;
  }

  const style: React.CSSProperties = {
    position: 'absolute',
    left: 0,
    top: 0,
    width: el.w,
    height: el.h,
    transform: `translate(${el.x}px, ${el.y}px) rotate(${el.rotation}deg)`,
    transformOrigin: 'center center',
    opacity: el.opacity ?? 1,
    pointerEvents: interactive && !el.locked ? 'auto' : 'none',
    cursor: interactive ? 'move' : 'default',
  };

  return (
    <div ref={ref} data-element-id={el.id} style={style}>
      {renderInner(el, theme, interactive)}
    </div>
  );
}
