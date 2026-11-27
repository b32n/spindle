// ContextMenu — the right-click menu for elements on the stage. Actions operate
// on the current selection. Rendering + dismissal live in the shared Menu.

import React from 'react';
import type { ConnectorKind } from '@b32nio/spindle-slides-core';
import { useDeck, useSelection, useClipboard } from '../hooks';
import { Menu, type MenuItem } from './Menu';

export interface ContextMenuProps {
  x: number;
  y: number;
  onClose: () => void;
}

export function ContextMenu({ x, y, onClose }: ContextMenuProps): React.ReactElement {
  const deck = useDeck();
  const selection = useSelection();
  const clipboard = useClipboard();
  const ids = selection.elementIds;
  const hasSel = ids.length > 0;
  const only = ids.length === 1 ? deck.getElement(ids[0]) : undefined;
  const line = only?.type === 'line' ? only : undefined;
  // Elbow and curved routing lives here (not in the toolbar) so the toolbar row is unchanged.
  const route = (kind: ConnectorKind, label: string): MenuItem => ({
    label: (line?.connector ?? 'straight') === kind ? `${label} ✓` : label,
    run: () => deck.setConnectorKind(line!.id, kind),
  });

  const items: Array<MenuItem | 'sep'> = [
    { label: 'Cut', run: () => clipboard.cut(), disabled: !hasSel },
    { label: 'Copy', run: () => clipboard.copy(), disabled: !hasSel },
    { label: 'Paste', run: () => clipboard.paste(), disabled: !clipboard.hasContent() },
    { label: 'Duplicate', run: () => { const c = ids.map((id) => deck.duplicateElement(id)).filter(Boolean); if (c.length) deck.setSelection({ slideId: deck.getActiveSlideId(), elementIds: c.map((x) => x!.id) }); }, disabled: !hasSel },
    'sep',
    { label: 'Bring to front', run: () => deck.bringToFront(ids), disabled: !hasSel },
    { label: 'Bring forward', run: () => deck.bringForward(ids), disabled: !hasSel },
    { label: 'Send backward', run: () => deck.sendBackward(ids), disabled: !hasSel },
    { label: 'Send to back', run: () => deck.sendToBack(ids), disabled: !hasSel },
    ...(line ? ['sep' as const, route('straight', 'Straight connector'), route('elbow', 'Elbow connector'), route('curved', 'Curved connector')] : []),
    'sep',
    { label: 'Group', run: () => deck.groupElements(ids), disabled: ids.length < 2 },
    { label: 'Ungroup', run: () => deck.ungroupElements(ids), disabled: !hasSel },
    'sep',
    { label: 'Delete', run: () => { deck.deleteElements(ids); deck.setSelection({ slideId: deck.getActiveSlideId(), elementIds: [] }); }, disabled: !hasSel },
  ];

  return <Menu x={x} y={y} items={items} onClose={onClose} />;
}
