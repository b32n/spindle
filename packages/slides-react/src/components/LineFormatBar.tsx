// LineFormatBar — inline arrowhead controls, rendered *inside* the main Toolbar
// row (like TextFormatBar) so selecting a line doesn't shift the layout. Sets a
// selected line's start/end arrowheads to one of: no arrow, arrow on one side,
// or arrows on both ends. Renders nothing unless a single line is selected.

import React from 'react';
import { Minus, MoveLeft, MoveRight, MoveHorizontal } from 'lucide-react';
import type { ArrowHead, LineElement } from '@b32nio/spindle-slides-core';
import { useDeck, useSelection, useElement } from '../hooks';
import { ToolbarButton, ToolbarDivider } from './toolbarUI';

const HEAD: ArrowHead = 'triangle';

export function LineFormatBar(): React.ReactElement | null {
  const deck = useDeck();
  const selection = useSelection();
  const id = selection.elementIds.length === 1 ? selection.elementIds[0] : null;
  const el = useElement(id ?? '') as LineElement | undefined;
  if (!id || !el || el.type !== 'line') return null;

  const hasStart = (el.startArrow ?? 'none') !== 'none';
  const hasEnd = (el.endArrow ?? 'none') !== 'none';
  const mode: 'none' | 'start' | 'end' | 'both' =
    hasStart && hasEnd ? 'both' : hasStart ? 'start' : hasEnd ? 'end' : 'none';

  const set = (startArrow: ArrowHead, endArrow: ArrowHead) => deck.updateElement(id, { startArrow, endArrow });

  return (
    <>
      <ToolbarDivider />
      <ToolbarButton title="No arrow" active={mode === 'none'} onClick={() => set('none', 'none')}>
        <Minus size={15} />
      </ToolbarButton>
      <ToolbarButton title="Arrow (start)" active={mode === 'start'} onClick={() => set(HEAD, 'none')}>
        <MoveLeft size={15} />
      </ToolbarButton>
      <ToolbarButton title="Arrow (end)" active={mode === 'end'} onClick={() => set('none', HEAD)}>
        <MoveRight size={15} />
      </ToolbarButton>
      <ToolbarButton title="Arrows (both sides)" active={mode === 'both'} onClick={() => set(HEAD, HEAD)}>
        <MoveHorizontal size={15} />
      </ToolbarButton>
    </>
  );
}
