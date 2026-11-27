// SlideContextMenu — the right-click menu for a slide thumbnail in the
// filmstrip. Add / duplicate / reorder / delete operate on that slide.

import React from 'react';
import { useDeck, useSlideIds } from '../hooks';
import { Menu, type MenuItem } from './Menu';

export interface SlideContextMenuProps {
  slideId: string;
  x: number;
  y: number;
  onClose: () => void;
}

export function SlideContextMenu({ slideId, x, y, onClose }: SlideContextMenuProps): React.ReactElement {
  const deck = useDeck();
  const slideIds = useSlideIds();
  const i = slideIds.indexOf(slideId);
  const isLast = slideIds.length <= 1;
  const hidden = !!deck.getSlide(slideId)?.hidden;

  const deleteSlide = () => {
    if (isLast) return;
    const neighbor = slideIds[i + 1] ?? slideIds[i - 1];
    deck.deleteSlide(slideId);
    if (neighbor) deck.setActiveSlide(neighbor);
  };

  const items: Array<MenuItem | 'sep'> = [
    { label: 'New slide', run: () => { const s = deck.addSlide({ afterSlideId: slideId, layoutId: 'titleContent' }); deck.setActiveSlide(s.id); } },
    { label: 'Duplicate slide', run: () => { const s = deck.duplicateSlide(slideId); if (s) deck.setActiveSlide(s.id); } },
    'sep',
    { label: 'Move up', run: () => { const prev = slideIds[i - 2]; deck.moveSlide(slideId, prev ? { afterSlideId: prev } : {}); }, disabled: i <= 0 },
    { label: 'Move down', run: () => { const nextAfter = slideIds[i + 1]; if (nextAfter) deck.moveSlide(slideId, { afterSlideId: nextAfter }); }, disabled: i >= slideIds.length - 1 },
    'sep',
    { label: hidden ? 'Show slide' : 'Hide slide', run: () => deck.setSlideHidden(slideId, !hidden) },
    'sep',
    { label: 'Delete slide', run: deleteSlide, disabled: isLast },
  ];

  return <Menu x={x} y={y} items={items} onClose={onClose} />;
}
