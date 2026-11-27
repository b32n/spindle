// TextOptionsButton — "Text box options" (autofit, insets, direction, columns) for the selected or edited text box.
// A separate toolbar item next to the text format bar: an anchored popover on the toolbar, a row plus modal inside the
// overflow menu (an anchored popover would paint under that surface), like the image crop control.

import React, { useState } from 'react';
import { Settings2 } from 'lucide-react';
import { autofitMode, mergeBodyStyle, type BodyStyle } from '@b32nio/spindle-slides-core';
import { useToolbarMenu, MenuRow } from '@b32nio/spindle-shared/react';
import { useDeck, useSelection, useEditingId, useElement } from '../hooks';
import { ToolbarButton } from './toolbarUI';
import { Popover } from './Popover';
import { TextBodyFields } from './TextBodyFields';
import { refitElement } from '../text-fit';

const LABEL = 'Text box options';

/** The text body fields bound live to one element (the overflow modal's content is a static node, so it reads the model itself). */
export function LiveTextBodyFields({ id }: { id: string }): React.ReactElement | null {
  const deck = useDeck();
  const el = useElement(id);
  if (!el || (el.type !== 'text' && el.type !== 'shape')) return null;
  const apply = (patch: Partial<BodyStyle>) => {
    deck.updateElement(id, { bodyStyle: mergeBodyStyle(el.bodyStyle, patch) });
    // A fit depends on the text as laid out with the new settings: measure once the new layout is on screen.
    if (['autofit', 'insets', 'vert', 'columns', 'anchorCtr'].some((k) => k in patch)) {
      requestAnimationFrame(() => refitElement(deck, id));
    }
  };
  return <TextBodyFields idPrefix={`text-body-${id}`} body={el.bodyStyle} onChange={apply} allowResize={el.type === 'text'} />;
}

export function TextOptionsButton(): React.ReactElement | null {
  const selection = useSelection();
  const editingId = useEditingId();
  const menu = useToolbarMenu();
  const [anchor, setAnchor] = useState<HTMLElement | null>(null);
  // The element the popover is open for, so it closes when the selection moves to another box.
  const [openFor, setOpenFor] = useState<string | null>(null);
  const id = editingId ?? (selection.elementIds.length === 1 ? selection.elementIds[0] : null);
  const el = useElement(id ?? '');
  if (!id || !el || (el.type !== 'text' && el.type !== 'shape')) return null;
  // A selected shape shows its own bar; its text options appear once its text is edited (as the text format bar does).
  if (el.type === 'shape' && editingId !== id) return null;
  const open = openFor === id;

  if (menu.inMenu) {
    return (
      <MenuRow
        icon={<Settings2 size={17} />}
        label={LABEL}
        hint={autofitMode(el.bodyStyle) === 'none' ? undefined : autofitMode(el.bodyStyle) === 'shrink' ? 'Shrink' : 'Resize'}
        onClick={() => menu.openModal(LABEL, <LiveTextBodyFields id={id} />)}
      />
    );
  }
  return (
    <>
      <span ref={setAnchor} style={{ display: 'inline-flex' }}>
        <ToolbarButton title={LABEL} active={open} onClick={() => setOpenFor(open ? null : id)}>
          <Settings2 size={15} />
        </ToolbarButton>
      </span>
      {open && (
        <Popover anchor={anchor} onClose={() => setOpenFor(null)}>
          <LiveTextBodyFields id={id} />
        </Popover>
      )}
    </>
  );
}
