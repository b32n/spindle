// Toolbar — insert elements, edit, arrange, and align. Enabled state follows
// the current selection. Insertions drop near the slide centre and select the
// new element.

import React, { useRef, useSyncExternalStore } from 'react';
import {
  Type, Minus, MoveUpRight, Image as ImageIcon, Link2, Table,
  Trash2, Copy, Undo2, Redo2, Group, Ungroup, GalleryVerticalEnd,
  BringToFront, SendToBack, AlignHorizontalJustifyCenter, AlignVerticalJustifyCenter,
  AlignStartVertical, AlignEndVertical, AlignStartHorizontal, AlignEndHorizontal,
  AlignHorizontalDistributeCenter, AlignVerticalDistributeCenter,
} from 'lucide-react';
import type { NewElementSpec, AlignMode } from '@b32nio/spindle-slides-core';
import { ResponsiveToolbar } from '@b32nio/spindle-shared/react';
import { useDeck, useSelection, useEditingId, useFilmstripOpen } from '../hooks';
import { useDeckContext } from '../context/DeckContext';
import { DeckControls } from './DeckControls';
import { TextFormatBar } from './TextFormatBar';
import { LineFormatBar } from './LineFormatBar';
import { ShapeFormatBar } from './ShapeFormatBar';
import { ImageFormatBar } from './ImageFormatBar';
import { TableFormatBar } from './TableFormatBar';
import { ShapePicker } from './ShapePicker';
import { TB, ToolbarButton as IconButton, ToolbarDivider } from './toolbarUI';

export function Toolbar({ extras }: { extras?: React.ReactNode } = {}): React.ReactElement {
  const deck = useDeck();
  const selection = useSelection();
  const { tableSel, ui } = useDeckContext();
  const cellSel = useSyncExternalStore(tableSel.subscribe, tableSel.getState);
  const filmstripOpen = useFilmstripOpen();
  const fileRef = useRef<HTMLInputElement>(null);
  const editingId = useEditingId();
  const ids = selection.elementIds;
  const hasSel = ids.length > 0;
  const hasMulti = ids.length > 1;
  const canDistribute = ids.length >= 3;
  const { w, h } = deck.getSlideSize();

  // Contextual toolbar: while formatting an element the insert/arrange groups
  // step aside so only the relevant controls show (one row, less clutter);
  // they return when nothing is selected. Text is the space-hungry case, so
  // arrange also hides for a selected text element / while editing.
  const editing = editingId != null;
  const single = ids.length === 1 ? deck.getElement(ids[0]) : null;
  const isTextSingle = single?.type === 'text';
  const isTableSingle = single?.type === 'table';
  // A live cell-range selection turns the toolbar into a cell-formatting bar
  // (text + table controls), so the object-level actions step aside for room.
  const cellsActive = isTableSingle && !!cellSel && cellSel.tableId === single!.id;
  const showInsert = !hasSel && !editing;
  const showActions = hasSel && !editing && !cellsActive; // duplicate / delete
  // Arrange hides for text and tables (both bring space-hungry format bars).
  const showArrange = hasSel && !editing && !isTextSingle && !isTableSingle; // z-order / group / align

  const insert = (spec: NewElementSpec, size: { w: number; h: number }) => {
    const slideId = deck.getActiveSlideId();
    const el = deck.addElement(slideId, { ...spec, x: (w - size.w) / 2, y: (h - size.h) / 2 } as NewElementSpec);
    deck.setSelection({ slideId, elementIds: [el.id] });
  };

  const insertImage = (src: string) => {
    const img = new Image();
    img.onload = () => {
      const maxW = w * 0.6;
      const scale = Math.min(1, maxW / img.naturalWidth);
      insert(
        { type: 'image', src, naturalW: img.naturalWidth, naturalH: img.naturalHeight } as NewElementSpec,
        { w: img.naturalWidth * scale, h: img.naturalHeight * scale }
      );
    };
    img.src = src;
  };

  const onImageFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => insertImage(reader.result as string);
    reader.readAsDataURL(file);
    e.target.value = '';
  };

  const onImageUrl = () => {
    const url = window.prompt('Image URL', '');
    if (url && url.trim()) insertImage(url.trim());
  };

  const align = (mode: AlignMode) => {
    if (ids.length >= 2) deck.alignElements(ids, mode);
    else if (ids.length === 1) deck.alignElements(ids, mode, { x: 0, y: 0, w, h });
  };

  return (
    <div style={TB.strip}>
      <div style={TB.pill}>
      <input ref={fileRef} type="file" accept="image/*" onChange={onImageFile} style={{ display: 'none' }} />
      {/* Pinned outside ResponsiveToolbar so it never collapses into the overflow
          menu — the filmstrip toggle stays a first-class, always-visible button. */}
      <IconButton title="Slides panel" label="Slides" active={filmstripOpen} onClick={() => ui.toggleFilmstrip()}>
        <GalleryVerticalEnd size={16} />
      </IconButton>
      <ToolbarDivider />
      <ResponsiveToolbar gap={2}>
      <DeckControls />

      {/* Insert — only when nothing is selected. */}
      {showInsert && (
        <>
          <ToolbarDivider />
          <IconButton title="Text box" onClick={() => insert({ type: 'text' } as NewElementSpec, { w: 400, h: 100 })}>
            <Type size={16} />
          </IconButton>
          <ShapePicker />
          <IconButton title="Line" onClick={() => insert({ type: 'line' } as NewElementSpec, { w: 300, h: 0 })}>
            <Minus size={16} />
          </IconButton>
          <IconButton title="Arrow" onClick={() => insert({ type: 'line', endArrow: 'triangle' } as NewElementSpec, { w: 300, h: 0 })}>
            <MoveUpRight size={16} />
          </IconButton>
          <IconButton title="Image (upload)" onClick={() => fileRef.current?.click()}>
            <ImageIcon size={16} />
          </IconButton>
          <IconButton title="Image from URL" onClick={onImageUrl}>
            <Link2 size={16} />
          </IconButton>
          <IconButton title="Table (3×3)" onClick={() => insert({ type: 'table', rows: 3, cols: 3 } as NewElementSpec, { w: 720, h: 300 })}>
            <Table size={16} />
          </IconButton>
        </>
      )}

      {/* Duplicate / delete — for a selected element (not while editing text). */}
      {showActions && (
        <>
          <ToolbarDivider />
          <IconButton title="Duplicate (⌘D)" onClick={() => { const c = ids.map((id) => deck.duplicateElement(id)).filter(Boolean); if (c.length) deck.setSelection({ slideId: deck.getActiveSlideId(), elementIds: c.map((x) => x!.id) }); }}>
            <Copy size={16} />
          </IconButton>
          <IconButton title="Delete (⌦)" onClick={() => { deck.deleteElements(ids); deck.setSelection({ slideId: deck.getActiveSlideId(), elementIds: [] }); }}>
            <Trash2 size={16} />
          </IconButton>
        </>
      )}

      {/* Arrange — hidden for text (space) and while editing. */}
      {showArrange && (
        <>
          <ToolbarDivider />
          <IconButton title="Bring to front (⌘⇧])" onClick={() => deck.bringToFront(ids)}>
            <BringToFront size={16} />
          </IconButton>
          <IconButton title="Send to back (⌘⇧[)" onClick={() => deck.sendToBack(ids)}>
            <SendToBack size={16} />
          </IconButton>
          <IconButton title="Group (⌘G)" disabled={!hasMulti} onClick={() => deck.groupElements(ids)}>
            <Group size={16} />
          </IconButton>
          <IconButton title="Ungroup (⌘⇧G)" onClick={() => deck.ungroupElements(ids)}>
            <Ungroup size={16} />
          </IconButton>
          <ToolbarDivider />
          <IconButton title="Align left" onClick={() => align('left')}>
            <AlignStartVertical size={16} />
          </IconButton>
          <IconButton title="Align centre" onClick={() => align('centerH')}>
            <AlignHorizontalJustifyCenter size={16} />
          </IconButton>
          <IconButton title="Align right" onClick={() => align('right')}>
            <AlignEndVertical size={16} />
          </IconButton>
          <IconButton title="Align top" onClick={() => align('top')}>
            <AlignStartHorizontal size={16} />
          </IconButton>
          <IconButton title="Align middle" onClick={() => align('centerV')}>
            <AlignVerticalJustifyCenter size={16} />
          </IconButton>
          <IconButton title="Align bottom" onClick={() => align('bottom')}>
            <AlignEndHorizontal size={16} />
          </IconButton>
          <IconButton title="Distribute horizontally" disabled={!canDistribute} onClick={() => deck.distributeElements(ids, 'h')}>
            <AlignHorizontalDistributeCenter size={16} />
          </IconButton>
          <IconButton title="Distribute vertically" disabled={!canDistribute} onClick={() => deck.distributeElements(ids, 'v')}>
            <AlignVerticalDistributeCenter size={16} />
          </IconButton>
        </>
      )}

      <ToolbarDivider />
      <IconButton title="Undo (⌘Z)" onClick={() => deck.undo()}>
        <Undo2 size={16} />
      </IconButton>
      <IconButton title="Redo (⌘⇧Z)" onClick={() => deck.redo()}>
        <Redo2 size={16} />
      </IconButton>

      {/* Contextual format bars. Text formatting shows for a text element or a
          cell being edited; the structure bars hide while editing text. */}
      <TextFormatBar />
      {!editing && <ShapeFormatBar />}
      {!editing && <ImageFormatBar />}
      {!editing && <TableFormatBar />}
      {!editing && <LineFormatBar />}
      {/* Host-injected controls (e.g. app-specific actions). */}
      {extras}
      </ResponsiveToolbar>
      </div>
    </div>
  );
}
