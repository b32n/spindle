// Clipboard for elements. Uses an in-memory buffer (works even where the async
// Clipboard API is unavailable) and best-effort mirrors a marker onto the
// system clipboard as text/plain. Copy stores type-specific specs; paste
// re-adds them to the active slide with a small offset and selects them.

import { useCallback } from 'react';
import type { DeckImpl, NewElementSpec, SlideElement } from '@b32nio/spindle-slides-core';
import { useDeck } from './index';

// Module-scoped so copy in one place and paste in another share it.
let buffer: NewElementSpec[] = [];
// Parallel to `buffer`: each spec's source group id (undefined if ungrouped), so
// paste can re-group the copies that belonged together into a fresh group.
let bufferGroups: (string | undefined)[] = [];
let pasteCount = 0;

function toSpec(el: SlideElement): NewElementSpec {
  // Strip engine-assigned fields; the rest is a valid NewElementSpec.
  const clone = structuredClone(el) as unknown as Record<string, unknown>;
  delete clone.id;
  delete clone.containerId;
  delete clone.index;
  delete clone.groupId;
  return clone as unknown as NewElementSpec;
}

export interface ClipboardApi {
  copy(): void;
  cut(): void;
  paste(): void;
  hasContent(): boolean;
}

export function copyElements(deck: DeckImpl): void {
  const ids = deck.getSelection().elementIds;
  const els = ids.map((id) => deck.getElement(id)).filter((e): e is SlideElement => !!e);
  if (els.length === 0) return;
  buffer = els.map(toSpec);
  bufferGroups = els.map((el) => el.groupId);
  pasteCount = 0;
  try {
    void navigator.clipboard?.writeText('application/x-spindle-slides');
  } catch {
    /* clipboard unavailable — the in-memory buffer still works */
  }
}

/** Whether the in-memory element buffer holds anything to paste. */
export function hasClipboardContent(): boolean {
  return buffer.length > 0;
}

export function pasteElements(deck: DeckImpl): void {
  if (buffer.length === 0) return;
  const slideId = deck.getActiveSlideId();
  pasteCount += 1;
  const offset = 16 * pasteCount;
  const newIds: string[] = [];
  const groupToNewIds = new Map<string, string[]>();
  buffer.forEach((spec, i) => {
    const base = spec as unknown as { x?: number; y?: number };
    const el = deck.addElement(slideId, {
      ...spec,
      x: (base.x ?? 0) + offset,
      y: (base.y ?? 0) + offset,
    } as NewElementSpec);
    newIds.push(el.id);
    const g = bufferGroups[i];
    if (g) {
      const arr = groupToNewIds.get(g) ?? [];
      arr.push(el.id);
      groupToNewIds.set(g, arr);
    }
  });
  // Re-form each source group among the pasted copies (fresh group ids).
  for (const ids of groupToNewIds.values()) {
    if (ids.length > 1) deck.groupElements(ids);
  }
  deck.setSelection({ slideId, elementIds: newIds });
}

export function useClipboard(): ClipboardApi {
  const deck = useDeck();
  const copy = useCallback(() => copyElements(deck), [deck]);
  const cut = useCallback(() => {
    copyElements(deck);
    deck.deleteElements(deck.getSelection().elementIds);
    deck.setSelection({ slideId: deck.getActiveSlideId(), elementIds: [] });
  }, [deck]);
  const paste = useCallback(() => pasteElements(deck), [deck]);
  const hasContent = useCallback(() => buffer.length > 0, []);
  return { copy, cut, paste, hasContent };
}
