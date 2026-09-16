// Fine-grained subscription hooks. Each one subscribes to exactly the slice of
// engine state it needs, so re-renders stay local (a single element edit
// re-renders only that ElementView).

import { useCallback, useSyncExternalStore } from 'react';
import type {
  DeckImpl,
  Slide,
  SlideElement,
  ThemeData,
  DeckSelection,
} from '@b32nio/spindle-slides-core';
import { useDeckContext } from '../context/DeckContext';

export function useDeck(): DeckImpl {
  return useDeckContext().deck;
}

export function useSlideIds(): string[] {
  const { store } = useDeckContext();
  return useSyncExternalStore(store.subscribeSlideIds, store.getSlideIds);
}

export function useSlide(slideId: string): Slide | undefined {
  const { store } = useDeckContext();
  const subscribe = useCallback((cb: () => void) => store.subscribeSlide(slideId)(cb), [store, slideId]);
  const getSnapshot = useCallback(() => store.getSlide(slideId), [store, slideId]);
  return useSyncExternalStore(subscribe, getSnapshot);
}

export function useSlideElementIds(slideId: string): string[] {
  const { store } = useDeckContext();
  const subscribe = useCallback(
    (cb: () => void) => store.subscribeSlideElements(slideId)(cb),
    [store, slideId]
  );
  const getSnapshot = useCallback(() => store.getSlideElementIds(slideId), [store, slideId]);
  return useSyncExternalStore(subscribe, getSnapshot);
}

export function useElement(elementId: string): SlideElement | undefined {
  const { store } = useDeckContext();
  const subscribe = useCallback(
    (cb: () => void) => store.subscribeElement(elementId)(cb),
    [store, elementId]
  );
  const getSnapshot = useCallback(() => store.getElement(elementId), [store, elementId]);
  return useSyncExternalStore(subscribe, getSnapshot);
}

export function useSelection(): DeckSelection {
  const { store } = useDeckContext();
  return useSyncExternalStore(store.subscribeSelection, store.getSelection);
}

export function useActiveSlideId(): string {
  const { store } = useDeckContext();
  return useSyncExternalStore(store.subscribeActiveSlide, store.getActiveSlideId);
}

export function useTheme(): ThemeData {
  const { store } = useDeckContext();
  return useSyncExternalStore(store.subscribeTheme, store.getTheme);
}

export { useClipboard } from './useClipboard';
export { useKeyboardShortcuts } from './useKeyboardShortcuts';

/** The id of the element currently being text-edited, or null. */
export function useEditingId(): string | null {
  const { editing } = useDeckContext();
  return useSyncExternalStore(editing.subscribe, editing.getEditingId);
}

/** Whether the comments sidebar is open (shared UI state). */
export function useCommentsOpen(): boolean {
  const { ui } = useDeckContext();
  return useSyncExternalStore(ui.subscribe, ui.getCommentsOpen);
}

export function useFilmstripOpen(): boolean {
  const { ui } = useDeckContext();
  return useSyncExternalStore(ui.subscribe, ui.getFilmstripOpen);
}
