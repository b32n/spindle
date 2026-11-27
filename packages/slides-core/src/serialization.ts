// setData validation + defaulting. getData() emits a fully-populated DeckData,
// but hand- or AI-authored JSON may omit ids, fractional indices, or the
// theme. normalizeDeckData fills those in so authoring a deck by hand stays
// easy: give it slides with elements in array order and it assigns everything
// order-sensitive.

import { generateId } from './utils/id';
import { indexesBetween, isValidIndex } from './scene/fractional-index';
import { getBuiltinTheme, DEFAULT_SLIDE_SIZE } from './theme/builtin';
import type { SlideElement } from './scene/types';
import type { DeckData, SlideData } from './types';

/**
 * Assign ascending fractional indices, preserving array order. The given indices
 * are trusted only when EVERY item has a well-formed, unique key — otherwise the
 * set can't be relied on for ordering, and a malformed key (missing, a bare '0',
 * out-of-alphabet, or a duplicate) would crash the next structural edit. In that
 * case we re-key the whole array in its current order.
 */
function ensureIndices<T extends { index?: string }>(items: T[]): Array<T & { index: string }> {
  const allValid = items.every((it) => isValidIndex(it.index));
  const unique = new Set(items.map((it) => it.index)).size === items.length;
  if (allValid && unique) return items as Array<T & { index: string }>;
  const keys = indexesBetween(null, null, items.length);
  return items.map((it, i) => ({ ...it, index: keys[i] }));
}

function normalizeSlide(slide: SlideData): SlideData {
  const id = slide.id || generateId();
  const rawElements = slide.elements ?? [];
  const withIndices = ensureIndices(
    rawElements.map((el) => ({ ...el, id: el.id || generateId(), containerId: id }))
  ) as SlideElement[];
  return { ...slide, id, elements: withIndices };
}

export function normalizeDeckData(data: DeckData): DeckData {
  const rawSlides = data.slides ?? [];
  const slides = ensureIndices(rawSlides.map(normalizeSlide)) as SlideData[];

  return {
    id: data.id ?? generateId(),
    title: data.title ?? 'Untitled deck',
    slideSize: data.slideSize ?? { ...DEFAULT_SLIDE_SIZE },
    theme: data.theme ?? getBuiltinTheme(),
    layouts: data.layouts,
    slides,
    threads: data.threads,
    opaqueParts: data.opaqueParts,
    selection: data.selection,
  };
}
