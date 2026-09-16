// Deck layer types — slides, the serializable DeckData shape, comment types,
// and the semantic event union.

import type { CommentThread } from '@b32nio/spindle-shared';
import type { Fill, SlideElement } from './scene/types';
import type { RichTextDoc } from './text/model';
import type { ThemeData, LayoutData } from './theme/types';

/** A slide's own record (its elements are stored flat on the engine). */
export interface Slide {
  id: string;
  /** Fractional order key — slides are sorted ascending by this. */
  index: string;
  /** Id of the layout this slide was created from (for placeholder styling). */
  layoutRef?: string;
  /** Slide background fill; defaults to the theme light background. */
  background?: Fill;
  /** Speaker notes. */
  notes?: RichTextDoc;
}

/** A slide plus its elements — the nested form used by getData/setData. */
export interface SlideData extends Slide {
  elements: SlideElement[];
}

/** Per-user view state. Local only — never serialized into the CRDT doc. */
export interface DeckSelection {
  slideId?: string;
  elementIds: string[];
}

// ── Comments (store implemented in Phase 6; types live here for DeckData) ─────

/** Anchors a comment thread to an element on a slide. */
export interface ElementCommentAnchor {
  slideId: string;
  elementId: string;
}

export type SlidesCommentThread = CommentThread<ElementCommentAnchor>;

// ── Serialization ─────────────────────────────────────────────────────────────

export interface DeckData {
  id: string;
  title: string;
  slideSize: { w: number; h: number };
  theme: ThemeData;
  /** Deck-scoped layout definitions; defaults to the built-ins when omitted. */
  layouts?: LayoutData[];
  slides: SlideData[];
  threads?: SlidesCommentThread[];
  /** Optional local selection snapshot (round-trips but is not collab state). */
  selection?: DeckSelection;
}

// ── Events ─────────────────────────────────────────────────────────────────────

export type DeckEventType =
  | 'deckChange'
  | 'themeChange'
  | 'slideAdd'
  | 'slideDelete'
  | 'slideMove'
  | 'slideChange'
  | 'elementAdd'
  | 'elementDelete'
  | 'elementChange'
  | 'selectionChange'
  | 'activeSlideChange'
  | 'commentChange'
  | 'commentEvent';

export interface SlideAddPayload {
  slideId: string;
}
export interface SlideDeletePayload {
  slideId: string;
}
export interface SlideMovePayload {
  slideId: string;
}
export interface SlideChangePayload {
  slideId: string;
  keys?: string[];
}
export interface ElementAddPayload {
  slideId: string;
  elementId: string;
}
export interface ElementDeletePayload {
  slideId: string;
  elementId: string;
}
export interface ElementChangePayload {
  slideId: string;
  elementId: string;
  /** Which fields changed, when known (drives fine-grained React updates). */
  keys?: string[];
}
export interface SelectionChangePayload {
  selection: DeckSelection;
}
export interface ActiveSlideChangePayload {
  slideId: string;
}
