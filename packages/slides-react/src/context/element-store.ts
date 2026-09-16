// ElementStore — a keyed listener registry that turns the engine's semantic
// events into fine-grained useSyncExternalStore subscriptions. The point is
// that a single element change re-renders only that ElementView, not the whole
// deck.
//
// Snapshots must be referentially stable between renders unless the thing
// actually changed (useSyncExternalStore compares with Object.is, and an
// always-new snapshot loops forever). Engine records are immutable, so
// getElement/getSlide/getTheme return the engine's own object. Derived arrays
// (slide ids, per-slide element ids) and the selection object are cached here
// and only replaced when the relevant event fires.

import type {
  DeckImpl,
  Slide,
  SlideElement,
  ThemeData,
  DeckSelection,
  ElementChangePayload,
  ElementAddPayload,
  ElementDeletePayload,
  SlideChangePayload,
  SelectionChangePayload,
  ActiveSlideChangePayload,
} from '@b32nio/spindle-slides-core';

type Listener = () => void;

const keyForElement = (id: string) => `el:${id}`;
const keyForSlide = (id: string) => `slide:${id}`;
const keyForSlideElements = (slideId: string) => `els:${slideId}`;
const KEY_SLIDE_IDS = 'slideIds';
const KEY_SELECTION = 'selection';
const KEY_ACTIVE_SLIDE = 'activeSlide';
const KEY_THEME = 'theme';

export class ElementStore {
  private listeners = new Map<string, Set<Listener>>();

  private slideIdsCache: string[] = [];
  private elementIdsCache = new Map<string, string[]>();
  private selectionCache: DeckSelection;
  private activeSlideCache: string;

  private offFns: Array<() => void> = [];

  constructor(public readonly deck: DeckImpl) {
    this.slideIdsCache = deck.getSlideIds();
    this.selectionCache = deck.getSelection();
    this.activeSlideCache = deck.getActiveSlideId();
  }

  /**
   * Subscribe to the engine's events. Called from a DeckProvider effect (not
   * the constructor) so it survives React StrictMode's mount→unmount→mount:
   * connect → dispose → connect leaves the store subscribed. Idempotent.
   */
  connect(): void {
    if (this.offFns.length > 0) return; // already connected
    // Caches may have gone stale between construction and connect — refresh.
    this.slideIdsCache = this.deck.getSlideIds();
    this.selectionCache = this.deck.getSelection();
    this.activeSlideCache = this.deck.getActiveSlideId();
    this.wire();
  }

  private wire(): void {
    const { deck } = this;
    this.offFns.push(
      deck.on('elementChange', (e) => {
        const p = e.payload as ElementChangePayload;
        this.notify(keyForElement(p.elementId));
        // A z-order (index) change reorders paint order — refresh the id list.
        if (p.keys?.includes('index')) this.refreshElementIds(p.slideId);
      }),
      deck.on('elementAdd', (e) => {
        const p = e.payload as ElementAddPayload;
        this.refreshElementIds(p.slideId);
        this.notify(keyForElement(p.elementId));
      }),
      deck.on('elementDelete', (e) => {
        const p = e.payload as ElementDeletePayload;
        this.refreshElementIds(p.slideId);
        this.notify(keyForElement(p.elementId));
      }),
      deck.on('slideChange', (e) => {
        this.notify(keyForSlide((e.payload as SlideChangePayload).slideId));
      }),
      deck.on('slideAdd', () => this.refreshSlideIds()),
      deck.on('slideDelete', () => this.refreshSlideIds()),
      deck.on('slideMove', () => this.refreshSlideIds()),
      deck.on('themeChange', () => this.notify(KEY_THEME)),
      deck.on('selectionChange', (e) => {
        this.selectionCache = (e.payload as SelectionChangePayload).selection;
        this.notify(KEY_SELECTION);
      }),
      deck.on('activeSlideChange', (e) => {
        this.activeSlideCache = (e.payload as ActiveSlideChangePayload).slideId;
        this.notify(KEY_ACTIVE_SLIDE);
      }),
      // Broad invalidation (setData, undo/redo): rebuild caches, notify everyone.
      deck.on('deckChange', () => this.rebuildAndNotifyAll()),
    );
  }

  /** Unsubscribe from the engine. Component listeners (managed by the hooks'
   *  own useSyncExternalStore subscribe/unsubscribe) are left intact so a
   *  StrictMode reconnect keeps working. */
  dispose(): void {
    for (const off of this.offFns) off();
    this.offFns = [];
  }

  // ── Subscriptions ────────────────────────────────────────────────────────────

  private subscribe = (key: string, cb: Listener): (() => void) => {
    let set = this.listeners.get(key);
    if (!set) {
      set = new Set();
      this.listeners.set(key, set);
    }
    set.add(cb);
    return () => {
      set!.delete(cb);
      if (set!.size === 0) this.listeners.delete(key);
    };
  };

  private notify(key: string): void {
    const set = this.listeners.get(key);
    if (set) for (const cb of set) cb();
  }

  private notifyAll(): void {
    for (const set of this.listeners.values()) {
      for (const cb of set) cb();
    }
  }

  private refreshSlideIds(): void {
    this.slideIdsCache = this.deck.getSlideIds();
    this.notify(KEY_SLIDE_IDS);
  }

  private refreshElementIds(slideId: string): void {
    this.elementIdsCache.set(slideId, this.deck.getElementIdsForSlide(slideId));
    this.notify(keyForSlideElements(slideId));
  }

  private rebuildAndNotifyAll(): void {
    this.slideIdsCache = this.deck.getSlideIds();
    this.elementIdsCache.clear();
    this.selectionCache = this.deck.getSelection();
    this.activeSlideCache = this.deck.getActiveSlideId();
    this.notifyAll();
  }

  // ── Subscribe/snapshot pairs for hooks ───────────────────────────────────────

  subscribeSlideIds = (cb: Listener) => this.subscribe(KEY_SLIDE_IDS, cb);
  getSlideIds = (): string[] => this.slideIdsCache;

  subscribeSlide = (id: string) => (cb: Listener) => this.subscribe(keyForSlide(id), cb);
  getSlide = (id: string): Slide | undefined => this.deck.getSlide(id);

  subscribeSlideElements = (slideId: string) => (cb: Listener) =>
    this.subscribe(keyForSlideElements(slideId), cb);
  getSlideElementIds = (slideId: string): string[] => {
    let cached = this.elementIdsCache.get(slideId);
    if (!cached) {
      cached = this.deck.getElementIdsForSlide(slideId);
      this.elementIdsCache.set(slideId, cached);
    }
    return cached;
  };

  subscribeElement = (id: string) => (cb: Listener) => this.subscribe(keyForElement(id), cb);
  getElement = (id: string): SlideElement | undefined => this.deck.getElement(id);

  subscribeSelection = (cb: Listener) => this.subscribe(KEY_SELECTION, cb);
  getSelection = (): DeckSelection => this.selectionCache;

  subscribeActiveSlide = (cb: Listener) => this.subscribe(KEY_ACTIVE_SLIDE, cb);
  getActiveSlideId = (): string => this.activeSlideCache;

  subscribeTheme = (cb: Listener) => this.subscribe(KEY_THEME, cb);
  getTheme = (): ThemeData => this.deck.getTheme();
}
