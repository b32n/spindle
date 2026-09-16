// DeckImpl — the presentation engine (the WorkbookImpl analogue).
//
// Instance-owned flat maps of slides and elements. Records are IMMUTABLE:
// every mutation replaces the object, so React snapshots can version by
// reference. Every mutator: update maps → emit a semantic event →
// recordHistory() when not attached to collab (the Y.Doc mirror + Y.UndoManager
// routing is layered on in Phase 5). Multi-element ops run inside one
// EventEmitter.batch() and one history entry.
//
// Selection and activeSlide are per-user view state: they live here for
// convenience but are LOCAL ONLY — never mirrored into the CRDT doc (a hard
// project rule).

import { EventEmitter } from '@b32nio/spindle-shared';
import { indexBetween, indexesBetween, sortByIndex } from './scene/fractional-index';
import { buildPlaceholderElement } from './theme/materialize';
import type { PlaceholderMeta } from './scene/types';
import { generateId } from './utils/id';
import {
  createTextElement,
  createShapeElement,
  createImageElement,
  createLineElement,
  createTableElement,
  type TextElementInput,
  type ShapeElementInput,
  type ImageElementInput,
  type LineElementInput,
  type TableElementInput,
} from './scene/elements';
import { insertRow, insertColumn, removeRow, removeColumn, resizeColumn, setRowHeight } from './scene/table';
import {
  bringToFront as zBringToFront,
  sendToBack as zSendToBack,
  bringForward as zBringForward,
  sendBackward as zSendBackward,
  type ZItem,
  type ZResult,
} from './scene/z-order';
import { alignFrames, distributeFrames, type AlignMode, type FrameItem } from './scene/align';
import type { Rect } from './scene/geometry';
import { resolveConnectorFrame, resolveEndpoints } from './scene/connector';
import type { Frame, Fill, SlideElement, LineElement, EndpointBind, TableElement, TableCell } from './scene/types';
import { applyTextFormat, type RichTextDoc, type TextFormatSpec } from './text/model';
import type { ThemeData, LayoutData } from './theme/types';
import { getBuiltinTheme, BUILTIN_LAYOUTS, DEFAULT_SLIDE_SIZE } from './theme/builtin';
import { DeckHistory, type DeckSnapshot } from './history';
import { normalizeDeckData } from './serialization';
import { SlidesCommentStore } from './comments';
import type { CollabHandle, AttachCollabOptions } from './collab/binding';
import type { CollabIdentity, CollabProvider } from '@b32nio/spindle-shared';
import type {
  Slide,
  DeckData,
  DeckSelection,
  DeckEventType,
} from './types';

/** A request to add a new element to a slide; the engine assigns id + z-index. */
export type NewElementSpec =
  | ({ type: 'text' } & Omit<TextElementInput, 'containerId' | 'index'>)
  | ({ type: 'shape' } & Omit<ShapeElementInput, 'containerId' | 'index'>)
  | ({ type: 'image' } & Omit<ImageElementInput, 'containerId' | 'index'>)
  | ({ type: 'line' } & Omit<LineElementInput, 'containerId' | 'index'>)
  | ({ type: 'table' } & Omit<TableElementInput, 'containerId' | 'index'>);

const FRAME_KEYS = ['x', 'y', 'w', 'h', 'rotation'] as const;
/** Whether a patch changes any geometry a bound connector must track. */
function touchesFrame(patch: Partial<SlideElement>): boolean {
  return FRAME_KEYS.some((k) => k in patch);
}

export interface AddSlideOptions {
  /** Insert immediately after this slide; defaults to the end. */
  afterSlideId?: string;
  /** Layout to associate with the slide. */
  layoutId?: string;
}

export class DeckImpl {
  id: string;
  private title: string;
  private slideSize: { w: number; h: number };
  private theme: ThemeData;
  private layouts: LayoutData[];

  private slides: Map<string, Slide> = new Map();
  private elements: Map<string, SlideElement> = new Map();

  private selection: DeckSelection = { elementIds: [] };
  private activeSlideId = '';

  private comments = new SlidesCommentStore();
  private events = new EventEmitter<DeckEventType>();
  private history = new DeckHistory();
  private isRestoring = false;
  /** Set by attachCollab. Non-null means a Y.Doc mirror is active. */
  private collabHandle: CollabHandle | null = null;

  constructor(id?: string, title = 'Untitled deck') {
    this.id = id ?? generateId();
    this.title = title;
    this.slideSize = { ...DEFAULT_SLIDE_SIZE };
    this.theme = getBuiltinTheme();
    this.layouts = BUILTIN_LAYOUTS.map((l) => structuredClone(l) as LayoutData);

    // Start with one empty slide so activeSlideId always resolves. Placeholders
    // are materialized when a slide is added with a layout (or via setData).
    const first = this.createSlideRecord({ layoutId: 'blank' });
    this.slides.set(first.id, first);
    this.activeSlideId = first.id;
    this.selection = { slideId: first.id, elementIds: [] };
    this.wireCommentListener();
  }

  private wireCommentListener(): void {
    this.comments.__setChangeListener((event) => {
      // commentEvent fires for the local user's own actions only (the store
      // notifies on API mutations, not on loadJSON from a remote/setData).
      this.emit('commentEvent', event);
      this.emit('commentChange', {});
    });
  }

  // ── Events ─────────────────────────────────────────────────────────────────

  on(event: DeckEventType, handler: (data: { type: DeckEventType; payload: unknown }) => void): () => void {
    return this.events.on(event, handler);
  }

  off(event: DeckEventType, handler: (data: { type: DeckEventType; payload: unknown }) => void): void {
    this.events.off(event, handler);
  }

  private emit(event: DeckEventType, payload: unknown): void {
    this.events.emit(event, payload);
  }

  // ── Deck-level ───────────────────────────────────────────────────────────────

  getTitle(): string {
    return this.title;
  }

  setTitle(title: string): void {
    if (title === this.title) return;
    this.recordHistory();
    this.title = title;
    this.emit('deckChange', { keys: ['title'] });
  }

  getSlideSize(): { w: number; h: number } {
    return { ...this.slideSize };
  }

  setSlideSize(size: { w: number; h: number }): void {
    this.recordHistory();
    this.slideSize = { ...size };
    this.emit('deckChange', { keys: ['slideSize'] });
  }

  getTheme(): ThemeData {
    return this.theme;
  }

  setTheme(theme: ThemeData): void {
    this.recordHistory();
    this.theme = theme;
    this.emit('themeChange', { theme });
  }

  getLayouts(): LayoutData[] {
    return this.layouts;
  }

  getLayout(id: string): LayoutData | undefined {
    return this.layouts.find((l) => l.id === id);
  }

  // ── Slides ─────────────────────────────────────────────────────────────────

  /** Slide ids in presentation order. */
  getSlideIds(): string[] {
    return sortByIndex([...this.slides.values()]).map((s) => s.id);
  }

  /** Slides in presentation order. */
  getSlides(): Slide[] {
    return sortByIndex([...this.slides.values()]);
  }

  getSlide(id: string): Slide | undefined {
    return this.slides.get(id);
  }

  slideCount(): number {
    return this.slides.size;
  }

  private createSlideRecord(opts: { index?: string; layoutId?: string }): Slide {
    const slide: Slide = {
      id: generateId(),
      index: opts.index ?? indexBetween(null, null),
    };
    if (opts.layoutId) slide.layoutRef = opts.layoutId;
    return slide;
  }

  /** Fractional index that sorts a slide immediately after `afterSlideId`. */
  private slideIndexAfter(afterSlideId?: string): string {
    const ordered = this.getSlides();
    if (!afterSlideId) {
      const last = ordered[ordered.length - 1];
      return indexBetween(last ? last.index : null, null);
    }
    const pos = ordered.findIndex((s) => s.id === afterSlideId);
    if (pos === -1) {
      const last = ordered[ordered.length - 1];
      return indexBetween(last ? last.index : null, null);
    }
    const after = ordered[pos];
    const next = ordered[pos + 1];
    return indexBetween(after.index, next ? next.index : null);
  }

  addSlide(opts: AddSlideOptions = {}): Slide {
    this.recordHistory();
    const slide = this.createSlideRecord({
      index: this.slideIndexAfter(opts.afterSlideId),
      layoutId: opts.layoutId,
    });
    this.slides.set(slide.id, slide);
    this.emit('slideAdd', { slideId: slide.id });
    if (opts.layoutId) this.materializePlaceholders(slide.id, opts.layoutId);
    return slide;
  }

  /** Create real placeholder elements (and background) for a slide's layout. */
  private materializePlaceholders(slideId: string, layoutId: string): void {
    const layout = this.getLayout(layoutId);
    if (!layout) return;
    if (layout.background) {
      const slide = this.slides.get(slideId);
      if (slide) this.slides.set(slideId, { ...slide, background: layout.background });
    }
    if (layout.placeholders.length === 0) return;
    const keys = indexesBetween(null, null, layout.placeholders.length);
    layout.placeholders.forEach((ph, i) => {
      const el = buildPlaceholderElement(ph, slideId, keys[i]);
      this.elements.set(el.id, el);
      this.emit('elementAdd', { slideId, elementId: el.id });
    });
  }

  /** The layout prompt text for a placeholder element, if any (for rendering). */
  getPlaceholderPrompt(slideId: string, placeholder: PlaceholderMeta | undefined): string | undefined {
    if (!placeholder) return undefined;
    const slide = this.slides.get(slideId);
    if (!slide?.layoutRef) return undefined;
    const layout = this.getLayout(slide.layoutRef);
    return layout?.placeholders.find((p) => p.type === placeholder.type && p.idx === placeholder.idx)?.prompt;
  }

  /** Duplicate a slide and all its elements, inserting the copy just after it. */
  duplicateSlide(slideId: string): Slide | undefined {
    const src = this.slides.get(slideId);
    if (!src) return undefined;
    this.recordHistory();
    const copy: Slide = {
      ...structuredClone(src),
      id: generateId(),
      index: this.slideIndexAfter(slideId),
    };
    this.slides.set(copy.id, copy);
    // Announce the new slide before its elements so the collab mirror has the
    // slide's Y.Map in place when each elementAdd lands.
    this.emit('slideAdd', { slideId: copy.id });
    const srcElements = this.getElementsForSlide(slideId);
    // Map every source element id to its copy's id first, so internal references
    // (connector binds) can be rewritten to point at the copies rather than
    // lingering on the originals. Group ids get their own fresh id per source
    // group, so the copied group is a distinct group of only the copies.
    const idMap = new Map<string, string>();
    for (const el of srcElements) idMap.set(el.id, generateId());
    const groupMap = new Map<string, string>();
    for (const el of srcElements) {
      const elCopy = { ...structuredClone(el), id: idMap.get(el.id)!, containerId: copy.id } as SlideElement;
      // Give the copied group a fresh shared id so getGroupMembers on the copy
      // returns only the copies (not the source slide's members).
      if (elCopy.groupId) {
        let mapped = groupMap.get(elCopy.groupId);
        if (!mapped) { mapped = generateId(); groupMap.set(elCopy.groupId, mapped); }
        elCopy.groupId = mapped;
      }
      // Re-point connector binds at the copied shapes; leave binds to off-slide
      // elements untouched.
      if (elCopy.type === 'line') {
        const line = elCopy as LineElement;
        if (line.startBind && idMap.has(line.startBind.elementId)) {
          line.startBind = { ...line.startBind, elementId: idMap.get(line.startBind.elementId)! };
        }
        if (line.endBind && idMap.has(line.endBind.elementId)) {
          line.endBind = { ...line.endBind, elementId: idMap.get(line.endBind.elementId)! };
        }
      }
      this.elements.set(elCopy.id, elCopy);
      // Emit per copied element so the collab mirror writes it into the Y.Doc —
      // otherwise a duplicated slide syncs (and later reloads) with no elements.
      this.emit('elementAdd', { slideId: copy.id, elementId: elCopy.id });
    }
    return copy;
  }

  deleteSlide(slideId: string): void {
    if (!this.slides.has(slideId)) return;
    if (this.slides.size <= 1) throw new Error('Cannot delete the last slide');
    this.recordHistory();
    // Remove the slide's elements too (emit per-element so collab + React see it).
    for (const el of [...this.elements.values()]) {
      if (el.containerId === slideId) {
        this.elements.delete(el.id);
        this.emit('elementDelete', { slideId, elementId: el.id });
      }
    }
    this.slides.delete(slideId);
    if (this.activeSlideId === slideId) {
      this.activeSlideId = this.getSlideIds()[0];
      this.selection = { slideId: this.activeSlideId, elementIds: [] };
      this.emit('activeSlideChange', { slideId: this.activeSlideId });
      this.emit('selectionChange', { selection: this.getSelection() });
    }
    this.emit('slideDelete', { slideId });
  }

  /** Move a slide so it sorts immediately after `afterSlideId` (or to the front). */
  moveSlide(slideId: string, opts: { afterSlideId?: string } = {}): void {
    const slide = this.slides.get(slideId);
    if (!slide) return;
    const ordered = this.getSlides().filter((s) => s.id !== slideId);
    let index: string;
    if (!opts.afterSlideId) {
      index = indexBetween(null, ordered[0] ? ordered[0].index : null);
    } else {
      const pos = ordered.findIndex((s) => s.id === opts.afterSlideId);
      const after = ordered[pos];
      const next = ordered[pos + 1];
      index = indexBetween(after ? after.index : null, next ? next.index : null);
    }
    this.recordHistory();
    this.slides.set(slideId, { ...slide, index });
    this.emit('slideMove', { slideId });
  }

  setSlideBackground(slideId: string, background: Fill | undefined): void {
    const slide = this.slides.get(slideId);
    if (!slide) return;
    this.recordHistory();
    const next = { ...slide };
    if (background) next.background = background;
    else delete next.background;
    this.slides.set(slideId, next);
    this.emit('slideChange', { slideId, keys: ['background'] });
  }

  setSlideNotes(slideId: string, notes: RichTextDoc): void {
    const slide = this.slides.get(slideId);
    if (!slide) return;
    this.recordHistory();
    this.slides.set(slideId, { ...slide, notes });
    this.emit('slideChange', { slideId, keys: ['notes'] });
  }

  /** Associate a slide with a layout (placeholder materialization is Phase 4). */
  setSlideLayout(slideId: string, layoutId: string): void {
    const slide = this.slides.get(slideId);
    if (!slide) return;
    this.recordHistory();
    this.slides.set(slideId, { ...slide, layoutRef: layoutId });
    this.emit('slideChange', { slideId, keys: ['layoutRef'] });
  }

  // ── Elements ─────────────────────────────────────────────────────────────────

  getElement(id: string): SlideElement | undefined {
    return this.elements.get(id);
  }

  /** Elements on a slide, sorted back-to-front by z-order. */
  getElementsForSlide(slideId: string): SlideElement[] {
    const list = [...this.elements.values()].filter((e) => e.containerId === slideId);
    return sortByIndex(list);
  }

  getElementIdsForSlide(slideId: string): string[] {
    return this.getElementsForSlide(slideId).map((e) => e.id);
  }

  /** Fractional z-index that sorts above every element currently on a slide. */
  private topZIndex(slideId: string): string {
    const list = this.getElementsForSlide(slideId);
    const top = list[list.length - 1];
    return indexBetween(top ? top.index : null, null);
  }

  addElement(slideId: string, spec: NewElementSpec, opts?: { index?: string }): SlideElement {
    if (!this.slides.has(slideId)) throw new Error(`Slide not found: ${slideId}`);
    const index = opts?.index ?? this.topZIndex(slideId);
    const common = { containerId: slideId, index };
    let el: SlideElement;
    switch (spec.type) {
      case 'text':
        el = createTextElement({ ...spec, ...common });
        break;
      case 'shape':
        el = createShapeElement({ ...spec, ...common });
        break;
      case 'image':
        el = createImageElement({ ...spec, ...common });
        break;
      case 'line':
        el = createLineElement({ ...spec, ...common });
        break;
      case 'table':
        el = createTableElement({ ...spec, ...common });
        break;
    }
    this.recordHistory();
    this.elements.set(el.id, el);
    this.emit('elementAdd', { slideId, elementId: el.id });
    return el;
  }

  /** Patch an element; the record is replaced (immutable). */
  updateElement(id: string, patch: Partial<SlideElement>): void {
    const el = this.elements.get(id);
    if (!el) return;
    this.recordHistory();
    this.events.batch(() => {
      this.applyElementPatch(id, patch);
      this.emit('elementChange', {
        slideId: el.containerId,
        elementId: id,
        keys: Object.keys(patch),
      });
      if (touchesFrame(patch)) this.reconcileConnectors(new Set([id]));
    });
  }

  /**
   * Set an element's height to match its rendered content (a table whose rows
   * are content-driven). This is a layout reflow, not a user edit, so it records
   * NO history entry — otherwise undo would fight the reflow and the selection/
   * gutter overlays (which read the frame) wouldn't line up with the table.
   * A no-op if the height already matches (within a pixel).
   */
  syncElementHeight(id: string, h: number): void {
    const el = this.elements.get(id);
    if (!el || Math.abs(el.h - h) < 0.5) return;
    this.events.batch(() => {
      this.applyElementPatch(id, { h } as Partial<SlideElement>);
      this.emit('elementChange', { slideId: el.containerId, elementId: id, keys: ['h'] });
      this.reconcileConnectors(new Set([id]));
    });
  }

  /** Patch several elements as one undo entry and one event batch. */
  updateElements(patches: Array<{ id: string; patch: Partial<SlideElement> }>): void {
    this.recordHistory();
    this.events.batch(() => {
      const framed = new Set<string>();
      for (const { id, patch } of patches) {
        const el = this.elements.get(id);
        if (!el) continue;
        this.applyElementPatch(id, patch);
        this.emit('elementChange', {
          slideId: el.containerId,
          elementId: id,
          keys: Object.keys(patch),
        });
        if (touchesFrame(patch)) framed.add(id);
      }
      if (framed.size) this.reconcileConnectors(framed);
    });
  }

  /**
   * Recompute the box of every connector bound to one of `changedIds`, so it
   * keeps tracking its shape. Applied in-place inside the caller's history/emit
   * cycle (no new undo entry). Connectors bound to nothing that moved, and
   * connectors whose own box was just set directly, are left alone.
   */
  private reconcileConnectors(changedIds: Set<string>): void {
    const getFrame = (elementId: string): Frame | undefined => {
      const e = this.elements.get(elementId);
      return e ? { x: e.x, y: e.y, w: e.w, h: e.h, rotation: e.rotation } : undefined;
    };
    for (const el of this.elements.values()) {
      if (el.type !== 'line') continue;
      const line = el as LineElement;
      if (changedIds.has(line.id)) continue;
      const hit =
        (line.startBind && changedIds.has(line.startBind.elementId)) ||
        (line.endBind && changedIds.has(line.endBind.elementId));
      if (!hit) continue;
      const box = resolveConnectorFrame(line, getFrame);
      this.applyElementPatch(line.id, box as Partial<SlideElement>);
      this.emit('elementChange', { slideId: line.containerId, elementId: line.id, keys: ['x', 'y', 'w', 'h', 'flipH', 'flipV'] });
    }
  }

  /** Clear any connector endpoints bound to `deletedIds`, pinning them at their
   *  last resolved position (an explicit point) so they stay put rather than
   *  reverting to the ambiguous box corner. Applied in-place; call inside a
   *  history/emit cycle. */
  private detachConnectors(deletedIds: Set<string>): void {
    const getFrame = (elementId: string): Frame | undefined => {
      const e = this.elements.get(elementId);
      return e ? { x: e.x, y: e.y, w: e.w, h: e.h, rotation: e.rotation } : undefined;
    };
    for (const el of this.elements.values()) {
      if (el.type !== 'line') continue;
      const line = el as LineElement;
      if (deletedIds.has(line.id)) continue;
      const startGone = line.startBind && deletedIds.has(line.startBind.elementId);
      const endGone = line.endBind && deletedIds.has(line.endBind.elementId);
      if (!startGone && !endGone) continue;
      const { start, end } = resolveEndpoints(line, getFrame);
      const patch: Partial<LineElement> = {};
      if (startGone) { patch.startBind = undefined; patch.startPoint = { x: start.x, y: start.y }; }
      if (endGone) { patch.endBind = undefined; patch.endPoint = { x: end.x, y: end.y }; }
      this.applyElementPatch(line.id, patch as Partial<SlideElement>);
      this.emit('elementChange', { slideId: line.containerId, elementId: line.id, keys: Object.keys(patch) });
    }
  }

  private applyElementPatch(id: string, patch: Partial<SlideElement>): void {
    const el = this.elements.get(id);
    if (!el) return;
    // Merge without changing discriminant; cast is safe because callers never
    // change `type`.
    this.elements.set(id, { ...el, ...patch } as SlideElement);
  }

  deleteElement(id: string): void {
    const el = this.elements.get(id);
    if (!el) return;
    this.recordHistory();
    this.events.batch(() => {
      this.elements.delete(id);
      this.emit('elementDelete', { slideId: el.containerId, elementId: id });
      this.detachConnectors(new Set([id]));
    });
  }

  deleteElements(ids: string[]): void {
    if (ids.length === 0) return;
    this.recordHistory();
    this.events.batch(() => {
      const deleted = new Set<string>();
      for (const id of ids) {
        const el = this.elements.get(id);
        if (!el) continue;
        this.elements.delete(id);
        this.emit('elementDelete', { slideId: el.containerId, elementId: id });
        deleted.add(id);
      }
      this.detachConnectors(deleted);
    });
  }

  duplicateElement(id: string): SlideElement | undefined {
    const el = this.elements.get(id);
    if (!el) return undefined;
    this.recordHistory();
    const copy = {
      ...structuredClone(el),
      id: generateId(),
      index: this.topZIndex(el.containerId),
      x: el.x + 16,
      y: el.y + 16,
    } as SlideElement;
    // A line's box shifted, so shift its explicit endpoints too (mirror
    // moveElements) — otherwise the box and its diagonal desync.
    if (copy.type === 'line') {
      const l = copy as LineElement;
      if (l.startPoint) l.startPoint = { x: l.startPoint.x + 16, y: l.startPoint.y + 16 };
      if (l.endPoint) l.endPoint = { x: l.endPoint.x + 16, y: l.endPoint.y + 16 };
    }
    delete copy.groupId;
    this.elements.set(copy.id, copy);
    this.emit('elementAdd', { slideId: el.containerId, elementId: copy.id });
    return copy;
  }

  /**
   * Duplicate several elements as one action, preserving grouping (a fresh
   * shared group id per source group) and re-pointing connector binds among the
   * copied set — so Cmd+D of a group yields a grouped copy, not loose elements.
   */
  duplicateElements(ids: string[]): SlideElement[] {
    const src = ids.map((id) => this.elements.get(id)).filter(Boolean) as SlideElement[];
    if (src.length === 0) return [];
    this.recordHistory();
    const idMap = new Map<string, string>();
    for (const el of src) idMap.set(el.id, generateId());
    const groupMap = new Map<string, string>();
    const copies: SlideElement[] = [];
    for (const el of src) {
      const copy = {
        ...structuredClone(el),
        id: idMap.get(el.id)!,
        index: this.topZIndex(el.containerId),
        x: el.x + 16,
        y: el.y + 16,
      } as SlideElement;
      if (copy.groupId) {
        let mapped = groupMap.get(copy.groupId);
        if (!mapped) { mapped = generateId(); groupMap.set(copy.groupId, mapped); }
        copy.groupId = mapped;
      }
      if (copy.type === 'line') {
        const l = copy as LineElement;
        if (l.startPoint) l.startPoint = { x: l.startPoint.x + 16, y: l.startPoint.y + 16 };
        if (l.endPoint) l.endPoint = { x: l.endPoint.x + 16, y: l.endPoint.y + 16 };
        if (l.startBind && idMap.has(l.startBind.elementId)) l.startBind = { ...l.startBind, elementId: idMap.get(l.startBind.elementId)! };
        if (l.endBind && idMap.has(l.endBind.elementId)) l.endBind = { ...l.endBind, elementId: idMap.get(l.endBind.elementId)! };
      }
      this.elements.set(copy.id, copy);
      copies.push(copy);
    }
    for (const c of copies) this.emit('elementAdd', { slideId: c.containerId, elementId: c.id });
    return copies;
  }

  // ── Rich text ────────────────────────────────────────────────────────────────

  /** Replace a text-bearing element's body (commits one history entry). */
  setElementRichText(id: string, doc: RichTextDoc): void {
    const el = this.elements.get(id);
    if (!el || (el.type !== 'text' && el.type !== 'shape')) return;
    this.updateElement(id, { richText: doc } as Partial<SlideElement>);
  }

  /** Apply an idle-path text format across an element's whole body. */
  applyTextFormat(id: string, spec: TextFormatSpec): void {
    const el = this.elements.get(id);
    if (!el) return;
    const rich = (el as { richText?: RichTextDoc }).richText;
    if (!rich) return;
    this.setElementRichText(id, applyTextFormat(rich, spec));
  }

  // ── Transforms (gesture-commit path; Phase 2 builds the gestures) ────────────

  /** Commit new frames for one or more elements as a single undo entry. */
  setFrames(frames: Array<{ id: string; frame: Partial<Frame> }>): void {
    this.updateElements(frames.map(({ id, frame }) => ({ id, patch: frame as Partial<SlideElement> })));
  }

  /** Translate several elements by (dx, dy). Line connectors also shift their
   *  explicit free endpoints (bound ends stay on their shapes via reconcile). */
  moveElements(ids: string[], dx: number, dy: number): void {
    const patches = ids
      .map((id) => this.elements.get(id))
      .filter((el): el is SlideElement => !!el)
      .map((el) => {
        const patch: Partial<SlideElement> = { x: el.x + dx, y: el.y + dy };
        if (el.type === 'line') {
          const l = el as LineElement;
          if (l.startPoint) (patch as Partial<LineElement>).startPoint = { x: l.startPoint.x + dx, y: l.startPoint.y + dy };
          if (l.endPoint) (patch as Partial<LineElement>).endPoint = { x: l.endPoint.x + dx, y: l.endPoint.y + dy };
        }
        return { id: el.id, patch };
      });
    this.updateElements(patches);
  }

  /**
   * Set one end of a line to either bind to an element's anchor or a free point.
   * Converts the other end to an explicit point if it was a bare box corner, so
   * the geometry (and arrow direction) survives; recomputes the box. One undo.
   */
  setLineEndpoint(id: string, end: 'start' | 'end', to: { bind?: EndpointBind; point?: { x: number; y: number } }): void {
    const el = this.elements.get(id);
    if (!el || el.type !== 'line') return;
    const line = el as LineElement;
    const getFrame = (elementId: string): Frame | undefined => {
      const e = this.elements.get(elementId);
      return e ? { x: e.x, y: e.y, w: e.w, h: e.h, rotation: e.rotation } : undefined;
    };
    const resolved = resolveEndpoints(line, getFrame);
    const patch: Partial<LineElement> = {};
    // Freeze the untouched end if it's still a bare box corner.
    if (end === 'end' && !line.startBind && !line.startPoint) patch.startPoint = { x: resolved.start.x, y: resolved.start.y };
    if (end === 'start' && !line.endBind && !line.endPoint) patch.endPoint = { x: resolved.end.x, y: resolved.end.y };
    // Apply the edited end.
    if (end === 'start') {
      if (to.bind) { patch.startBind = to.bind; patch.startPoint = undefined; }
      else { patch.startPoint = to.point; patch.startBind = undefined; }
    } else {
      if (to.bind) { patch.endBind = to.bind; patch.endPoint = undefined; }
      else { patch.endPoint = to.point; patch.endBind = undefined; }
    }
    const box = resolveConnectorFrame({ ...line, ...patch } as LineElement, getFrame);
    Object.assign(patch, box);
    this.updateElement(id, patch as Partial<SlideElement>);
  }

  // ── Tables (thin wrappers over scene/table pure fns) ─────────────────────────

  private tableOp(id: string, op: (t: TableElement) => Partial<TableElement>): void {
    const el = this.elements.get(id);
    if (!el || el.type !== 'table') return;
    const patch = op(el as TableElement);
    if (Object.keys(patch).length) this.updateElement(id, patch as Partial<SlideElement>);
  }

  insertTableRow(id: string, at: number): void { this.tableOp(id, (t) => insertRow(t, at)); }
  insertTableColumn(id: string, at: number): void { this.tableOp(id, (t) => insertColumn(t, at)); }
  removeTableRow(id: string, at: number): void { this.tableOp(id, (t) => removeRow(t, at)); }
  removeTableColumn(id: string, at: number): void { this.tableOp(id, (t) => removeColumn(t, at)); }

  /** Remove an inclusive range of rows/columns (a whole selection) in one undo. */
  removeTableRows(id: string, from: number, to: number): void {
    this.tableOp(id, (t) => {
      let cur = t;
      for (let r = to; r >= from; r--) cur = { ...cur, ...removeRow(cur, r) } as TableElement;
      // Include rowHeights — removeRow updates it, and dropping it here leaves a
      // stale, index-misaligned array (wrong per-row min heights render).
      return { rows: cur.rows, rowFractions: cur.rowFractions, cells: cur.cells, rowHeights: cur.rowHeights };
    });
  }
  removeTableColumns(id: string, from: number, to: number): void {
    this.tableOp(id, (t) => {
      let cur = t;
      for (let c = to; c >= from; c--) cur = { ...cur, ...removeColumn(cur, c) } as TableElement;
      return { cols: cur.cols, colFractions: cur.colFractions, cells: cur.cells };
    });
  }
  resizeTableColumn(id: string, index: number, delta: number): void { this.tableOp(id, (t) => resizeColumn(t, index, delta)); }
  /** Set row `index`'s minimum height in px (content can grow it further). */
  setTableRowHeight(id: string, index: number, height: number): void { this.tableOp(id, (t) => setRowHeight(t, index, height)); }

  /** Replace one cell's rich text (cell editing commits through here). */
  setTableCellRichText(id: string, row: number, col: number, doc: RichTextDoc): void {
    this.updateTableCell(id, row, col, { richText: doc });
  }

  /** Merge a patch into one cell (rich text, fill, body style). */
  updateTableCell(id: string, row: number, col: number, patch: Partial<TableCell>): void {
    this.updateTableCells(id, [[row, col]], patch);
  }

  /** Merge a patch into many cells at once as a single undo (e.g. fill a whole
   *  row/column/range). Cells outside the grid are ignored. */
  updateTableCells(id: string, cells: ReadonlyArray<readonly [number, number]>, patch: Partial<TableCell>): void {
    this.tableOp(id, (t) => {
      const set = new Set<number>();
      for (const [r, c] of cells) if (r >= 0 && r < t.rows && c >= 0 && c < t.cols) set.add(r * t.cols + c);
      if (!set.size) return {};
      const grid = t.cells.map((r, ri) => r.map((c, ci) => (set.has(ri * t.cols + ci) ? { ...c, ...patch } : c)));
      return { cells: grid };
    });
  }

  /** Apply a text format across many cells' rich text as a single undo (e.g.
   *  bold an entire header row). */
  applyTableCellsFormat(id: string, cells: ReadonlyArray<readonly [number, number]>, spec: TextFormatSpec): void {
    this.tableOp(id, (t) => {
      const set = new Set<number>();
      for (const [r, c] of cells) if (r >= 0 && r < t.rows && c >= 0 && c < t.cols) set.add(r * t.cols + c);
      if (!set.size) return {};
      const grid = t.cells.map((r, ri) =>
        r.map((c, ci) => (set.has(ri * t.cols + ci) ? { ...c, richText: applyTextFormat(c.richText, spec) } : c))
      );
      return { cells: grid };
    });
  }

  // ── Z-order / align / group (thin wrappers over scene/* pure fns) ────────────

  private slideIdOf(ids: string[]): string | undefined {
    for (const id of ids) {
      const el = this.elements.get(id);
      if (el) return el.containerId;
    }
    return undefined;
  }

  private zItemsFor(ids: string[]): ZItem[] {
    const slideId = this.slideIdOf(ids);
    if (!slideId) return [];
    return this.getElementsForSlide(slideId).map((e) => ({ id: e.id, index: e.index }));
  }

  private applyZResult(res: ZResult): void {
    if (res.length) this.updateElements(res.map((r) => ({ id: r.id, patch: { index: r.index } })));
  }

  bringToFront(ids: string[]): void {
    this.applyZResult(zBringToFront(this.zItemsFor(ids), ids));
  }
  sendToBack(ids: string[]): void {
    this.applyZResult(zSendToBack(this.zItemsFor(ids), ids));
  }
  bringForward(ids: string[]): void {
    this.applyZResult(zBringForward(this.zItemsFor(ids), ids));
  }
  sendBackward(ids: string[]): void {
    this.applyZResult(zSendBackward(this.zItemsFor(ids), ids));
  }

  private frameItemsFor(ids: string[]): FrameItem[] {
    return ids
      .map((id) => this.elements.get(id))
      .filter((el): el is SlideElement => !!el)
      .map((el) => ({ id: el.id, frame: { x: el.x, y: el.y, w: el.w, h: el.h, rotation: el.rotation } }));
  }

  alignElements(ids: string[], mode: AlignMode, bounds?: Rect): void {
    const res = alignFrames(this.frameItemsFor(ids), mode, bounds);
    this.setFrames(res.map((r) => ({ id: r.id, frame: r.frame })));
  }

  distributeElements(ids: string[], axis: 'h' | 'v'): void {
    const res = distributeFrames(this.frameItemsFor(ids), axis);
    this.setFrames(res.map((r) => ({ id: r.id, frame: r.frame })));
  }

  /** Group the given elements under a fresh shared groupId (needs ≥2). */
  groupElements(ids: string[]): string | undefined {
    const valid = ids.filter((id) => this.elements.has(id));
    if (valid.length < 2) return undefined;
    const groupId = generateId();
    this.updateElements(valid.map((id) => ({ id, patch: { groupId } })));
    return groupId;
  }

  ungroupElements(ids: string[]): void {
    const patches = ids
      .filter((id) => this.elements.get(id)?.groupId)
      .map((id) => ({ id, patch: { groupId: undefined } as Partial<SlideElement> }));
    if (patches.length) this.updateElements(patches);
  }

  /** All element ids sharing a groupId. */
  getGroupMembers(groupId: string): string[] {
    return [...this.elements.values()].filter((e) => e.groupId === groupId).map((e) => e.id);
  }

  // ── Selection & active slide (LOCAL view state; never serialized to CRDT) ────

  getSelection(): DeckSelection {
    return { slideId: this.selection.slideId, elementIds: [...this.selection.elementIds] };
  }

  setSelection(selection: DeckSelection): void {
    this.selection = { slideId: selection.slideId, elementIds: [...selection.elementIds] };
    this.emit('selectionChange', { selection: this.getSelection() });
  }

  getActiveSlideId(): string {
    return this.activeSlideId;
  }

  setActiveSlide(slideId: string): void {
    if (!this.slides.has(slideId) || slideId === this.activeSlideId) return;
    this.activeSlideId = slideId;
    // The element selection belonged to the previous slide — clear it so its
    // bounding box doesn't linger on the new slide.
    this.selection = { slideId, elementIds: [] };
    this.emit('activeSlideChange', { slideId });
    this.emit('selectionChange', { selection: this.getSelection() });
  }

  // ── History ────────────────────────────────────────────────────────────────

  private createSnapshot(): DeckSnapshot {
    return {
      title: this.title,
      slideSize: { ...this.slideSize },
      theme: this.theme,
      layouts: this.layouts,
      slides: new Map(this.slides),
      elements: new Map(this.elements),
      selection: this.getSelection(),
      activeSlideId: this.activeSlideId,
    };
  }

  private recordHistory(): void {
    if (this.isRestoring) return;
    if (this.collabHandle) return; // collab mode uses Y.UndoManager (Phase 5)
    this.history.record(this.createSnapshot());
  }

  private restoreSnapshot(snap: DeckSnapshot): void {
    this.isRestoring = true;
    try {
      this.title = snap.title;
      this.slideSize = { ...snap.slideSize };
      this.theme = snap.theme;
      this.layouts = snap.layouts;
      this.slides = new Map(snap.slides);
      this.elements = new Map(snap.elements);
      this.selection = { slideId: snap.selection.slideId, elementIds: [...snap.selection.elementIds] };
      if (this.slides.has(snap.activeSlideId)) this.activeSlideId = snap.activeSlideId;
      else this.activeSlideId = this.getSlideIds()[0] ?? '';
    } finally {
      this.isRestoring = false;
    }
    // Undo is not on the hot path; a broad invalidation is fine here.
    this.emit('deckChange', { keys: ['*'] });
    this.emit('activeSlideChange', { slideId: this.activeSlideId });
    this.emit('selectionChange', { selection: this.getSelection() });
  }

  undo(): boolean {
    if (this.collabHandle) return this.collabHandle.undoManager.undo() !== null;
    const prev = this.history.undo(this.createSnapshot());
    if (!prev) return false;
    this.restoreSnapshot(prev);
    return true;
  }

  redo(): boolean {
    if (this.collabHandle) return this.collabHandle.undoManager.redo() !== null;
    const next = this.history.redo(this.createSnapshot());
    if (!next) return false;
    this.restoreSnapshot(next);
    return true;
  }

  canUndo(): boolean {
    return this.collabHandle ? this.collabHandle.undoManager.canUndo() : this.history.canUndo();
  }

  canRedo(): boolean {
    return this.collabHandle ? this.collabHandle.undoManager.canRedo() : this.history.canRedo();
  }

  // ── Collaboration ────────────────────────────────────────────────────────────

  // ── Comments ─────────────────────────────────────────────────────────────────

  getComments(): SlidesCommentStore {
    return this.comments;
  }

  /** True when a thread's anchored element no longer exists (orphaned). */
  isThreadOrphaned(threadId: string): boolean {
    const t = this.comments.getThread(threadId);
    return !!t && !this.elements.has(t.anchor.elementId);
  }

  /** Replace the comment store from remote-synced JSON, then notify the UI. */
  _applyRemoteComments(threads: import('./types').SlidesCommentThread[]): void {
    this.comments.loadJSON(threads);
    this.emit('commentChange', {});
  }

  isCollabAttached(): boolean {
    return this.collabHandle !== null;
  }

  getCollabHandle(): CollabHandle | null {
    return this.collabHandle;
  }

  /** Attach a Yjs mirror over a provider. Lazy-loads ./collab (keeps Yjs out of the base bundle). */
  async attachCollab(provider: CollabProvider, identity: CollabIdentity, options?: AttachCollabOptions): Promise<CollabHandle> {
    if (this.collabHandle) return this.collabHandle;
    const { attachCollabToDeck } = await import('./collab/binding');
    this.collabHandle = await attachCollabToDeck(this, provider, identity, options);
    this.history.clear(); // collab uses Y.UndoManager
    return this.collabHandle;
  }

  detachCollab(): void {
    this.collabHandle?.detach();
    this.collabHandle = null;
  }

  // Remote-apply hooks — called ONLY by the collab binding while applying a
  // remote (or undo) change. They mutate maps + emit events but never record
  // history; the binding suppresses mirror echo around these calls.

  _applyRemoteElementUpsert(el: SlideElement): void {
    const existed = this.elements.has(el.id);
    this.elements.set(el.id, el);
    if (existed) this.emit('elementChange', { slideId: el.containerId, elementId: el.id, keys: ['*'] });
    else this.emit('elementAdd', { slideId: el.containerId, elementId: el.id });
  }

  _applyRemoteElementDelete(id: string): void {
    const el = this.elements.get(id);
    if (!el) return;
    this.elements.delete(id);
    this.emit('elementDelete', { slideId: el.containerId, elementId: id });
  }

  _applyRemoteSlideUpsert(slide: Slide): void {
    const existed = this.slides.has(slide.id);
    this.slides.set(slide.id, slide);
    if (existed) this.emit('slideChange', { slideId: slide.id, keys: ['*'] });
    else this.emit('slideAdd', { slideId: slide.id });
  }

  _applyRemoteSlideDelete(id: string): void {
    if (!this.slides.has(id)) return;
    this.slides.delete(id);
    if (this.activeSlideId === id) {
      this.activeSlideId = this.getSlideIds()[0] ?? '';
      this.emit('activeSlideChange', { slideId: this.activeSlideId });
    }
    this.emit('slideDelete', { slideId: id });
  }

  /** Escape hatch: replace the whole engine state from a Y-derived DeckData
   *  (used on attach to reconcile, and available if observers ever drift). */
  _resyncFromY(data: DeckData): void {
    this.id = data.id;
    this.title = data.title;
    this.slideSize = { ...data.slideSize };
    if (data.theme) this.theme = data.theme;
    if (data.layouts) this.layouts = data.layouts;
    this.slides = new Map();
    this.elements = new Map();
    for (const slideData of data.slides) {
      const { elements, ...slide } = slideData;
      this.slides.set(slide.id, slide);
      for (const el of elements) this.elements.set(el.id, el);
    }
    this.comments.loadJSON(data.threads);
    const first = this.getSlideIds()[0] ?? '';
    if (!this.slides.has(this.activeSlideId)) this.activeSlideId = first;
    this.selection = { slideId: this.activeSlideId, elementIds: [] };
    this.emit('deckChange', { keys: ['*'] });
    this.emit('themeChange', { theme: this.theme });
    this.emit('activeSlideChange', { slideId: this.activeSlideId });
    this.emit('selectionChange', { selection: this.getSelection() });
  }

  _applyRemoteMeta(meta: { title?: string; slideSize?: { w: number; h: number }; theme?: ThemeData; layouts?: LayoutData[] }): void {
    let deckChanged = false;
    if (meta.title !== undefined && meta.title !== this.title) { this.title = meta.title; deckChanged = true; }
    if (meta.slideSize) { this.slideSize = { w: meta.slideSize.w, h: meta.slideSize.h }; deckChanged = true; }
    if (meta.layouts) { this.layouts = meta.layouts; deckChanged = true; }
    if (meta.theme) { this.theme = meta.theme; this.emit('themeChange', { theme: this.theme }); }
    if (deckChanged) this.emit('deckChange', { keys: ['*'] });
  }

  // ── Serialization ────────────────────────────────────────────────────────────

  getData(): DeckData {
    const slides = this.getSlides().map((slide) => ({
      ...structuredClone(slide),
      elements: this.getElementsForSlide(slide.id).map((el) => structuredClone(el)),
    }));
    return {
      id: this.id,
      title: this.title,
      slideSize: { ...this.slideSize },
      theme: structuredClone(this.theme),
      layouts: this.layouts.map((l) => structuredClone(l)),
      slides,
      threads: this.comments.toJSON().length ? this.comments.toJSON() : undefined,
      selection: this.getSelection(),
    };
  }

  setData(data: DeckData): void {
    if (this.collabHandle) throw new Error('Cannot setData while collab is attached');
    const normalized = normalizeDeckData(data);
    this.id = normalized.id;
    this.title = normalized.title;
    this.slideSize = { ...normalized.slideSize };
    this.theme = normalized.theme;
    this.layouts = normalized.layouts ?? BUILTIN_LAYOUTS.map((l) => structuredClone(l) as LayoutData);

    this.slides = new Map();
    this.elements = new Map();
    for (const slideData of normalized.slides) {
      const { elements, ...slide } = slideData;
      this.slides.set(slide.id, slide);
      for (const el of elements) {
        this.elements.set(el.id, el);
      }
    }
    this.comments.loadJSON(normalized.threads);

    this.history.clear();

    // Restore or repair local view state.
    const firstId = this.getSlideIds()[0] ?? '';
    const wantActive = normalized.selection?.slideId;
    this.activeSlideId = wantActive && this.slides.has(wantActive) ? wantActive : firstId;
    this.selection = normalized.selection
      ? { slideId: this.activeSlideId, elementIds: [...normalized.selection.elementIds] }
      : { slideId: this.activeSlideId, elementIds: [] };

    this.emit('deckChange', { keys: ['*'] });
    this.emit('themeChange', { theme: this.theme });
    this.emit('activeSlideChange', { slideId: this.activeSlideId });
    this.emit('selectionChange', { selection: this.getSelection() });
  }
}
