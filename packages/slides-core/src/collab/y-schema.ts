// Y.Doc schema for collaborative decks.
//
// Layout:
//   meta:     Y.Map — id, title, slideSize(JSON), theme(JSON), layouts(JSON)   (LWW)
//   slides:   Y.Map<slideId, Y.Map{ index, layoutRef?, background?(JSON), notes?(JSON), hidden? }>
//   elements: Y.Map<elId, Y.Map{ containerId, type, x, y, w, h, rotation, index,
//                                groupId?, parentGroupId?, locked?, opacity?, placeholder?(JSON),
//                                props(JSON), richText?: Y.XmlFragment }>
//   threads:  Y.Map<threadId, JSON>   — outside undo scope (Phase 6)
//   opaqueParts: Y.Map<hash, Uint8Array> — content-addressed parts kept byte for byte (schemaVersion 10, plan K11); the
//                value is `encodeOpaqueEntry` (content type + bytes); outside undo scope (parts are only ever added)
//   (a slide map also carries `opaque` JSON: the transition and timing XML of an imported slide, schemaVersion 10)
//
// schemaVersion (meta): the version of this layout the room was written with (missing = 1, the layout above).
// attachCollabToDeck refuses a room whose version is newer than COLLAB_SCHEMA_VERSION (see assertSchemaVersion), so an
// older client never re-saves a room and drops keys it does not know.
//
// Frame fields are individual Y.Map keys so concurrent move/rotate/z-reorder on
// one element merge per-key. Type-specific styling lives in one `props` JSON
// key (LWW, v1-acceptable). The rich-text body is a Y.XmlFragment bound via
// y-prosemirror. Fractional `index` + jitter → concurrent reorders converge.

import * as Y from 'yjs';
import { encodeOpaqueEntry, type OpaquePart } from '@b32nio/spindle-shared/opaque';
import { prosemirrorToYXmlFragment, yXmlFragmentToProseMirrorRootNode } from 'y-prosemirror';
import { slidesSchema } from '../text/schema';
import type { SlideElement } from '../scene/types';
import type { RichTextDoc } from '../text/model';
import type { DeckData, Slide, SlideData, SlidesCommentThread, DeckSelection } from '../types';

/**
 * The Y layout version this client reads and writes. 1 is the layout described above (rooms written before the key
 * existed have no `schemaVersion` and read as 1). Raised by any later change an older client could not read.
 *
 * 2: paragraph nodes of the rich-text fragment carry K3b attributes (`bullet`, `lineHeightPx`, `spaceBeforePct`,
 * `spaceAfterPct`, `marginLeft`, `firstLineIndent`, `tabStops`) as XmlElement attributes. A v1 client's schema does not
 * know them and would drop them on its next write, so it must not join a v2 room.
 *
 * 3: the rich-text fragment carries K3c run marks (`baseline`, `caps`, `charSpacing`, `lang`, as text formatting
 * attributes) and `field` nodes (slide number, date/time, footer; an XmlElement with `kind` and `format`
 * attributes whose child text is the cached text). A v2 client's schema has neither and would drop them on its next write, so it must not join a v3 room.
 *
 * 4: K4 layouts and slide visibility. `meta.layouts` entries carry `elements` (static layout content), `textStyles` and
 * `masterName`, and slide maps carry `hidden`. A v3 client would show hidden slides in its presentation and drop the
 * `hidden` key the next time it rewrites a slide, so it must not join a v4 room.
 *
 * 5: K5 shape presets. A shape element's `props.shape` may be one of the ~58 added presets or `unknown:<prst>` (an
 * OOXML preset the editor cannot draw, kept for re-export), and `props.adjustments` carries per-preset guides. A v4
 * client would draw those shapes as plain rectangles and offer no way to see what they are, so it must not join a v5 room.
 *
 * 6: K6 nested groups. An element map carries `parentGroupId` (the group containing its `groupId`) next to `groupId`.
 * A v5 client does not know the key, would read the inner group as top level (selecting and moving it apart from its
 * outer group) and drop the key on its next write of that element, flattening the nesting, so it must not join a v6 room.
 *
 * 7: K6b elbow and curved connectors. A line element's `props` carry `connector` ('elbow' | 'curved') and `adjustments`
 * (the elbow's middle-leg position), and a bound anchor on an ellipse, diamond or triangle sits on the outline rather than
 * the box. A v6 client would draw these as straight lines to the wrong points, so it must not join a v7 room.
 *
 * 8: K6c custom geometry. A shape element's `props.shape` may be `path`, and `props.path` carries the `CustomGeometry`
 * (paths of move/line/cubic/quad/arc/close commands, fill modes). A v7 client would draw such a shape as a plain
 * rectangle and, on its next write of the element, keep the key but could not edit or re-export it faithfully, so it
 * must not join a v8 room.
 *
 * 9: K7 tables. A table element's `props.cells` cells may carry `gridSpan`, `rowSpan`, `hMerge`, `vMerge` (merged
 * regions) and `borders` (per-side strokes), and the table `props.styleFlags`. A v8 client would draw a merged table as
 * a plain grid (covered cells shown), and its next edit of the cells would keep the keys while ignoring the spans on
 * insert/delete, so it must not join a v9 room. Table cell text is part of the same `props` JSON (last writer wins,
 * not CRDT-merged); that is an existing limit, unchanged here.
 *
 * 10: K11 opaque content. An element map may have `type: 'opaque'` (a chart, SmartArt diagram, OLE object, clip or ink
 * kept but not editable, `props` carrying `kind`, `partRefs`, ...), element `props` may carry `pptxId`, a slide map
 * carries `opaque` (transition and timing XML; the timing names shapes by `pptxId` and is pruned when a shape is deleted),
 * and the document has a top-level `opaqueParts` map (hash to part). A v9 client would draw the element as nothing,
 * keep `type: 'opaque'` it cannot render, and on deleting a shape leave the slide's timing pointing at it, so it must not
 * join a v10 room.
 */
export const COLLAB_SCHEMA_VERSION = 10;

/** Thrown when a room was written with a newer Y deck layout than this client understands. */
export class CollabSchemaVersionError extends Error {
  readonly roomVersion: number;
  readonly supportedVersion: number;
  constructor(roomVersion: number, supportedVersion: number = COLLAB_SCHEMA_VERSION) {
    super(
      `This collaboration room uses deck schema version ${roomVersion}, but this editor only supports up to ` +
        `version ${supportedVersion}. Update the editor to join the room.`,
    );
    this.name = 'CollabSchemaVersionError';
    this.roomVersion = roomVersion;
    this.supportedVersion = supportedVersion;
  }
}

/** The schema version a Y.Doc was written with (1 when the key is absent or not a number). */
export function readSchemaVersion(ydoc: Y.Doc): number {
  const v = ydoc.getMap<unknown>('meta').get('schemaVersion');
  return typeof v === 'number' && Number.isFinite(v) && v >= 1 ? v : 1;
}

/** Throw a CollabSchemaVersionError when the Y.Doc was written with a newer schema than this client supports. */
export function assertSchemaVersion(ydoc: Y.Doc, supported: number = COLLAB_SCHEMA_VERSION): void {
  const version = readSchemaVersion(ydoc);
  if (version > supported) throw new CollabSchemaVersionError(version, supported);
}

export interface DeckYTypes {
  meta: Y.Map<unknown>;
  slides: Y.Map<Y.Map<unknown>>;
  elements: Y.Map<Y.Map<unknown>>;
  threads: Y.Map<SlidesCommentThread>;
  /** hash → `encodeOpaqueEntry` of the part (v10). */
  opaqueParts: Y.Map<Uint8Array>;
}

/**
 * Normalize a value to a plain, same-realm JSON structure before storing it in
 * Yjs. Yjs's content check is strict (`value.constructor === Object`), so a
 * cross-realm object — e.g. structuredClone output under some runtimes — is
 * rejected as "Unexpected content type". A JSON round-trip guarantees plain
 * objects/arrays. Only used for the small metadata/props blobs, never for the
 * hot-path primitive frame keys.
 */
export function toPlain<T>(value: T): T {
  return value === undefined ? value : (JSON.parse(JSON.stringify(value)) as T);
}

export function getDeckYTypes(ydoc: Y.Doc): DeckYTypes {
  return {
    meta: ydoc.getMap<unknown>('meta'),
    slides: ydoc.getMap<Y.Map<unknown>>('slides'),
    elements: ydoc.getMap<Y.Map<unknown>>('elements'),
    threads: ydoc.getMap<SlidesCommentThread>('threads'),
    opaqueParts: ydoc.getMap<Uint8Array>('opaqueParts'),
  };
}

// Element fields stored as individual Y.Map keys (per-key CRDT merge).
const SCALAR_KEYS = ['containerId', 'type', 'x', 'y', 'w', 'h', 'rotation', 'index', 'groupId', 'parentGroupId', 'locked', 'opacity', 'placeholder'] as const;

/** The type-specific styling fields bundled into the `props` JSON key. */
export function elementProps(el: SlideElement): Record<string, unknown> {
  const rest: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(el)) {
    if ((SCALAR_KEYS as readonly string[]).includes(k)) continue;
    if (k === 'id' || k === 'richText') continue;
    rest[k] = v;
  }
  return rest;
}

/** True when an element actually carries a rich-text body (not just a type that could). */
export function elementHasRichText(el: SlideElement): boolean {
  return (el.type === 'text' || el.type === 'shape') && !!(el as { richText?: unknown }).richText;
}

/**
 * Build (and attach) a Y.Map for an element. The richText fragment is created
 * empty here; populate it with `populateElementRichText` AFTER the map is added
 * to the elements map (a fragment must be integrated into the doc first).
 */
export function createElementYMap(el: SlideElement): Y.Map<unknown> {
  const m = new Y.Map<unknown>();
  for (const k of SCALAR_KEYS) {
    const v = (el as unknown as Record<string, unknown>)[k];
    if (v !== undefined) m.set(k, typeof v === 'object' ? toPlain(v) : v);
  }
  m.set('props', toPlain(elementProps(el)));
  if (elementHasRichText(el)) m.set('richText', new Y.XmlFragment());
  return m;
}

/** Populate an element's richText fragment from JSON (after the map is attached). */
export function populateElementRichText(m: Y.Map<unknown>, richText: RichTextDoc | undefined): void {
  const frag = m.get('richText') as Y.XmlFragment | undefined;
  // Yjs only stores attribute values whose constructor is exactly Object/Array (object-valued paragraph attrs such as
  // `bullet`), so rebuild the JSON from plain literals: a structuredClone'd value (other realm) or a null-prototype one
  // from an importer would otherwise throw "Unexpected content type".
  if (frag && richText) prosemirrorToYXmlFragment(slidesSchema.nodeFromJSON(JSON.parse(JSON.stringify(richText))), frag);
}

/** Read an element's richText fragment back to JSON. */
export function fragmentToRichText(frag: Y.XmlFragment): RichTextDoc {
  return yXmlFragmentToProseMirrorRootNode(frag, slidesSchema).toJSON() as RichTextDoc;
}

/** Reconstruct a SlideElement from its Y.Map (id is the elements-map key). */
export function yMapToElement(id: string, m: Y.Map<unknown>): SlideElement {
  const el: Record<string, unknown> = { id };
  for (const k of SCALAR_KEYS) {
    if (m.has(k)) el[k] = m.get(k);
  }
  // Copy peer-controlled props, skipping keys that would poison el's prototype
  // (Object.assign's [[Set]] would honor a __proto__ key).
  const props = (m.get('props') as Record<string, unknown>) ?? {};
  for (const [k, v] of Object.entries(props)) {
    if (k === '__proto__' || k === 'constructor' || k === 'prototype') continue;
    el[k] = v;
  }
  const frag = m.get('richText') as Y.XmlFragment | undefined;
  if (frag) el.richText = fragmentToRichText(frag);
  return el as unknown as SlideElement;
}

export function createSlideYMap(slide: Slide): Y.Map<unknown> {
  const m = new Y.Map<unknown>();
  m.set('index', slide.index);
  if (slide.layoutRef) m.set('layoutRef', slide.layoutRef);
  if (slide.background) m.set('background', toPlain(slide.background));
  if (slide.notes) m.set('notes', toPlain(slide.notes));
  if (slide.hidden) m.set('hidden', true);
  if (slide.opaque) m.set('opaque', toPlain(slide.opaque));
  return m;
}

export function yMapToSlide(id: string, m: Y.Map<unknown>): Slide {
  const slide: Slide = { id, index: m.get('index') as string };
  if (m.has('layoutRef')) slide.layoutRef = m.get('layoutRef') as string;
  if (m.has('background')) slide.background = m.get('background') as Slide['background'];
  if (m.has('notes')) slide.notes = m.get('notes') as RichTextDoc;
  if (m.get('hidden') === true) slide.hidden = true;
  if (m.has('opaque')) slide.opaque = m.get('opaque') as Slide['opaque'];
  return slide;
}

/** First-time hydration of an empty Y.Doc from a DeckData snapshot. */
export function hydrateYDocFromData(ydoc: Y.Doc, data: DeckData): void {
  const y = getDeckYTypes(ydoc);
  ydoc.transact(() => {
    y.meta.set('schemaVersion', COLLAB_SCHEMA_VERSION);
    y.meta.set('id', data.id);
    y.meta.set('title', data.title);
    y.meta.set('slideSize', toPlain(data.slideSize));
    y.meta.set('theme', toPlain(data.theme));
    if (data.layouts) y.meta.set('layouts', toPlain(data.layouts));

    for (const slideData of data.slides) {
      const { elements, ...slide } = slideData;
      y.slides.set(slide.id, createSlideYMap(slide));
      for (const el of elements) {
        const m = createElementYMap(el);
        y.elements.set(el.id, m);
        if (elementHasRichText(el)) populateElementRichText(m, (el as { richText?: RichTextDoc }).richText);
      }
    }

    if (data.threads) for (const t of data.threads) y.threads.set(t.id, toPlain(t));
    for (const part of data.opaqueParts ?? []) y.opaqueParts.set(part.hash, encodeOpaqueEntry(part));
  });
}

/** Read the Y.Doc back to a DeckData object (local view state supplied by caller). */
export function serializeYDocToData(
  ydoc: Y.Doc,
  selection: DeckSelection,
  activeSlideId: string,
  opaqueParts?: OpaquePart[]
): DeckData {
  const y = getDeckYTypes(ydoc);

  // Group elements by slide.
  const bySlide = new Map<string, SlideElement[]>();
  for (const [id, m] of y.elements.entries()) {
    const el = yMapToElement(id, m);
    const list = bySlide.get(el.containerId) ?? [];
    list.push(el);
    bySlide.set(el.containerId, list);
  }

  const slides: SlideData[] = [];
  for (const [id, m] of y.slides.entries()) {
    const slide = yMapToSlide(id, m);
    const els = (bySlide.get(id) ?? []).sort((a, b) =>
      a.index < b.index ? -1 : a.index > b.index ? 1 : a.id < b.id ? -1 : 1
    );
    slides.push({ ...slide, elements: els });
  }
  slides.sort((a, b) => (a.index < b.index ? -1 : a.index > b.index ? 1 : a.id < b.id ? -1 : 1));

  const threads: SlidesCommentThread[] = [];
  for (const t of y.threads.values()) threads.push(t);

  return {
    id: (y.meta.get('id') as string) ?? '',
    title: (y.meta.get('title') as string) ?? '',
    slideSize: (y.meta.get('slideSize') as DeckData['slideSize']) ?? { w: 1280, h: 720 },
    theme: y.meta.get('theme') as DeckData['theme'],
    layouts: y.meta.get('layouts') as DeckData['layouts'],
    slides,
    threads: threads.length ? threads : undefined,
    opaqueParts: opaqueParts && opaqueParts.length > 0 ? opaqueParts : undefined,
    selection: { slideId: activeSlideId, elementIds: [...selection.elementIds] },
  };
}
