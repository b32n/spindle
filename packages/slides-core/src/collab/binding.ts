// Bridge a deck's Y.Doc to the library's opaque CollabProvider transport, with
// granular observers (no whole-doc reload on the hot path).
//
// Two directions:
//   Mirror   — engine semantic events → Y writes (origin LOCAL_ORIGIN)
//   Observe  — Y changes (remote or undo) → engine _applyRemote* methods
//
// Echo prevention: mirror writes carry LOCAL_ORIGIN; the observers skip
// LOCAL_ORIGIN transactions (the engine already has that state). While applying
// a remote change to the engine, `applyingRemote` is set so the mirror
// listeners don't bounce it back.

import * as Y from 'yjs';
import { ySyncPluginKey } from 'y-prosemirror';
import {
  Awareness,
  applyAwarenessUpdate,
  encodeAwarenessUpdate,
  removeAwarenessStates,
} from 'y-protocols/awareness';
import * as syncProtocol from 'y-protocols/sync';
import * as encoding from 'lib0/encoding';
import * as decoding from 'lib0/decoding';
import { IndexeddbPersistence } from 'y-indexeddb';
import type { CollabIdentity, CollabProvider } from '@b32nio/spindle-shared';

import {
  getDeckYTypes,
  hydrateYDocFromData,
  serializeYDocToData,
  createElementYMap,
  populateElementRichText,
  createSlideYMap,
  yMapToElement,
  yMapToSlide,
  elementProps,
  elementHasRichText,
  toPlain,
  type DeckYTypes,
} from './y-schema';
import type { DeckImpl } from '../deck';
import type { Frame } from '../scene/types';
import type { RichTextDoc } from '../text/model';
import type { SlidesCommentThread } from '../types';

export const LOCAL_ORIGIN = Symbol('spindle-slides-local');

export interface CollabHandle {
  ydoc: Y.Doc;
  awareness: Awareness;
  identity: CollabIdentity;
  undoManager: Y.UndoManager;
  /** The richText fragment for an element (for binding a live y-prosemirror editor). */
  getElementFragment(elementId: string): Y.XmlFragment | undefined;
  detach(): void;
}

export interface AttachCollabOptions {
  roomId?: string;
  /** Mirror the Y.Doc to IndexedDB under this key (browser-only). */
  persistenceKey?: string;
}

const SCALAR_KEYS = new Set(['containerId', 'type', 'x', 'y', 'w', 'h', 'rotation', 'index', 'groupId', 'locked', 'opacity', 'placeholder']);

export async function attachCollabToDeck(
  deck: DeckImpl,
  provider: CollabProvider,
  identity: CollabIdentity,
  options: AttachCollabOptions = {}
): Promise<CollabHandle> {
  const ydoc = new Y.Doc();
  const initialData = deck.getData();

  let persistence: IndexeddbPersistence | undefined;
  if (options.persistenceKey) {
    persistence = new IndexeddbPersistence(options.persistenceKey, ydoc);
    await persistence.whenSynced;
  }

  const y = getDeckYTypes(ydoc);
  // Only locally-restored content might be missing from the relay and need
  // re-contributing on connect; a joiner's replay and a creator's seed already
  // reach it, so gating the re-push avoids bloating the relay's log per attach.
  const restoredWithContent = y.slides.size > 0;

  const awareness = new Awareness(ydoc);
  awareness.setLocalStateField('user', { userId: identity.userId, name: identity.displayName, color: identity.color });

  const undoManager = new Y.UndoManager([y.meta, y.slides, y.elements], {
    trackedOrigins: new Set<unknown>([LOCAL_ORIGIN, ySyncPluginKey]),
    captureTimeout: 500,
  });

  let applyingRemote = false;
  const applyRemote = (fn: () => void) => {
    applyingRemote = true;
    try {
      fn();
    } finally {
      applyingRemote = false;
    }
  };

  // ── Mirror: engine events → Y (origin LOCAL_ORIGIN) ────────────────────────

  const tx = (fn: () => void) => ydoc.transact(fn, LOCAL_ORIGIN);

  const offs: Array<() => void> = [];
  const on = (event: Parameters<DeckImpl['on']>[0], handler: (payload: unknown) => void) => {
    offs.push(deck.on(event, (e) => {
      if (applyingRemote) return;
      handler(e.payload);
    }));
  };

  // After a local mirror write, re-read the element back from Y and replace the
  // engine's record with it, so the writer holds the same canonical form peers
  // derive (notably: rich text picks up the PM default paragraph attrs the
  // fragment round-trip adds). Guarded so the resulting event isn't re-mirrored.
  const reconcileLocal = (id: string) => {
    const m = y.elements.get(id);
    if (m) applyRemote(() => deck._applyRemoteElementUpsert(yMapToElement(id, m)));
  };

  const writeElement = (id: string) => {
    const el = deck.getElement(id);
    if (!el) return;
    const m = createElementYMap(el);
    y.elements.set(id, m);
    if (elementHasRichText(el)) populateElementRichText(m, (el as { richText?: RichTextDoc }).richText);
    reconcileLocal(id);
  };

  const updateElementKeys = (id: string, keys: string[]) => {
    const m = y.elements.get(id);
    const el = deck.getElement(id);
    if (!m || !el) return;
    let propsDirty = false;
    let touchedRichText = false;
    for (const k of keys) {
      if (k === '*') {
        writeElement(id);
        return;
      }
      if (SCALAR_KEYS.has(k)) {
        const v = (el as unknown as Record<string, unknown>)[k];
        if (v === undefined) m.delete(k);
        else m.set(k, typeof v === 'object' ? toPlain(v) : v);
      } else if (k === 'richText') {
        touchedRichText = true;
        const rt = (el as { richText?: RichTextDoc }).richText;
        let frag = m.get('richText') as Y.XmlFragment | undefined;
        if (frag) frag.delete(0, frag.length);
        else if (rt) {
          frag = new Y.XmlFragment();
          m.set('richText', frag);
        }
        if (rt) populateElementRichText(m, rt);
      } else {
        propsDirty = true;
      }
    }
    if (propsDirty) m.set('props', toPlain(elementProps(el)));
    if (touchedRichText) reconcileLocal(id);
  };

  const attachMirror = () => {
    on('elementAdd', (p) => tx(() => writeElement((p as { elementId: string }).elementId)));
    on('elementChange', (p) => {
      const { elementId, keys } = p as { elementId: string; keys?: string[] };
      tx(() => updateElementKeys(elementId, keys ?? ['*']));
    });
    on('elementDelete', (p) => tx(() => y.elements.delete((p as { elementId: string }).elementId)));

    on('slideAdd', (p) => {
      const slide = deck.getSlide((p as { slideId: string }).slideId);
      if (slide) tx(() => y.slides.set(slide.id, createSlideYMap(slide)));
    });
    on('slideDelete', (p) => tx(() => y.slides.delete((p as { slideId: string }).slideId)));
    on('slideMove', (p) => {
      const slide = deck.getSlide((p as { slideId: string }).slideId);
      if (slide) tx(() => (y.slides.get(slide.id) as Y.Map<unknown>)?.set('index', slide.index));
    });
    on('slideChange', (p) => {
      const { slideId } = p as { slideId: string };
      const slide = deck.getSlide(slideId);
      const m = y.slides.get(slideId) as Y.Map<unknown> | undefined;
      if (!slide || !m) return;
      tx(() => {
        m.set('index', slide.index);
        if (slide.layoutRef) m.set('layoutRef', slide.layoutRef); else m.delete('layoutRef');
        if (slide.background) m.set('background', toPlain(slide.background)); else m.delete('background');
        if (slide.notes) m.set('notes', toPlain(slide.notes)); else m.delete('notes');
      });
    });

    on('deckChange', () => tx(() => {
      y.meta.set('id', deck.id);
      y.meta.set('title', deck.getTitle());
      y.meta.set('slideSize', toPlain(deck.getSlideSize()));
      y.meta.set('layouts', toPlain(deck.getLayouts()));
    }));
    on('themeChange', () => tx(() => y.meta.set('theme', toPlain(deck.getTheme()))));

    // Comments live in a top-level Y.Map outside the undo scope (not undoable,
    // matching sheets). Full-sync on each change — thread volume is small.
    on('commentChange', () => tx(() => {
      const current = new Set<string>();
      for (const t of deck.getComments().getThreads()) {
        current.add(t.id);
        y.threads.set(t.id, toPlain(t));
      }
      for (const id of [...y.threads.keys()]) if (!current.has(id)) y.threads.delete(id);
    }));
  };

  // ── Observe: Y changes (remote / undo) → engine ────────────────────────────

  const onElements = (events: Y.YEvent<Y.AbstractType<unknown>>[], transaction: Y.Transaction) => {
    if (transaction.origin === LOCAL_ORIGIN) return;
    const affected = new Set<string>();
    for (const e of events) {
      if (e.target === y.elements) {
        for (const key of (e as Y.YMapEvent<unknown>).keysChanged) affected.add(key);
      } else if (typeof e.path[0] === 'string') {
        affected.add(e.path[0] as string);
      }
    }
    applyRemote(() => {
      for (const id of affected) {
        const m = y.elements.get(id);
        if (m) deck._applyRemoteElementUpsert(yMapToElement(id, m));
        else deck._applyRemoteElementDelete(id);
      }
    });
  };

  const onSlides = (events: Y.YEvent<Y.AbstractType<unknown>>[], transaction: Y.Transaction) => {
    if (transaction.origin === LOCAL_ORIGIN) return;
    const affected = new Set<string>();
    for (const e of events) {
      if (e.target === y.slides) {
        for (const key of (e as Y.YMapEvent<unknown>).keysChanged) affected.add(key);
      } else if (typeof e.path[0] === 'string') {
        affected.add(e.path[0] as string);
      }
    }
    applyRemote(() => {
      for (const id of affected) {
        const m = y.slides.get(id);
        if (m) deck._applyRemoteSlideUpsert(yMapToSlide(id, m));
        else deck._applyRemoteSlideDelete(id);
      }
    });
  };

  const onMeta = (_e: Y.YMapEvent<unknown>, transaction: Y.Transaction) => {
    if (transaction.origin === LOCAL_ORIGIN) return;
    applyRemote(() => deck._applyRemoteMeta({
      title: y.meta.get('title') as string | undefined,
      slideSize: y.meta.get('slideSize') as Frame | undefined,
      theme: y.meta.get('theme') as never,
      layouts: y.meta.get('layouts') as never,
    }));
  };

  const onThreads = (_e: Y.YMapEvent<SlidesCommentThread>, transaction: Y.Transaction) => {
    if (transaction.origin === LOCAL_ORIGIN) return;
    applyRemote(() => deck._applyRemoteComments([...y.threads.values()]));
  };

  const attachObservers = () => {
    y.elements.observeDeep(onElements);
    y.slides.observeDeep(onSlides);
    y.meta.observe(onMeta);
    y.threads.observe(onThreads);
  };

  // ── Wire transport (identical shape to sheets/docs) ────────────────────────

  const roomId = options.roomId ?? initialData.id;

  const onDocUpdate = (update: Uint8Array, origin: unknown) => {
    if (origin !== provider) {
      const encoder = encoding.createEncoder();
      syncProtocol.writeUpdate(encoder, update);
      provider.send('doc', encoding.toUint8Array(encoder));
    }
  };
  ydoc.on('update', onDocUpdate);

  const onAwarenessUpdate = (changes: { added: number[]; updated: number[]; removed: number[] }, origin: unknown) => {
    if (origin === provider) return;
    const clients = changes.added.concat(changes.updated, changes.removed);
    if (clients.length === 0) return;
    provider.send('awareness', encodeAwarenessUpdate(awareness, clients));
  };
  awareness.on('update', onAwarenessUpdate);

  const unsubDoc = provider.onMessage('doc', (payload) => {
    const decoder = decoding.createDecoder(payload);
    const reply = encoding.createEncoder();
    syncProtocol.readSyncMessage(decoder, reply, ydoc, provider);
    if (encoding.length(reply) > 0) provider.send('doc', encoding.toUint8Array(reply));
  });
  const unsubAwareness = provider.onMessage('awareness', (payload) => {
    applyAwarenessUpdate(awareness, payload, provider);
  });

  await provider.connect(roomId);

  // connect() replays the room's existing state before it resolves. Seed only
  // if never seeded (a persistent marker) AND the room is empty — a joiner (or
  // a joiner to an emptied room, or a restored doc) must not stack a duplicate
  // or resurrect deleted content. Then reconcile the engine to exactly match
  // the Y.Doc, and only now attach the ongoing mirror + observers.
  const seedMeta = ydoc.getMap('__spindle');
  if (seedMeta.get('seeded') !== true && y.slides.size === 0) {
    ydoc.transact(() => {
      hydrateYDocFromData(ydoc, initialData);
      seedMeta.set('seeded', true);
    });
  }
  deck._resyncFromY(serializeYDocToData(ydoc, deck.getSelection(), deck.getActiveSlideId()));
  attachObservers();
  attachMirror();

  // Contribute locally-restored state so a relay that never saw it gets it.
  // Idempotent (Yjs dedupes).
  if (restoredWithContent && y.slides.size > 0) {
    const stateEncoder = encoding.createEncoder();
    syncProtocol.writeUpdate(stateEncoder, Y.encodeStateAsUpdate(ydoc));
    provider.send('doc', encoding.toUint8Array(stateEncoder));
  }

  // Broadcast presence, and re-broadcast on reconnect so a blip doesn't drop
  // our cursor for peers.
  const sendAwareness = (): void =>
    provider.send('awareness', encodeAwarenessUpdate(awareness, [ydoc.clientID]));
  sendAwareness();
  const unsubStatus = provider.onStatusChange?.((status) => {
    if (status === 'connected') sendAwareness();
  });

  const detach = () => {
    removeAwarenessStates(awareness, [ydoc.clientID], 'detach');
    ydoc.off('update', onDocUpdate);
    awareness.off('update', onAwarenessUpdate);
    y.elements.unobserveDeep(onElements);
    y.slides.unobserveDeep(onSlides);
    y.meta.unobserve(onMeta);
    y.threads.unobserve(onThreads);
    for (const off of offs) off();
    unsubStatus?.();
    unsubDoc();
    unsubAwareness();
    undoManager.destroy();
    awareness.destroy();
    provider.disconnect();
    void persistence?.destroy();
    ydoc.destroy();
  };

  return {
    ydoc,
    awareness,
    identity,
    undoManager,
    getElementFragment: (id) => (y.elements.get(id) as Y.Map<unknown> | undefined)?.get('richText') as Y.XmlFragment | undefined,
    detach,
  };
}

export type { DeckYTypes };
