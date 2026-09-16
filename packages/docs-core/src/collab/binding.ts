// Bridge a Y.Doc to the library's opaque CollabProvider transport.
//
// Wire protocol per channel:
//   'doc'        — y-protocols/sync messages (state vector exchange + updates).
//                  Both the initial sync handshake and incremental edits travel
//                  here; readSyncMessage decides what to do with each payload.
//   'awareness'  — y-protocols/awareness encoded updates. Ephemeral; not stored.
//
// Outbound flow:
//   ydoc.on('update')          → writeUpdate → provider.send('doc', bytes)
//   awareness.on('update')     → encodeAwarenessUpdate → provider.send('awareness', bytes)
//
// Inbound flow:
//   provider.onMessage('doc')        → readSyncMessage → maybe-reply on 'doc'
//   provider.onMessage('awareness')  → applyAwarenessUpdate
//
// On connect, we proactively send a SyncStep1 so any existing peer responds
// with the updates we're missing — that's how a late joiner converges.

import * as Y from 'yjs';
import { Awareness, applyAwarenessUpdate, encodeAwarenessUpdate, removeAwarenessStates } from 'y-protocols/awareness';
import * as syncProtocol from 'y-protocols/sync';
import * as encoding from 'lib0/encoding';
import * as decoding from 'lib0/decoding';
import { IndexeddbPersistence } from 'y-indexeddb';
import type { CollabIdentity, CollabProvider } from '@b32nio/spindle-shared';

import { getYDocFields, hydrateYDocFromData } from './y-schema';
import type { DocumentData } from '../types';

/**
 * Live handle returned by attachCollab. The React layer reads `xmlFragment`
 * for ySyncPlugin and `awareness` for yCursorPlugin. `detach()` releases
 * every handler and tears the Y.Doc down.
 */
export interface CollabHandle {
  ydoc: Y.Doc;
  xmlFragment: Y.XmlFragment;
  awareness: Awareness;
  identity: CollabIdentity;
  detach(): void;
}

export interface AttachCollabOptions {
  /**
   * Optional roomId override. Defaults to the document's meta.id once the
   * Y.Doc is hydrated. Consumers can pass a different room name to share
   * sessions under a stable URL even if the document id changes.
   */
  roomId?: string;
  /**
   * When set, the Y.Doc is mirrored to IndexedDB under this key. Edits
   * survive a page refresh / offline period: on the next attach the Y.Doc
   * is restored from IndexedDB BEFORE deciding whether to hydrate from
   * initialData. Browser-only — has no effect (and would throw) outside a
   * browser, so leave it unset for server-side rendering.
   */
  persistenceKey?: string;
}

/**
 * Hydrate a Y.Doc from the given DocumentData, attach it to the provider,
 * and return a live handle. Caller must call detach() when done.
 *
 * The provider is expected to NOT be connected yet — attachCollab calls
 * provider.connect(roomId) as the final step so subscriptions are all in
 * place before any inbound payload can arrive.
 */
export async function attachCollabToYDoc(
  initialData: DocumentData,
  provider: CollabProvider,
  identity: CollabIdentity,
  options: AttachCollabOptions = {},
): Promise<CollabHandle> {
  const ydoc = new Y.Doc();

  // Optional IndexedDB persistence. Load any prior local state FIRST, then
  // only hydrate from initialData if the doc came back empty — otherwise
  // we'd append a duplicate copy of the content on top of the restored one.
  let persistence: IndexeddbPersistence | undefined;
  if (options.persistenceKey) {
    persistence = new IndexeddbPersistence(options.persistenceKey, ydoc);
    await persistence.whenSynced;
  }
  // Only content restored from local persistence (applied before onDocUpdate is
  // wired) might be missing from the relay and need re-contributing on connect.
  // A joiner's replayed state and a creator's seed already reach the relay, so
  // they must not re-push (that would bloat the relay's log on every attach).
  const restoredWithContent = getYDocFields(ydoc).content.length > 0;

  // Stable reference to the top-level fragment. It exists whether or not the
  // doc has content yet; connect() (below) fills it in for a joiner, and the
  // creator's seed fills it in after connect.
  const { content: xmlFragment } = getYDocFields(ydoc);

  // Awareness carries cursor / presence state. Identity goes on local state.
  const awareness = new Awareness(ydoc);
  awareness.setLocalStateField('user', {
    userId: identity.userId,
    name: identity.displayName,
    color: identity.color,
  });

  const roomId = options.roomId ?? initialData.id;

  // --- Outbound: local edits → wire ----------------------------------------

  const onDocUpdate = (update: Uint8Array, origin: unknown): void => {
    // Don't echo updates that came from this very provider — readSyncMessage
    // tags applies with `provider` as origin, so this filter prevents loops.
    if (origin === provider) return;
    const encoder = encoding.createEncoder();
    syncProtocol.writeUpdate(encoder, update);
    provider.send('doc', encoding.toUint8Array(encoder));
  };
  ydoc.on('update', onDocUpdate);

  const onAwarenessUpdate = (
    changes: { added: number[]; updated: number[]; removed: number[] },
    origin: unknown,
  ): void => {
    if (origin === provider) return;
    const changedClients = changes.added.concat(changes.updated, changes.removed);
    if (changedClients.length === 0) return;
    const payload = encodeAwarenessUpdate(awareness, changedClients);
    provider.send('awareness', payload);
  };
  awareness.on('update', onAwarenessUpdate);

  // --- Inbound: wire → local Y types ----------------------------------------

  const unsubDoc = provider.onMessage('doc', (payload: Uint8Array) => {
    const decoder = decoding.createDecoder(payload);
    const reply = encoding.createEncoder();
    // readSyncMessage handles all three message types (SyncStep1, SyncStep2,
    // Update). When it writes into `reply`, we send that back to whoever sent
    // us this message — that's how the SyncStep1/SyncStep2 handshake closes.
    syncProtocol.readSyncMessage(decoder, reply, ydoc, provider);
    if (encoding.length(reply) > 0) {
      provider.send('doc', encoding.toUint8Array(reply));
    }
  });

  const unsubAwareness = provider.onMessage('awareness', (payload: Uint8Array) => {
    applyAwarenessUpdate(awareness, payload, provider);
  });

  // --- Connect + seed --------------------------------------------------------

  // connect() replays the room's existing state before it resolves, so once it
  // returns the fragment already holds whatever the room contains. Seed from
  // initialData only if it's still empty — i.e. we're the room's creator. A
  // joiner (or a reconnecting peer whose state was restored) seeds nothing,
  // which prevents duplicated content in the common case. (Two peers creating
  // the *same* brand-new room within one round-trip can still both seed — that
  // needs relay-side arbitration, which an opaque relay can't provide.)
  await provider.connect(roomId);

  // Seed only if the room has never been seeded (a persistent marker) AND the
  // doc is still empty. The marker means a joiner to a room whose content was
  // legitimately deleted won't resurrect it, and a restored/pre-existing doc
  // isn't re-seeded.
  const seedMeta = ydoc.getMap('__spindle');
  if (seedMeta.get('seeded') !== true && getYDocFields(ydoc).content.length === 0) {
    ydoc.transact(() => {
      hydrateYDocFromData(ydoc, initialData);
      seedMeta.set('seeded', true);
    });
  }

  // Contribute locally-restored state so a relay that never saw it (e.g. a
  // doc created offline and synced later) gets it. Idempotent (Yjs dedupes).
  if (restoredWithContent && getYDocFields(ydoc).content.length > 0) {
    const stateEncoder = encoding.createEncoder();
    syncProtocol.writeUpdate(stateEncoder, Y.encodeStateAsUpdate(ydoc));
    provider.send('doc', encoding.toUint8Array(stateEncoder));
  }

  // Broadcast our awareness state so existing peers see us immediately, and
  // re-broadcast on every reconnect so peers don't lose our cursor after a blip.
  const sendAwareness = (): void =>
    provider.send('awareness', encodeAwarenessUpdate(awareness, [ydoc.clientID]));
  sendAwareness();
  const unsubStatus = provider.onStatusChange?.((status) => {
    if (status === 'connected') sendAwareness();
  });

  // --- Teardown -------------------------------------------------------------

  const detach = (): void => {
    // Tell peers we're leaving so their cursor overlays clean up.
    removeAwarenessStates(awareness, [ydoc.clientID], 'detach');
    ydoc.off('update', onDocUpdate);
    awareness.off('update', onAwarenessUpdate);
    unsubStatus?.();
    unsubDoc();
    unsubAwareness();
    awareness.destroy();
    provider.disconnect();
    void persistence?.destroy();
    ydoc.destroy();
  };

  return {
    ydoc,
    xmlFragment,
    awareness,
    identity,
    detach,
  };
}
