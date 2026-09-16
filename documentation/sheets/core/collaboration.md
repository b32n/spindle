# Sheets Collaboration Internals

This page explains how a spreadsheet is wired for real-time collaboration under the hood. For the user-facing API — `attachCollab`, providers, offline persistence, connection status, and end-to-end encryption — start with the general **[Real-time Collaboration](../../collaboration.md)** guide. Everything here is sheet-specific detail on top of that.

## The binding

Collaboration is built on [Yjs](https://yjs.dev), a CRDT. `WorkbookImpl` keeps a **`Y.Doc`** as the merge-conflict-free shadow of the workbook, and moves opaque bytes to peers through a **`CollabProvider`** — the library never talks to the network directly.

```
WorkbookImpl
     │  workbook.attachCollab(provider, identity, options?)
     ▼
   Y.Doc  ⇄  CollabProvider  ⇄  (network / memory / encryption wrapper)
     │
     ├── 'doc'        channel — durable state (y-protocols/sync)
     └── 'awareness'  channel — ephemeral presence (y-protocols/awareness)
```

Attaching hydrates a fresh `Y.Doc` from the workbook's current `WorkbookData`, subscribes to both channels, and only then calls `provider.connect(roomId)` — so every subscription is in place before the first inbound payload can arrive. The `roomId` defaults to the workbook's `id` and can be overridden via options.

```ts
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';
import { WebSocketProvider } from '@b32nio/spindle-transport-websocket';

const workbook = new WorkbookImpl('wb_1', 'Quarterly Plan');
workbook.setData(savedJson);

const provider = new WebSocketProvider({ url: 'wss://collab.example.com' });

await workbook.attachCollab(
  provider,
  { userId: 'u_42', displayName: 'Bharat', color: '#4ecdc4' },
  { roomId: 'quarterly-plan' },
);

// …later
workbook.detachCollab();
```

`attachCollab` returns a handle (`{ ydoc, awareness, identity, undoManager, detach }`). `<WorkbookCanvas>` picks it up automatically, so remote cursors and selection highlights render with no extra wiring.

## What flows on each channel

**Outbound.** A local edit mutates the workbook's top-level Y types; Yjs emits an update, which is encoded and sent on the `'doc'` channel. Presence changes (the local user's cursor and selection) are encoded from the Yjs `Awareness` instance and sent on `'awareness'`.

**Inbound.**

- `'doc'` payloads are fed through the Yjs sync protocol and applied to the `Y.Doc` with the provider as the update origin.
- `'awareness'` payloads are applied to the `Awareness` instance, updating remote presence.

## Reloading remote changes (v1)

When a remote `'doc'` update lands, the current strategy re-serializes the **whole** `Y.Doc` back to `WorkbookData` and calls the workbook's internal reload path. This is `O(workbook size)` per remote update — simple and always correct. Granular `Y.Map.observe` per type, to touch only the changed cells, is a planned optimization for very large sheets.

## Undo, redo, and presence

- **Undo/redo** routes through a `Y.UndoManager` that tracks every local write on the workbook's top-level Y types (metadata, style/format pools, sheet list, and per-sheet data transitively). Undo therefore produces **inverse Yjs operations that broadcast to peers**, rather than a local snapshot restore that would bypass the shared document.
- **Awareness** carries the `identity` you pass in — `userId`, `displayName`, and `color` — which drive the remote cursor and selection colors.

## Offline persistence

Pass a `persistenceKey` in the options and the `Y.Doc` is mirrored to IndexedDB. On the next attach, prior local state is restored from IndexedDB **before** deciding whether to hydrate from the workbook's data, so edits survive a refresh or an offline period and sync on reconnect. It is browser-only — leave it unset for server-side rendering.

## See also

- **[Real-time Collaboration](../../collaboration.md)** — the full API, transports, connection status, and end-to-end encryption
- **[Architecture](architecture.md)** — the core data model these Y types mirror
