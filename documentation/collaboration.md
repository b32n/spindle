# Real-Time Collaboration

Spindle documents, spreadsheets, and slide decks can sync live between users.
The collaboration layer is built on [Yjs](https://yjs.dev) (a CRDT), but the
library never exposes Yjs to you directly — you work with a small
transport interface and an `attachCollab` call.

## Design in one paragraph

Every collaborative session has a **Y.Doc** (the merge-conflict-free shared
state) and a **`CollabProvider`** (a transport that moves opaque bytes
between peers). The library owns the Y.Doc; you supply the provider. The
provider never inspects payloads — which is what makes encryption a clean
wrapping concern (see [End-to-end encryption](#end-to-end-encryption)).

```
DocumentImpl / WorkbookImpl
        │  attachCollab(provider, identity, options)
        ▼
   Y.Doc  ⇄  CollabProvider  ⇄  (network / memory / encryption wrapper)
        │
        ├── 'doc'        channel — durable state (Yjs sync protocol)
        └── 'awareness'  channel — ephemeral presence (cursors, selections)
```

## Quick start

```ts
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';
import { WebSocketProvider } from '@b32nio/spindle-transport-websocket';

const workbook = new WorkbookImpl('wb_1', 'Quarterly Plan');
workbook.setData(savedJson); // your persisted WorkbookData

const provider = new WebSocketProvider({ url: 'wss://collab.example.com' });

await workbook.attachCollab(
  provider,
  { userId: 'u_42', displayName: 'Bharat', color: '#4ecdc4' },
  { roomId: 'quarterly-plan' },
);
```

Documents are identical — `new DocumentImpl()`, `document.setData(...)`,
`document.attachCollab(provider, identity, options)`.

`attachCollab` returns a **handle** (`{ ydoc, awareness, undoManager,
detach, ... }`). Call `detach()` when the session ends. The React editors
(`<WorkbookCanvas>`, `<DocumentEditor>`) pick the handle up automatically
via `getCollabHandle()` — remote cursors and selections render with no
extra wiring.

## The `CollabProvider` interface

```ts
interface CollabProvider {
  connect(roomId: string): Promise<void>;
  disconnect(): void;
  send(channel: 'doc' | 'awareness', payload: Uint8Array): void;
  onMessage(channel: 'doc' | 'awareness', handler): () => void;
  // optional:
  getStatus?(): 'connecting' | 'connected' | 'offline';
  onStatusChange?(handler): () => void;
}
```

That is the entire contract. Anything that can shuttle `Uint8Array`s
between peers can be a provider.

### Provided transports

| Provider | Package | Use |
|---|---|---|
| `InMemoryProvider` | `@b32nio/spindle-shared` | Tests / in-process demos. Rooms are a module-level registry. |
| `WebSocketProvider` | `@b32nio/spindle-transport-websocket` | Real cross-tab / cross-machine sync via a relay server. |

A reference relay server lives in `examples/collab-server` — a dumb
per-room broadcast hub. **It has no auth or persistence by design.** A
production deployment supplies those (see that folder's README).

## Offline persistence

Pass `persistenceKey` to mirror the Y.Doc into the browser's IndexedDB:

```ts
await workbook.attachCollab(provider, identity, {
  roomId: 'quarterly-plan',
  persistenceKey: 'wb:quarterly-plan',
});
```

On the next attach the Y.Doc is restored from IndexedDB *before* deciding
whether to hydrate from `setData` — so edits survive a refresh or an
offline period, and reconcile with peers on reconnect. Browser-only;
leave it unset for server-side rendering.

## Connection status

```ts
const off = provider.onStatusChange?.((status) => {
  // 'connecting' | 'connected' | 'offline'
  updateBadge(status);
});
```

`WebSocketProvider` reports status across its exponential-backoff
reconnect cycle. The demos render this as a colored dot in the header.

## What syncs and what doesn't

| Kind | Examples | Where it lives |
|---|---|---|
| **Document state** (shared) | cells, formulas, row/col structure, styles, block content, sheet existence & names | the Y.Doc |
| **Presence** (ephemeral) | a user's cursor / cell selection | the `awareness` channel |
| **View state** (per-user, local) | which sheet is active, scroll position, zoom | never synced |

If two users would reasonably want different values for something at the
same time, it's view state — keep it local.

## End-to-end encryption

Because a `CollabProvider` only ever handles opaque `Uint8Array` payloads,
**encryption is a decorator** — wrap any provider with one that encrypts
on `send` and decrypts on `onMessage`. The relay server then sees only
ciphertext: it stays a zero-knowledge hub.

The library deliberately does **not** ship a built-in encrypted provider —
key derivation, exchange, and rotation are application concerns. Here is a
complete, copy-pasteable recipe using the WebCrypto AES-GCM primitives
every modern browser ships.

### `EncryptingProvider`

```ts
import type {
  CollabChannel,
  CollabMessageHandler,
  CollabProvider,
  CollabStatus,
  CollabStatusHandler,
} from '@b32nio/spindle-shared';

/**
 * Wraps any CollabProvider, encrypting every payload with AES-GCM before
 * it leaves the process and decrypting every payload on the way in.
 *
 * Wire frame produced: [ 12-byte random IV | ciphertext+tag ].
 *
 * send() and the inbound handler are async (WebCrypto is async), so each
 * direction is funneled through a promise chain to preserve ordering.
 * Yjs tolerates reordering, but in-order delivery keeps behavior obvious.
 */
export class EncryptingProvider implements CollabProvider {
  private sendChain: Promise<unknown> = Promise.resolve();

  constructor(
    private readonly inner: CollabProvider,
    private readonly key: CryptoKey, // an AES-GCM CryptoKey, 128 or 256-bit
  ) {}

  connect(roomId: string): Promise<void> {
    return this.inner.connect(roomId);
  }

  disconnect(): void {
    this.inner.disconnect();
  }

  getStatus(): CollabStatus | undefined {
    return this.inner.getStatus?.();
  }

  onStatusChange(handler: CollabStatusHandler): () => void {
    return this.inner.onStatusChange?.(handler) ?? (() => {});
  }

  send(channel: CollabChannel, payload: Uint8Array): void {
    // Serialize encrypts so frames leave in the order send() was called.
    this.sendChain = this.sendChain
      .then(async () => {
        const iv = crypto.getRandomValues(new Uint8Array(12));
        const ct = await crypto.subtle.encrypt(
          { name: 'AES-GCM', iv },
          this.key,
          payload,
        );
        const framed = new Uint8Array(12 + ct.byteLength);
        framed.set(iv, 0);
        framed.set(new Uint8Array(ct), 12);
        this.inner.send(channel, framed);
      })
      .catch((err) => {
        console.error('[EncryptingProvider] encrypt failed:', err);
      });
  }

  onMessage(channel: CollabChannel, handler: CollabMessageHandler): () => void {
    let recvChain: Promise<unknown> = Promise.resolve();
    return this.inner.onMessage(channel, (framed, from) => {
      recvChain = recvChain
        .then(async () => {
          if (framed.length < 13) return; // too short to contain IV + tag
          const iv = framed.slice(0, 12);
          const ct = framed.slice(12);
          const pt = await crypto.subtle.decrypt(
            { name: 'AES-GCM', iv },
            this.key,
            ct,
          );
          handler(new Uint8Array(pt), from);
        })
        .catch((err) => {
          // A decrypt failure means a wrong key or a tampered frame —
          // drop it rather than crashing the session.
          console.error('[EncryptingProvider] decrypt failed:', err);
        });
    });
  }
}
```

### Using it

```ts
// Derive a key however your app manages secrets. A passphrase example:
async function deriveKey(passphrase: string, salt: Uint8Array) {
  const base = await crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(passphrase),
    'PBKDF2',
    false,
    ['deriveKey'],
  );
  return crypto.subtle.deriveKey(
    { name: 'PBKDF2', salt, iterations: 200_000, hash: 'SHA-256' },
    base,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt'],
  );
}

const key = await deriveKey(userPassphrase, perDocumentSalt);
const provider = new EncryptingProvider(
  new WebSocketProvider({ url: 'wss://collab.example.com' }),
  key,
);
await workbook.attachCollab(provider, identity, { roomId: 'quarterly-plan' });
```

### Key management is yours

The recipe handles *transport* encryption. The hard parts remain the
application's responsibility:

- **Key derivation** — from a passphrase (as above), or a per-document key
  unwrapped from the user's key.
- **Key exchange** — how a second user obtains the document key (key
  wrapping with each member's public key, an invite link fragment, etc.).
- **Rotation & membership** — revoking access means re-keying.

Awareness payloads (cursors) flow through the same wrapper, so presence
is encrypted too — a cursor position can leak structure, so this matters.

## Undo / redo

When collab is attached, undo/redo route through Yjs's `UndoManager`
(exposed on the handle as `undoManager`). Each user undoes **their own**
edits, and the inverse operations broadcast to peers like any other edit.
`workbook.undo()` / `document.undo()` do the right thing automatically.

## Known limitations (v1)

- **Documents must be single-section** to attach collab. Multi-section
  support needs section-boundary nodes in the editor schema.
- **Same-cell concurrent edits are last-writer-wins.** Edits to
  *different* cells always merge cleanly.
- The reference relay server keeps room state only in connected peers'
  memory — if everyone disconnects, add server-side persistence to retain
  the room.
