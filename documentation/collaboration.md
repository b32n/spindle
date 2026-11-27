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

- **Multi-section documents attach.** The shared document carries every
  section (a `section_break` node ends each section but the last, and
  `sectionsMeta` holds each section's page setup, header and footer) and the
  editor shows them section by section. A section's setup is read from the
  local document, so changing page setup in one client is not propagated to
  peers yet; the Insert → Section break UI comes with the section editing work.
- **Content the shared layout cannot hold blocks attach.** `getCollabBlockers(document)`
  lists it (no sections, two sections with one id, block or inline types this
  version does not know) and `attachCollab` rejects with a `CollabBlockedError`
  carrying that list, instead of silently dropping it. The list is empty for
  every document the importers produce today.
- **Slides: version gate and blockers.** A slides room records the layout version in `meta.schemaVersion`
  (a room without the key reads as 1, so existing rooms keep working). `attachCollab` refuses a room written by a
  newer client with a `CollabSchemaVersionError` that names both versions, before it seeds or edits anything.
  `getCollabBlockers(deck)` and `CollabBlockedError` mirror the documents API; the list is empty today, and a later
  deck construct either gets its own Y keys or registers a blocker. Version 4 added layout `elements`, `textStyles` and
  `masterName` (inside the `meta.layouts` JSON, last writer wins: two users editing layouts at the same moment keep one
  edit) and the per-slide `hidden` key. Version 5 added the shape presets of the preset geometry library and
  `unknown:<prst>` shapes (inside the element `props` JSON).
- **Sheets: version gate and blockers.** A sheets room records the layout version in `meta.schemaVersion`
  (a room without the key reads as 1, so existing rooms keep working). `attachCollab` refuses a room written by a
  newer client with a `CollabSchemaVersionError` that names both versions, before it seeds or edits anything.
  `getCollabBlockers(workbook)` and `CollabBlockedError` mirror the documents and slides API; the list is empty
  today, and a later workbook construct either gets its own Y keys or registers a blocker.
- **Content controls.** A Word content control is a ProseMirror node
  (`content_control` around blocks, `content_control_inline` in a paragraph), so it
  syncs like any other content and does not block attach. Concurrent edits merge as
  for text and attributes (two people toggling the same check box is
  last-writer-wins). A control's lock is not enforced against a collaborator.
- **Footnotes and endnotes.** A note reference is a ProseMirror node holding the
  text of its note (`note_ref`), so it syncs like other content and does not block
  attach. The footnotes at the bottom of the pages and the endnotes area are drawn from the references, not stored. A note's text is
  edited in a small editor of its own and written back into the reference when the edit
  ends, so two people editing the same note at once is last-writer-wins on its text.
- **Table of contents.** A table of contents is a ProseMirror node (`toc`) holding its entries, so it syncs like
  other content and does not block attach. Updating it writes the entries and adds bookmarks to the headings in one
  transaction; two people updating it at once is last-writer-wins on its entries.
- **Tracked changes.** An insertion, deletion or formatting change is a ProseMirror mark (`insertion`, `deletion`,
  `formatChange`: revision id, author, date; a move's name; the formatting before) on the text, and a paragraph's or table
  row's change is an attribute of the node, so they sync like other content and do not block attach. Two people editing
  inside the same change merge as for text; text a person adds to a change while someone else settles it is settled with it.
  A line break inside tracked text is a newline character in the text (a marked `hard_break` node would lose its change,
  because marks sync on text only); the editor inserts that for Shift+Enter in tracked text, and the DOCX reader makes the same.
  - *Settling.* Accepting or rejecting a change in the review panel is an ordinary edit (text deleted, marks removed,
    attributes set) that syncs like any other; the coordination of two people settling the same change at once is not done
    yet (one of two concurrent settlements wins, the same for everybody), and the Final / Original view is local to each
    client. A change in the panel is found by its position, so a remote edit that moves it between rendering the panel and
    clicking Accept makes that click do nothing (the panel shows the new list right after).
  - *Paragraph marks and rows.* A mark or row keeps one attribute per kind of change (`revisionInsertion`,
    `revisionDeletion`, `revisionFormat`), each a Yjs key of its own, so two people suggesting different things on the same
    paragraph mark or row at once (one presses Enter, the other deletes the mark) both keep theirs, and each can be accepted
    or rejected alone. Two people suggesting the same kind on the same mark at once is last-writer-wins on that attribute:
    one stays, the same on every client.
  - *Concurrent split (residual limit).* When one person splits a paragraph while another marks its mark, Yjs keeps the
    original element for the first half, so the other person's record lands on the first half (next to the insertion that the
    split recorded there), not on the second half that holds the original mark. Accepted together the two cancel (the
    paragraphs are joined again), and each can still be rejected alone: rejecting the split joins the paragraphs and moves the
    other person's deletion record to the joined mark (the original one), where it is still open. The same happens outside
    collaboration when somebody deleted a mark that somebody else had inserted and the insertion is rejected: the deletion
    record cannot be told from the concurrent-split case, so it stays open on the joined mark. Pressing Enter in a paragraph
    that carries a record copies the record to both halves (ProseMirror copies a node's attributes on a split): both halves
    then show the same change, which is accepted or rejected as one.
  - *Colours.* Two people under one display name are one author, drawn in the smallest of their cursor colours (the name is
    trimmed, as on a suggestion). A person in the room who has no change in the document keeps their cursor colour, which can
    equal the colour an author of a change was given from the palette, so two people can be drawn alike.
- **Suggesting mode** records the local user's own edits only: changes that arrive from a collaborator, and undo and redo, are
  never turned into suggestions. A collaborator's suggestions are ordinary marks and attributes that sync as such, also when they
  were made offline. The author on a suggestion is the name the user joined the room under (`CollabIdentity.displayName`; comments still use the provider's `currentUser`), and
  people who are in the room are drawn in the colour of their cursor (`collabAuthorColors`); once they leave, their changes are
  drawn in a colour the document's authors share out, so a colour can change when somebody joins or leaves. Two people accepting
  the same change both succeed; one accepting and the other rejecting it at the same time ends in one of the two, the same for
  everybody (it is two ordinary edits of the same text). For an inserted run the rejection always wins (removing the text outweighs removing the mark); for a deleted run, which of the two wins is not specified, only that every client ends in the same state. People who suggest in one room need distinct display names: a change is attributed by the author's name, so two people with the same name are one author to each other, and one's Backspace over the other's insertion removes the text outright.
- **Mixed-version rooms.** Rooms record a `schemaVersion` (a missing value means
  1, the layout without section breaks; 2 adds section breaks; 3 adds content
  controls; 4 adds footnote and endnote references; 5 adds the table of contents; 6 adds tracked insertions and deletions; 7 adds formatting changes, moves, and changes of paragraph marks and table rows). A client that finds a newer version
  than it supports rejects `attachCollab` with a `CollabSchemaVersionError`
  instead of editing a document whose newer nodes it would drop. Clients from
  before this check have no such guard: do not open a multi-section room with
  them. Existing single-section rooms need no migration.
- **Same-cell concurrent edits are last-writer-wins.** Edits to
  *different* cells always merge cleanly.
- The reference relay server keeps room state only in connected peers'
  memory — if everyone disconnects, add server-side persistence to retain
  the room.
