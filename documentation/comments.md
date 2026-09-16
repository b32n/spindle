# Comments

The spreadsheet, the document editor, and the slides editor all support
**comment threads** — select content, attach a thread, reply, resolve, and
@-mention teammates. Comments sync between users over the same collaboration
layer as the rest of the document.

## Design in one paragraph

A **thread** is anchored to a piece of content and holds an ordered list of
**comments** (`comments[0]` is the root; the rest are replies). The thread
*content* model is identical across editors — only the **anchor** differs.
Thread content lives in a top-level `threads` map inside the Y.Doc, so it
syncs like everything else; it is deliberately kept out of undo/redo.

```
CommentStore  (sheets)  /  DocsCommentStore  (docs)
        │  addThread / addReply / resolve / reopen / delete
        ▼
   threads  ──►  Y.Map<threadId, CommentThread>   (synced; not undoable)
        │
   anchor ──┬── sheets: { rowId, colId }    — a stable cell id
            └── docs:   a `comment` ProseMirror mark in the body text
```

Sheets anchors a thread to a cell by its **stable row/column id**, so the
thread survives row/column insert, delete and sort. Docs anchors a thread
with a **`comment` mark** on the text range — the mark rides in the document
body, so y-prosemirror syncs it and ProseMirror maps it through edits for
free.

## Quick start

Comments are on by default in the React editors — the toolbar gains a
**Comment** button and a comments panel. To attribute comments and enable
@-mentions, pass a few props to the provider:

```tsx
// Sheets
<WorkbookProvider
  workbook={workbook}
  currentUser={{ id: 'u_42', name: 'Alice' }}
  mentionableUsers={[
    { id: 'u_42', name: 'Alice' },
    { id: 'u_43', name: 'Bob' },
  ]}
  onCommentEvent={(event) => notifyBackend(event)}
>
  <WorkbookCanvas width={800} height={600} />
</WorkbookProvider>
```

```tsx
// Docs — identical props on DocumentProvider
<DocumentProvider
  document={doc}
  currentUser={{ id: 'u_42', name: 'Alice' }}
  mentionableUsers={mentionableUsers}
  onCommentEvent={(event) => notifyBackend(event)}
>
  <DocumentEditor />
</DocumentProvider>
```

All three props are optional. `currentUser` defaults to a generic local
user; without `mentionableUsers` the `@` autocomplete simply never appears.

## The model

The thread and comment types are editor-agnostic and exported from
`@b32nio/spindle-shared` (and re-exported from each `*-core` package):

```ts
type CommentStatus = 'open' | 'resolved';

interface Comment {
  id: string;
  authorId: string;
  authorName: string;   // snapshotted, so offline authors still render
  body: string;
  createdAt: string;    // ISO-8601
  editedAt?: string;
  mentions?: string[];  // user ids @-mentioned in the body
}

interface CommentThread<Anchor> {
  id: string;
  anchor: Anchor;       // { rowId, colId } for sheets; { quote } for docs
  status: CommentStatus;
  comments: Comment[];
  createdAt: string;
  resolvedBy?: string;
  resolvedAt?: string;
}
```

## Working with threads in code

Each editor owns a comment store. The store is the source of truth; the
React panels are just a view of it.

```ts
// Sheets — store for a sheet (defaults to the active sheet)
const store = workbook.getCommentStore();
const thread = store.addThread({ rowId, colId }, 'Looks off?', author);
store.addReply(thread.id, 'Fixed.', author);
store.resolveThread(thread.id, author);

// Docs — store for the document
const docStore = document.getComments();
docStore.getThreads();
```

Both stores expose the same surface: `addThread`, `addReply`,
`editComment`, `deleteComment`, `deleteThread`, `resolveThread`,
`reopenThread`, `getThreads`, `getThread`. `addThread` / `addReply` accept
an optional trailing `mentions: string[]`.

In React, the sheets package also offers a `useComments()` hook that binds
this store to the active sheet and re-renders on changes.

## @-mentions

Supply `mentionableUsers` (a `{ id, name }[]` directory) and the comment
composer gains an `@` autocomplete. Picking a user inserts `@Name` and
records their id; the rendered comment highlights `@Name`, and the id lands
in `comment.mentions` — so a notification handler knows exactly who to
ping.

The directory is a static array in v1. It need not include `currentUser`,
and an unknown mentioned id simply renders as plain text.

## The `onCommentEvent` callback

`onCommentEvent` is the hook for **notifications**. It fires for the local
user's own comment actions, carrying a semantic event the host app can
forward to a backend (push, email, etc.):

| `event.type` | Fired when | Carries |
|---|---|---|
| `thread-created` | a new thread is started | `comment`, `mentions` |
| `reply-added` | a reply is posted | `comment`, `mentions` |
| `comment-edited` | a comment body is edited | `commentId` |
| `comment-deleted` | a comment is removed | `commentId` |
| `thread-deleted` | a whole thread is removed | — |
| `thread-resolved` / `thread-reopened` | status changes | `by` |

Every event also carries `threadId` and `anchor` (and `sheetId` for
sheets). The same stream is available without React via
`workbook.on('commentEvent', …)` / `document.on('commentEvent', …)`.

It fires for **local actions only** — not for threads arriving from a
collaborator. The expected architecture is: the actor's client reports the
event to a backend, which fans notifications out to everyone else.

## What syncs

| Kind | Where it lives | Notes |
|---|---|---|
| Thread content (comments, replies, status) | a `threads` Y.Map in the Y.Doc | last-writer-wins per thread |
| Docs comment highlight (the anchor) | a `comment` mark in the body text | synced by y-prosemirror with the rest of the content |
| Which thread is open in the panel, draft reply text | local React state | never synced — per-user view state |

Comments work without collaboration too: the store is the source of truth
and round-trips through `getData()` / `setData()`. When collab is attached,
thread changes mirror into the Y.Doc automatically.

Comments are **excluded from undo/redo** — undoing a document edit never
removes a comment, and `restoreSnapshot` carries threads across the undo
timeline.

## Known limitations (v1)

- **Sheets comments anchor to a single cell** — range comments are not yet
  supported.
- **Thread content is last-writer-wins** — two replies to the *same* thread
  within one sync window can collide (consistent with cell-level LWW).
  Edits to *different* threads always merge cleanly.
- **Overlapping comments** on the same span of document text are deferred.
- `onCommentEvent` covers local actions only; client-side toasts for
  remote comments would need separate wiring.
