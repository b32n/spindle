# Collaboration

Real-time editing is opt-in and lives behind `deck.attachCollab()`. The base engine never imports Yjs — the binding is a lazily-loaded chunk, so consumers who don't collaborate never pay for it.

```ts
import { InMemoryProvider } from '@b32nio/spindle-shared';

const handle = await deck.attachCollab(
  new InMemoryProvider(),            // or a WebSocket CollabProvider
  { userId: 'u1', displayName: 'Alice', color: '#2D7FF9' },
  { persistenceKey: 'my-deck' },     // optional IndexedDB mirror
);
// …later
deck.detachCollab();
```

## Y.Doc schema

| Type       | Shape |
|------------|-------|
| `meta`     | `Y.Map` — id, title, slideSize, theme, layouts (LWW) |
| `slides`   | `Y.Map<slideId, Y.Map{ index, layoutRef?, background?, notes? }>` |
| `elements` | `Y.Map<elId, Y.Map{ containerId, type, x, y, w, h, rotation, index, groupId?, …, props, richText? }>` |
| `threads`  | `Y.Map<threadId, JSON>` — comments, outside the undo scope |

Frame fields are **individual Y.Map keys**, so concurrent move / rotate / z-reorder on one element merge per-key. Type-specific styling rides in one `props` JSON key (last-writer-wins, v1-acceptable). The rich-text body is a **`Y.XmlFragment`** bound via `y-prosemirror`. Fractional `index` + jitter make concurrent reorders converge.

## Mirror + observers

The binding runs two directions, both echo-guarded:

- **Mirror** — the engine's semantic events are translated into Y writes under a private `LOCAL_ORIGIN`.
- **Observe** — granular observers on `elements` / `slides` / `meta` / `threads` apply remote (and undo) changes back into the engine via dedicated `_applyRemote*` hooks. There's no whole-doc reload on the hot path.

A peer joining an existing room syncs first, hydrates only if the room was empty, then reconciles the engine to exactly match the Y.Doc — so joining never duplicates content.

## Undo

Under collaboration, `deck.undo()` / `redo()` route to a `Y.UndoManager` scoped to `[meta, slides, elements]` with `trackedOrigins` of `LOCAL_ORIGIN` and `ySyncPluginKey`. Undo therefore produces inverse Y operations that broadcast to peers, and it only ever reverts the local user's own edits. Rich-text fragments created after the manager exists are covered transitively.

## Presence

Each peer publishes `editing: { slideId, elementId }` in awareness while its live editor is open. `RemotePresenceOverlay` draws a colored outline + name tag around the element a collaborator is editing. In-text remote carets are deferred: `y-prosemirror`'s cursor plugin assumes one fragment per document, which is unsafe with per-element fragments.

## Comments

Comment threads mirror through the top-level `threads` map, deliberately outside the undo scope (comments aren't undoable). Threads whose anchored element is deleted are *orphaned* — kept, never silently dropped, and shown in the panel under "No longer attached".
