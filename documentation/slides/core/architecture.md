# Architecture

`slides-core` is layered internally: a **scene layer** (pure geometry + elements, container-agnostic) under a **deck layer** (slides, order, themes, layouts, notes). The React package never reaches into internals; it talks to the engine through semantic events and keyed subscriptions.

## The scene layer

Pure, fully unit-tested, and free of any notion of a slide or a viewport.

- **`scene/types.ts`** — the element union (`text` · `shape` · `image` · `line`), `Frame`, the `Color`/`Fill`/`Stroke` unions, and `PlaceholderMeta`. Every element extends `ElementBase` with an `id`, a `containerId`, a fractional `index` (z-order), and optional `groupId` / `parentGroupId` (nested groups) / `locked` / `opacity`.
- **`scene/fractional-index.ts`** — ordering keys. A base-62 lexicographic midpoint plus a 2-char random jitter suffix, so an item can be placed between two neighbours without moving anyone else (Yjs has no move op) and concurrent inserts into the same gap converge instead of colliding. Ties break by id.
- **`scene/geometry.ts`** — rotated-rectangle math: corners, axis-aligned bounds, point-in-frame, and a separating-axis test for marquee hit-testing.
- **`scene/transform.ts`** — `resizeFrame` and `rotateFrame`. Resize works in the frame's local space with the opposite anchor pinned; a property test asserts anchor-invariance at any rotation. Rotation snaps within 3° of each 45° step (15° with Shift).
- **`scene/snapping.ts`** — smart guides: collect targets once per gesture (other elements' AABB edges/centres + slide edges/centre), then compute the nudge and guide lines per move.
- **`scene/align.ts`**, **`scene/z-order.ts`**, **`scene/group.ts`** — align/distribute, front/back/forward/backward (returning new fractional indices), and group rotate/scale (a nested group transforms as the flat list of its leaves; `childFrameInParent` composes DrawingML group transforms for import).
- **`scene/group-tree.ts`** — nested groups: the tree is read back from each element's `groupId` (innermost group) and `parentGroupId` (the group that group sits in), so it travels with the elements and needs no group record. `groupParents` / `groupChain` / `rootGroupOf` resolve it deterministically (smallest element id wins on conflicting parents; cycles are cut).
- **`text/`** — the ProseMirror schema (deliberately flat: `doc > paragraph+ > inline*`, matching PPTX text bodies) plus pure JSON helpers used by the static renderer and the idle-formatting path.

## The deck layer

`DeckImpl` is the engine class (the `WorkbookImpl` analogue). It owns flat `Map`s of slides and elements.

- **Immutable records.** Every mutation replaces the object, so React can version a snapshot by reference.
- **Semantic events.** Each mutator emits a typed event (`elementChange { elementId, keys }`, `slideAdd`, `themeChange`, …) on the shared `EventEmitter`. Multi-element operations run inside one `batch()` and one history entry.
- **History.** A snapshot undo/redo stack — bypassed under collaboration, where undo routes to a `Y.UndoManager` instead.
- **View state is local.** Selection and the active slide live on the engine for convenience but are never serialized into the CRDT doc; they're shared through awareness instead.

## React reactivity

The load-bearing performance decision. `DeckProvider` owns an **`ElementStore`** — a keyed listener registry rather than a single trigger counter. Hooks use `useSyncExternalStore` with the engine's immutable records as snapshots, so:

- `useElement(id)` re-renders exactly one `ElementView`.
- `useSlideElementIds(slideId)` returns a stable array identity unless membership or z-order changes.
- A theme switch re-renders through `useTheme()` with no data migration — renderers resolve symbolic colors at paint time.

**Gestures never touch React state per move.** The controller snapshots frames + snap targets at gesture start and, per pointermove, writes `style.transform` straight onto the registered wrapper DOM nodes (a node registry in context) and pushes guides into a transient store consumed only by tiny overlay components. On pointerup a single `deck.setFrames()` commit reconciles React once.
