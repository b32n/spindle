# Spindle Slides — Implementation Plan

> **Status: implemented (v0.3.0).** This is the original implementation brief,
> kept as a historical design record — the slides library now ships. Some
> scope has since changed (notably **tables** and **presenter view** shipped,
> though the brief lists them as excluded/follow-ups below). For the current
> state and remaining work, see **[TODO.md](TODO.md)**.
>
> This is a self-contained implementation brief for the new slides library. It was produced after exploring the existing docs/sheets packages and researching shipped slide editors (Google Slides, Pitch, tldraw, Figma Slides) and OSS prior art (PPTist, Excalidraw, BlockSuite). All architectural decisions below are final — implement, don't re-litigate. Follow the phases in order; each leaves master shippable; commit at every phase boundary.

## Context

Spindle ships two of its three promised surfaces (docs, sheets); slides is the pending third. The website and CONTRIBUTING.md currently promise slides "on the docs engine," but we decided slides gets its own standalone `slides-core` — the docs model is a linear flowing block tree with no concept of positioned elements, z-order, rotation, or groups.

Research confirmed the niche is empty — no React + permissive-OSS + Yjs-collab + rich-text-in-shapes slides editor exists — and converged on the architecture below.

## Locked decisions

1. **Two new packages**: `packages/slides-core` (`@b32nio/spindle-slides-core`, zero React) + `packages/slides-react` (`@b32nio/spindle-slides-react`), following existing monorepo conventions (rollup dual cjs/esm, declarations from main entry only, `./collab` subpath keeps Yjs out of the base bundle, strict TS, co-located Jest tests).
2. **Slides-first, canvas-ready**: slides-core internally layered — container-agnostic **scene layer** (elements, transforms, selection math, snapping, z-order, groups, rich text) + **deck layer** (slides, order, masters/layouts/placeholders, themes, notes). Elements reference `containerId`; camera/viewport never enters the data model. A future Miro-style `spindle-canvas` reuses the scene layer.
3. **Rendering: DOM** — absolutely-positioned HTML per element, inline SVG shape geometry, CSS-transform zoom. Not canvas. (Basis: editable rich text requires contentEditable; every shipped slides editor with rich text is DOM; slides show 10–50 elements, far below canvas-justifying scale.)
4. **Rich text: raw ProseMirror** (already in-repo via docs-core). PM-JSON stored in the model; idle elements render statically; exactly one live PM editor mounts on the actively-edited element (tldraw's shipped pattern).
5. **Data model: PPTX-semantic, CRDT-mechanic** — theme = 12 symbolic color slots + major/minor fonts; `layoutRef` + placeholder `{type, idx}` metadata in the schema from day one (retrofitting caps PPTX fidelity — the PPTist lesson); fractional-index strings for slide order AND z-order, never array order (Yjs has no move op).
6. **Collab: Yjs behind `attachCollab()`** — Y.Map per element, Y.XmlFragment per text body via y-prosemirror, Y.UndoManager with trackedOrigins, awareness for presence. Per-user view state (active slide, zoom, selection) stays out of the CRDT doc (hard project rule). Granular observers, not sheets' v1 whole-doc reload.
7. **Scope = Tier 1+2**: text/shape/image/line elements, multi-select + move/resize/rotate handles, smart guides + snapping, align/distribute, group/ungroup (flat), z-order, filmstrip, speaker notes, themes + layout gallery, present mode, shortcuts, clipboard, undo/redo, element-anchored comments, collab, PDF export (print pipeline), `getData()/setData()`.
   **Excluded (follow-ups)**: PPTX import/export, tables, charts, video, animations, nested groups, presenter view w/ timer, in-text remote carets, gradient/image fills.
8. **Reuse from shared**: `CollabProvider`, `CollabIdentity`, `InMemoryProvider`, `CommentThread<Anchor>`, and the shared `EventEmitter<T>` directly (verified it has typed events + `batch()` — don't fork it a third time like docs/sheets did).
9. Defaults: slide size **1280×720** (16:9 @ 96 DPI, matching docs-core's DPI convention).

## Package layout

### packages/slides-core/src
```
index.ts                 public exports (no collab)
types.ts                 deck types: Slide, SlideData, DeckData, event union
deck.ts                  DeckImpl engine class (WorkbookImpl analogue)
history.ts               local snapshot undo (non-collab only)
comments.ts              SlidesCommentStore + ElementCommentAnchor (port sheets-core/src/comments.ts)
serialization.ts         normalizeDeckData() for setData validation/defaults
utils/id.ts
scene/                   SCENE LAYER — pure, container-agnostic, fully unit-tested
  types.ts               element union, Frame, Color/Fill/Stroke, PlaceholderMeta
  elements.ts            per-type factories + defaults
  fractional-index.ts    indexBetween/indexesBetween/sortByIndex — in-house ~60 lines, base-62
                         midpoint + 2-char random jitter suffix (concurrent-insert collision safety)
  geometry.ts            rotated-rect math: AABB, point-in-rect, frameIntersectsRect (marquee)
  transform.ts           resizeFrame()/rotateFrame() — pointer→local-space, fixed-anchor resize,
                         aspect lock, min 8×8, rotation snap within 3° of 45° steps
  snapping.ts            collectSnapTargets() once per gesture + computeSnap() per move → {dx,dy,guides}
  align.ts               alignFrames()/distributeFrames() (pure)
  z-order.ts             front/back/forward/backward → new fractional indices (pure)
  group.ts               group AABB, rotateGroup(), scaleGroup() (flat groups via shared groupId)
text/                    SCENE LAYER (rich text)
  schema.ts              slidesSchema: minimal PM schema (see Rich text)
  model.ts               RichTextDoc JSON type + pure helpers: emptyRichText, to/from plain text,
                         applyMarkToDoc, setParagraphAttrs (shared by toolbar idle-path + tests)
theme/                   DECK LAYER
  types.ts               ThemeData (12 slots + major/minor fonts), LayoutData, PlaceholderDef
  resolve.ts             resolveColor (slot→hex + lumMod/lumOff/alpha), resolveFont,
                         resolvePlaceholderStyle (element → layout placeholder → theme defaults)
  builtin.ts             3 themes, 5 layouts (title / title+content / section / two-content / blank)
collab/                  separate rollup entry — only place importing yjs
  y-schema.ts            Y.Doc layout, hydrate/serialize, createElementYMap, getElementRichTextFragment
  binding.ts             attachCollabToDeck(): provider wiring, granular observers, Y.UndoManager, IndexedDB
```

### packages/slides-react/src
```
context/DeckContext.tsx      DeckProvider (engine as prop; currentUser/mentionableUsers/onCommentEvent like WorkbookProvider)
context/element-store.ts     keyed subscriptions: subscribeElement/Slide/Deck/Selection
hooks/                       useDeck, useSlideIds, useSlide, useElement, useSelection, useActiveSlide,
                             useKeyboardShortcuts, useClipboard, useComments (port from sheets-react)
interactions/gesture.ts      imperative drag/resize/rotate/marquee controller (no React state per move)
interactions/transient-store.ts  per-gesture mutable state (guides, live sizes) + subscribe
interactions/coords.ts       screen↔slide mapping under zoom
components/
  SlidesEditor.tsx           root layout; owns keyboard + clipboard
  Toolbar.tsx  Filmstrip.tsx  SlideStage.tsx  SlideView.tsx
  elements/ElementView.tsx + Text/Shape/Image/Line views + StaticRichText.tsx
  RichTextEditor.tsx         the single live PM mount
  SelectionOverlay.tsx  GuidesOverlay.tsx  RemotePresenceOverlay.tsx
  NotesPanel.tsx  PresentMode.tsx  ContextMenu.tsx  CommentsPanel.tsx (port)
  pdf/export-pdf.ts
```

## Data model (core sketches)

```ts
// scene/types.ts
interface Frame { x: number; y: number; w: number; h: number; rotation: number } // deg, about center
type Color = { kind: 'theme'; slot: ThemeColorSlot; alpha?; lumMod?; lumOff? } | { kind: 'rgb'; hex: string; alpha? };
type Fill  = { kind: 'none' } | { kind: 'solid'; color: Color };
interface ElementBase extends Frame {
  id: string; containerId: string;
  index: string;                    // fractional z-order key
  groupId?: string; locked?: boolean; opacity?: number;
  placeholder?: { type: PlaceholderType; idx: number };   // PPTX fidelity, in schema NOW
}
// union: TextElement {richText, bodyStyle, fill?, stroke?} | ShapeElement {shape: ShapePreset (~18 presets),
// adjustments?, flipH/V, fill, stroke, richText?} | ImageElement {src, naturalW/H, flipH/V} |
// LineElement {flipH/V (diagonal), stroke, startArrow?, endArrow?}

// types.ts (deck)
interface Slide { id: string; index: string; layoutRef?: string; background?: Fill; notes?: RichTextDoc }
interface DeckData {
  id: string; title: string; slideSize: { w: number; h: number };
  theme: ThemeData; layouts?: LayoutData[];
  slides: Array<Slide & { elements: SlideElement[] }>;   // nested per-slide, sorted on emit (engine keeps flat maps)
  threads?: SlidesCommentThread[];
  selection?: { slideId?: string; elementIds: string[] };
}
```

## DeckImpl engine

Instance-owned; internal `slides: Map<id, Slide>`, `elements: Map<id, SlideElement>` (flat). **Records are immutable** — every mutation replaces the object so `useSyncExternalStore` snapshots version by reference. Every mutator: update maps → mirror to Y.Doc when attached (`ydoc.transact(fn, LOCAL_ORIGIN)`) → emit semantic event → `recordHistory()` when not attached. Multi-element ops (`updateElements`, `setFrames`) = one `batch()` + one Y transaction + one history entry.

Events: `deckChange, themeChange, slideAdd/Delete/Move/Change, elementAdd/Delete/Change {elementId, keys}, selectionChange, activeSlideChange, commentChange, commentEvent`.

API groups: deck (title, slideSize, theme, layouts) · slides (add w/ `{afterSlideId, layoutId}`, duplicate, delete, move via fractional reindex, background, notes, applyLayout) · elements (add/get/getForSlide-sorted/update/updateElements/delete/duplicate) · transforms (moveElements, setFrames — gesture commit path) · z-order/group/align-distribute (thin wrappers over scene/* pure fns) · rich text (setElementRichText, applyTextFormat) · selection + activeSlide (**local only, never serialized to CRDT; shared via awareness**) · history (undo routes to Y.UndoManager when attached — sheets pattern) · getData/setData (throws while attached) · attachCollab/detachCollab (lazy-import ./collab) · getComments.

## Interaction design (scene layer, all pure + tested)

- **Hit testing**: DOM-first — wrapper divs carry `data-element-id`, browser handles rotated hit-testing via CSS transforms; `closest()` on pointerdown. Geometry math only for marquee (separating-axis rotated-rect vs AABB) and snap bounds. Lines get a wide invisible SVG hit-stroke. Click selects whole group (expand via groupId); double-click drills in (PowerPoint convention).
- **Resize**: transform pointer to element-local space, fixed opposite anchor, clamp, aspect-lock (corners; images default-locked), recompute center so anchor stays fixed in world space. Property-test anchor invariance under rotation.
- **Snapping**: targets collected once at gesture start (other elements' AABB edges/centers + slide edges/center); threshold 5 screen px ÷ zoom; move snaps 3 x-lines + 3 y-lines, resize snaps moving edges only; emits GuideLine render data.
- **Groups v1 (flat)**: move trivial; rotate = orbit member centers + add Δθ to member rotations; resize = scale centers + w/h relative to group AABB — lossy for rotated members under non-uniform scale, so group/multi corner handles are aspect-locked by default (documented).

## Rich text

**PM schema deliberately flatter than docs** (PPTX text bodies are flat paragraphs with bullet attrs, not nested lists — keeps static render + future PPTX export honest): `doc > paragraph+ > inline*`; paragraph attrs `{align, listType: none|bullet|number, indent 0-8, lineHeight, spaceBefore/After}`; marks `bold, italic, underline, strikethrough, link, textColor (Color union as attr — theme slots survive!), fontFamily (literal | 'major'|'minor'), fontSize`.

- **Idle**: `StaticRichText` walks JSON → React nodes, memoized on `(richText, theme)`.
- **Lifecycle**: double-click/Enter → `editingElementId` set → that view swaps in `RichTextEditor` with identical box metrics (zero visual jump). Blur/Escape/slide-switch → destroy view, restore static. Never unmount mid-IME (`view.composing` guard).
- **Non-collab**: editor from JSON, commit on blur via `setElementRichText` (one history entry); local prosemirror-history during the session, discarded on blur (documented compromise).
- **Collab**: `ySyncPlugin(fragment)` + `yUndoPlugin({undoManager: handle.undoManager})`; fragment is truth; deep observer refreshes model JSON per fragment change.
- **Toolbar formatting**: live editor → PM transactions via registered view ref; idle selection → `applyTextFormat` walks stored JSON (pure helpers in text/model.ts).
- **Remote carets — flagged pitfall**: y-prosemirror's `yCursorPlugin` assumes one fragment per doc; with per-element fragments a cursor in element A would decorate element B. **v1: skip yCursorPlugin** — publish `editing: {slideId, elementId}` in awareness; `RemotePresenceOverlay` shows colored outline + name tag. In-text carets are a follow-up.

## Yjs schema & binding

```
meta:     Y.Map — id, title, slideSize, theme JSON, layouts JSON (LWW)
slides:   Y.Map<slideId, Y.Map{ index, layoutRef?, background?, notes? }>
elements: Y.Map<elId, Y.Map{ containerId, type, x, y, w, h, rotation, index, groupId?, locked?,
                             opacity?, placeholder?, props: JSON, richText?: Y.XmlFragment }>
threads:  Y.Map<threadId, JSON>   — outside undo scope (sheets precedent)
```

Frame props as individual Y.Map keys → concurrent move/rotate/z-reorder on one element merge per-key; type-specific styling in one `props` JSON key (LWW, v1-acceptable, mirrors sheets). Fractional `index` + jitter → concurrent reorders converge.

**Binding**: copy the sheets binding shell (`packages/sheets-core/src/collab/binding.ts`) — provider channels, SyncStep1 handshake, awareness encode, IndexedDB-restore-before-hydrate, teardown — but upgrade to **granular observers** (no whole-doc reload on the hot path): elements shallow observer → add/delete events; deep observer → patch single element record → `elementChange {keys}` (fragment events → re-derive that element's richText JSON only); slides observers → slideAdd/Delete/Change/Move; meta → deckChange/themeChange; threads → commentStore.loadJSON → `commentChange` only. Keep a `_resyncFromY()` escape hatch; decide in Phase 5 whether to retain.

**Undo**: `Y.UndoManager([meta, slides, elements], { trackedOrigins: new Set([LOCAL_ORIGIN, ySyncPluginKey]), captureTimeout: 500 })` — add `ySyncPluginKey` explicitly at construction. Fragments are covered transitively via the elements scope; add a test proving fragments created after the manager exists are tracked.

## React reactivity (the load-bearing perf decision)

- `DeckProvider` owns an **ElementStore** (keyed listener registry) instead of a single trigger counter; hooks use `useSyncExternalStore` with engine records as snapshots → each `ElementView` re-renders alone via `useElement(id)`; `useSlideElementIds(slideId)` returns stable array identity unless membership/z changes.
- **Gestures never touch React state per move**: `gesture.ts` snapshots frames + snap targets at start, per pointermove writes `style.transform` directly onto registered wrapper DOM nodes (node registry Map in context) and pushes guides/live-size into the transient store (consumed only by tiny overlay components). On pointerup → one `deck.setFrames()` commit → one batch/Y-transaction/history entry → React reconciles.
- Keyboard: one hook on a focused wrapper (not window); when editing or focus-in-input, only Escape handled globally, PM keymap owns the rest. Del, arrows nudge 1/10px, ⌘Z/⇧Z, ⌘D, ⌘C/X/V, ⌘A, ⌘G/⇧G, Tab cycle, Enter edit, ⌘]/[ z-order.
- Clipboard: `application/x-spindle-slides+json` + text/plain fallback (internal buffer if async API unavailable); paste images → data-URL ImageElement; paste text → TextElement.

## Present mode & PDF

- **PresentMode**: portal + `requestFullscreen()`, letterboxed `SlideView` at `min(vw/w, vh/h)` scale — same static renderers, zero editing chrome. Arrow/Space/PgUp-Dn/Home/End/number-jump nav; 150ms cross-fade (two stacked SlideViews) is the only v1 "transition".
- **PDF v1 (print pipeline)**: hidden iframe, clone static-rendered slides + collected styles, `@page { size: 1280px 720px; margin: 0 }` + page-break per slide, wait `document.fonts.ready` + image `decode()`, `print()`. Documented limitations (user picks "Save as PDF"; Firefox `@page size` quirks). Real PDF backend deferred.

## Comments & themes

- Port `sheets-core/src/comments.ts` → `SlidesCommentStore` with `ElementCommentAnchor {slideId, elementId}` on shared `CommentThread<Anchor>`. Orphaned threads (element deleted) stay in panel under "no longer attached" (sheets rule). Y `threads` map mirroring; `commentEvent` fires for local actions only. Port CommentsPanel + useComments from sheets-react; badge on commented elements; click thread → select element.
- Themes: renderers resolve colors at render time via `resolveColor` — theme switch is one `themeChange` re-render. `addSlide({layoutId})` **materializes** layout placeholders as real TextElements with `placeholder` meta + prompt text when empty; layout gallery on the "New slide" split-button; theme picker in Toolbar. Master-editing UI out of scope; schema carries everything PPTX export needs later.

## Phases (each leaves master shippable; commit at every phase boundary)

1. **Scaffolding + model + static render** — packages (copy docs-core rollup dual-entry pattern verbatim, sheets jest config), root build-script order, DeckImpl CRUD + events + getData/setData, fractional-index, all types, themes/layouts, StaticRichText, SlideView/ElementView (all 4 element types incl. SVG presets), read-only filmstrip + zoom stage, demo shell w/ hardcoded deck.
   *Exit*: root build/type-check/test green; round-trip + fractional-index + mutator/event tests; demo renders 3-slide deck at multiple zooms.
2. **Selection + transforms** — pointer/marquee selection, gesture controller + node registry, move/resize/rotate, snapping + guides, z-order, align/distribute, group/ungroup, snapshot undo, shortcuts, clipboard, context menu, insert from toolbar.
   *Exit*: geometry/transform/snapping/align/z/group unit tests (anchor-invariance property tests); 60fps drag w/ 50 elements (no React commits during move); undo covers every op.
3. **Rich text** — schema + model helpers, RichTextEditor lifecycle, toolbar live+idle formatting, NotesPanel.
   *Exit*: PM JSON round-trip + applyMarkToDoc tests (node env); no layout jump entering/exiting edit; IME-safe; **zoom decision checkpoint**: test PM at 50/100/200%, fall back to CSS `zoom` property if `transform: scale` coordinate math misbehaves.
4. **Deck features → usable single-user product** — layout gallery + placeholder materialization, theme picker, slide backgrounds, filmstrip DnD reorder, PresentMode, PDF export, deep duplicate-slide.
   *Exit*: author → present → PDF a full deck in the demo. **Single-user milestone.**
5. **Collab** — y-schema + binding + `./collab` rollup entry/exports subpath, granular observers, mutator mirroring, Y.UndoManager routing, per-element y-prosemirror editors, awareness presence + RemotePresenceOverlay, IndexedDB option, setData guard.
   *Exit*: y-schema round-trip tests; two-DeckImpl convergence over `InMemoryProvider` (concurrent moves, reorders, text in different elements → `getData()` deep-equal); undo-after-remote-edit test; fragment-created-after-UndoManager test; demo two-pane + `?ws=` modes work against collab-server (unchanged).
6. **Comments** — store + Y mirroring + panel + badges + onCommentEvent plumbing.
   *Exit*: ported store tests; threads sync in two-pane demo; orphaning verified.
7. **Demos, website, docs** — finish `examples/slides-demo` (sheets-demo parity: single/two-pane/`?ws=` modes), `website/src/demo/SlidesDemo.tsx` + route, `website/src/App.tsx` slide card (`pkg: '@b32nio/spindle-slides-react'`, `soon: false`, demo link) + packages-list rows + docs-core blurb fix, `CONTRIBUTING.md:16` ("Document & slides engine" → separate slides-core/slides-react lines), `documentation/slides/` (overview, core/architecture, core/collaboration, components, data-structures, TODO listing deferred Tier-3 items + manual QA checklist), package READMEs.
   *Exit*: website builds with live slides demo; root scripts green.

## Reference files (read before implementing each area)

- `packages/sheets-core/src/workbook.ts` — engine-class pattern (mutators, events, Y mirroring, history routing, setData guard)
- `packages/sheets-core/src/collab/binding.ts` + `y-schema.ts` — collab shell to copy; upgrade to granular observers
- `packages/docs-core/src/collab/y-schema.ts` — XmlFragment ↔ PM JSON via y-prosemirror helpers
- `packages/docs-core/rollup.config.js` — dual-entry build (main decl + ./collab `declaration:false`)
- `packages/docs-core/src/prosemirror/schema.ts` — PM schema conventions
- `packages/sheets-react/src/context/WorkbookContext.tsx` — provider API to mirror (improved with ElementStore)
- `packages/sheets-react/src/components/CommentsPanel.tsx` + `hooks/useComments.ts` — comment UI to port
- `examples/sheets-demo/src/App.tsx` — demo structure to mirror
- `packages/shared/src/event-emitter.ts` — use directly (verified: typed + batch())

## Verification

- Per phase: `npm run build && npm run type-check && npm test` at root (cross-package types go through dist/.d.ts — rebuild after API changes before trusting workspace type-check).
- Core logic verified by co-located Jest tests (node env): fractional-index, geometry/transform/snapping/align/z-order/group math, deck mutators/events/history, DeckData round-trip, PM model round-trip, y-schema round-trip, InMemoryProvider convergence + undo.
- No React test infra exists in the repo — don't invent it; `examples/slides-demo` is the smoke test with a manual checklist (drag perf via devtools, IME, zoom 25–400%, present mode, print) recorded in `documentation/slides/TODO.md`.
- End-to-end: run slides-demo in all three modes (single, two-pane InMemory, `?ws=` against `examples/collab-server`); author/present/export a deck; verify website demo page.

## Risks (watch during implementation)

1. PM inside CSS-scaled container — caret/coords at zoom ≠ 100% (Phase 3 checkpoint; CSS `zoom` fallback).
2. contentEditable inside rotated ancestors on iOS Safari — if buggy, temporarily render editor unrotated (PowerPoint-web does this).
3. Filmstrip thumbnails re-rendering on every commit — memoize on slide content version.
4. y-prosemirror + shared UndoManager coherence ("create element + type" as sane undo steps) — dedicated Phase 5 tests; tune captureTimeout.
5. Data-URL images bloat Y.Doc/IndexedDB — v1 accepts with size warning; upload-hook prop is a follow-up.
6. Idle-selection text formatting in collab rewrites whole fragment (LWW during someone else's typing) — v1-acceptable, revisit if it bites.
