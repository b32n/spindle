# Components

`@b32nio/spindle-slides-react` is a thin, composable layer over the engine. Everything hangs off a single provider.

## Provider

```tsx
<DeckProvider
  deck={deck}
  currentUser={{ id: 'u1', name: 'Alice' }}      // author of new comments
  mentionableUsers={[/* … */]}
  onCommentEvent={(e) => notify(e)}               // host hook, local actions only
>
  <SlidesEditor />
</DeckProvider>
```

`DeckProvider` owns the `ElementStore` (keyed subscriptions), a `NodeRegistry` (element-id → wrapper DOM node, for direct-to-DOM gestures), a `TransientStore` (per-gesture guides/marquee), and the `EditingStore` (which element hosts the single live editor).

## Top-level

- **`SlidesEditor`** — the full editor: header (title, present, PDF, comments), toolbar, filmstrip, interactive stage, notes panel. Owns keyboard shortcuts and the context menu. Pass `readOnly` for a viewer.
- **`SlideStage`** — the central canvas; fit-to-container or an explicit `zoom`. `interactive` turns on selection, gestures, and overlays.
- **`Filmstrip`** — slide navigator with click-to-activate, drag-to-reorder, and an add-slide button.
- **`PresentMode`** — fullscreen playback via a portal; arrow/space/home/end/number-jump navigation and a 150 ms cross-fade.

## Rendering

- **`SlideView` / `ScaledSlide`** — render a slide at native coordinates (e.g. 1280×720); the caller scales with CSS. The same static renderers back the stage, the filmstrip thumbnails, present mode, and PDF export.
- **`ElementView`** — the absolutely-positioned wrapper for one element; subscribes to just that element + the theme, so a single edit re-renders it alone.
- **`StaticRichText`** — walks the rich-text JSON to React nodes for idle text (memoized on doc + theme).
- **`RichTextEditor`** — the single live ProseMirror mount that swaps in over the actively-edited element with identical box metrics. Non-collab keeps a local history and commits on blur; collab binds directly to the element's `Y.XmlFragment`.

## Overlays (inside the scaled stage)

`SelectionOverlay` (handles), `GuidesOverlay` (smart guides + marquee), `CommentBadgesOverlay`, and `RemotePresenceOverlay`. Handle sizes are counter-scaled by `1/scale` so they stay a constant screen size at any zoom.

## Hooks

`useDeck`, `useSlideIds`, `useSlide`, `useSlideElementIds`, `useElement`, `useSelection`, `useActiveSlideId`, `useTheme`, `useEditingId`, `useKeyboardShortcuts`, `useClipboard`, and `useComments`. Each subscribes to exactly the slice it needs.

## Keyboard

Attached to the focused editor wrapper (not `window`). Delete, arrow-nudge (1 / 10 px), ⌘Z / ⇧⌘Z, ⌘D, ⌘C/X/V, ⌘A, ⌘G / ⇧⌘G, Tab to cycle, Enter to edit, ⌘] / ⌘[ (and with Shift, to front/back). While a text editor is focused, only Escape is handled here — the ProseMirror keymap owns the rest.

## Comments UI

`CommentsPanel` lists threads grouped by attachment (with a "No longer attached" section for orphans) and lets you reply / resolve / delete; clicking a thread selects its element. `CommentBadgesOverlay` marks commented elements on the canvas.
