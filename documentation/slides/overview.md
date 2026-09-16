# Spindle Slides

Spindle Slides is a React presentation editor built on a standalone, framework-agnostic engine. Unlike the docs surface — a linear flowing block tree — slides needs positioned elements, z-order, rotation, and groups, so it ships as its own pair of packages rather than reusing the document engine.

- **`@b32nio/spindle-slides-core`** — the engine. Zero React. Owns the deck model, the scene layer (elements, geometry, transforms, snapping, z-order, groups), rich text, themes/layouts, and the optional Yjs collaboration binding.
- **`@b32nio/spindle-slides-react`** — the editor. Hooks + DOM components: the interactive stage, gestures, overlays, toolbar, live rich-text editor, present mode, and PDF export.

## Why a separate engine

The document model is a single flowing sequence of blocks; a slide is a bag of absolutely-positioned elements with independent transforms. Retrofitting z-order, rotation, and groups onto the block tree would compromise both. Slides instead gets a purpose-built **scene layer** that is deliberately container-agnostic: every element references a `containerId` (a slide today) and carries no camera/viewport state, so a future open-canvas surface could reuse the same layer.

## What's in the box

- Text, shape (~18 SVG presets), image, line/connector, and **table** elements
- Tables: per-cell rich text, cell/row/column/range selection, and pasting a spreadsheet range as a table
- Multi-select with move / resize / rotate handles, smart guides + snapping
- Align & distribute, group / ungroup, z-order, filmstrip with drag reorder
- Themes (12 symbolic color slots + major/minor fonts) and a layout gallery with placeholder materialization
- Rich text per element via ProseMirror (a single live editor mounts on the element being edited)
- Present mode (fullscreen, letterboxed, cross-fade) with presenter view (timer, notes, next-slide preview); PDF export is wired up by the host app
- Real-time collaboration over Yjs with element-anchored comments
- A pure-JSON data model — `getData()` / `setData()` — that's easy to author by hand or generate

## Quick start

```tsx
import { DeckImpl } from '@b32nio/spindle-slides-core';
import { DeckProvider, SlidesEditor } from '@b32nio/spindle-slides-react';

const deck = new DeckImpl('my-deck', 'Untitled deck');
// deck.setData({ ... })  // optional: hydrate from JSON

export function App() {
  return (
    <DeckProvider deck={deck}>
      <SlidesEditor />
    </DeckProvider>
  );
}
```

The engine is instance-owned and passed in as a prop, so the same deck can be shared across panes or driven by collaboration. See [Architecture](/docs/slides/core/architecture) for the internals, [Collaboration](/docs/slides/core/collaboration) for real-time editing, and [Data Structures](/docs/slides/data-structures) for the JSON model.
