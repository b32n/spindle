# @b32nio/spindle-slides-react

React components and hooks for [Spindle Slides](https://spindle.b32n.io) — a full presentation editor over the [`@b32nio/spindle-slides-core`](https://www.npmjs.com/package/@b32nio/spindle-slides-core) engine.

```tsx
import { DeckImpl } from '@b32nio/spindle-slides-core';
import { DeckProvider, SlidesEditor } from '@b32nio/spindle-slides-react';

const deck = new DeckImpl('my-deck', 'Untitled deck');

export function App() {
  return (
    <DeckProvider deck={deck}>
      <SlidesEditor />
    </DeckProvider>
  );
}
```

Includes the interactive stage (selection, move/resize/rotate gestures with smart guides — gestures write straight to the DOM, so a 50-element slide drags at 60 fps with zero React commits mid-move), a toolbar, filmstrip, the single live ProseMirror rich-text editor, table editing (cell selection, per-cell formatting, paste-a-range), present mode with presenter view, real-time collaboration presence, and an element-anchored comments panel. (PDF/PNG export is intentionally left to the host app.)

Compose the whole editor with `SlidesEditor`, or drop down to `SlideStage`, `SlideView`, `Filmstrip`, `PresentMode`, and the `useDeck` / `useElement` / `useSelection` / `useComments` hooks.

## Open / Download (PPTX)

`SlidesEditor` can show **Open** (a local `.pptx` file picker) and **Download** (`.pptx`) in its toolbar, the same host contract as the docs and sheets editors. It is opt-in:

```tsx
<SlidesEditor
  fileActions
  importConfirmation="dialog"                // or 'never', or (report, file) => boolean | Promise<boolean>
  fileName="Quarterly review"                // optional; defaults to the deck title
  onFileOpened={({ fileName, warnings, loss }) => {}}
  onFileError={(error, action) => {}}        // action: 'open' | 'download'
  onExportWarning={(warnings) => {}}         // what the writer left out or changed
/>
```

Opening a file replaces the whole deck (and fails while collaboration is attached). When the file holds something the editor cannot take, a dialog lists it first: what will be lost, what is imported approximately and what is "kept, not editable" (charts, diagrams, objects, media, transitions and animations, saved back unchanged). Files are read locally; nothing is uploaded. For a host-owned button use the exported `exportDeckFile`, `importDeckFile` and `inspectDeckFile`.

Peer dependencies: `react` and `react-dom` (>= 18). Full docs: <https://spindle.b32n.io/docs/slides/components>.

MIT licensed.
