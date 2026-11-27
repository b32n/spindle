# Spindle Slides — remaining work & manual QA

Gaps between the current slides editor and a full-featured presentation tool.
The scene model already carries what most deferred items need, so several are
"UI/renderer only" rather than model changes.

> **Recently shipped** (so this list stays honest): **tables** are now a
> first-class element type — per-cell rich text, row/column/range **cell
> selection**, multi-cell formatting, content-driven row heights, and
> **paste-a-spreadsheet-range-as-a-table**. **Presenter view** also ships
> (elapsed timer, speaker notes, next-slide preview, toggled with `S`).
> Neither is a TODO anymore.

## Remaining features

Not in the current scope; verified absent in code.

- [ ] **PPTX import / export** — the schema deliberately keeps `layoutRef`,
      placeholder `{type, idx}`, symbolic theme colors, and flat text bodies so
      this stays honest, but no reader/writer exists yet. A corpus of test decks with a
      construct manifest (`fixtures/pptx/`) now defines what the importer has to cover. A loss detector
      (`detectPptxLoss` in `@b32nio/spindle-io/pptx`, loss ids in `documentation/import-loss.md`) and a golden runner
      (`npm run slides:pptx-fidelity`) already report what every corpus deck holds; both say "not imported" until the
      importer lands. Nothing opens a `.pptx` yet.
- [ ] **Charts** as a first-class element type (bar / line / pie, …). No chart
      element or renderer exists.
- [ ] **Video / audio** elements — absent from the `SlideElement` union.
- [ ] **Element animations & a slide-transition library** — present mode ships
      only a single 150 ms cross-fade; there are no per-element animations and
      no transition variety.
- [x] **Gradient fills & image/picture fills** — `Fill` has `gradient` (linear, radial, path; stops with alpha),
      `picture` (stretch or tile, alpha; data URLs only) and `pattern` (rendered as a solid blend of its two colours)
      next to `none` and `solid`, on shapes, text boxes, table cells and slide backgrounds. The fill controls offer
      gradient presets and an image fill from a file picker. Still open: a gradient stop editor, a pattern picker,
      layout backgrounds in the UI, tile offset/mirroring and picture-fill cropping (`a:srcRect`).
- [ ] **Nested groups** — v1 groups are flat via a shared `groupId`
      (`scene/group.ts`); no group hierarchy.
- [ ] **Master-slide editing UI** — the model supports masters/layouts and
      materializes placeholders, but there's no surface to edit a master/layout.
- [ ] **Image host-upload hook** — images inline as data URLs (size-unbounded);
      no callback for a host to upload and swap in a CDN URL
      (`ImageFormatBar.tsx`, `Toolbar.tsx`).
- [ ] **Free-form image crop** — only object-fit modes (fill/contain/cover);
      no region crop (`scene/types.ts`).
- [ ] **Text autofit / shrink-to-fit** — there is no autofit anywhere in text
      layout: text that overflows its box is clipped (and with middle/bottom
      vertical alignment the overflow clips at the *top*, dropping leading
      lines). No `bodyStyle` autofit flag and nothing scales the font size to fit
      (`text/` layout, `elements/*Text*`).
- [ ] **Rich speaker notes** — notes are plain-text only (`NotesPanel.tsx`).
- [ ] **Real PDF / PNG export backend** — intentionally **not** in these
      packages; export is delegated to the host app (`SlidesEditor.tsx`). A
      shared, headless export backend is a possible follow-up.

## Collaboration gaps

- [ ] **Table-cell text is not collaborative** — element rich text binds to the
      CRDT, but table **cell** rich text commits via snapshot with no Yjs
      binding (`RichTextEditor.tsx`). Concurrent edits to different cells of the
      same table are not CRDT-merged.
- [ ] **In-text remote carets** — presence is element-level only; per-character
      remote cursors need per-fragment cursor handling that `y-prosemirror`'s
      `yCursorPlugin` doesn't support with per-element fragments
      (`RemotePresenceOverlay.tsx`).

## Known limitations (accepted v1 constraints)

- Scalar element props sync last-write-wins (`collab/y-schema.ts`).
- No negative / flipped sizes in transforms (`scene/transform.ts`).
- Uniform inner text padding, not per-side (`scene/types.ts`).

## Manual QA checklist

No React test infrastructure exists in this repo by design; `examples/slides-demo`
and the website demo are the smoke tests. Run through this before shipping
editor changes:

- [ ] **Drag performance** — with ~50 elements, dragging stays at 60 fps with no
      React commits mid-move (transforms update on the DOM nodes; the tree
      doesn't re-render until pointerup).
- [ ] **Resize / rotate** — the opposite anchor stays fixed at several rotations;
      Shift aspect-locks; rotation snaps near 45° steps.
- [ ] **Snapping** — smart guides appear against other elements' edges/centres
      and the slide centre; the box snaps within threshold.
- [ ] **Rich text** — double-click / Enter enters edit with no visual jump;
      bold/italic/color/size/align/list work live and on an idle selection;
      **IME** composition is never interrupted; test the caret at
      **25 / 50 / 100 / 200 / 400 %** zoom.
- [ ] **Tables** — cell / row / column / range selection; multi-cell formatting;
      add/remove row·col; paste a spreadsheet range as a table (HTML and TSV
      clipboard); content-driven row heights.
- [ ] **Z-order / align / group** — front/back/forward/backward, align &
      distribute, group/ungroup, and undo of each.
- [ ] **Layouts & themes** — the new-slide gallery materializes placeholders; the
      theme picker recolors accents live; slide backgrounds apply.
- [ ] **Present mode** — fullscreen, letterboxed at 16:9, arrow/space/home/end/
      number-jump nav, cross-fade; **presenter view** (`S`) shows timer, notes,
      next-slide preview.
- [ ] **PDF export** — exercised in the **host/website** demo (export lives
      there, not in this package): one page per slide at the correct size; fonts
      and images render (test Chrome and Firefox — Firefox has `@page size`
      quirks).
- [ ] **Collaboration** — run the demo in two-pane and `?ws=` modes: concurrent
      moves/reorders/text-in-different-elements converge; undo reverts only the
      local user's edits; remote presence outlines appear; comments sync;
      orphaned threads survive element deletion. (Note: table **cell** text is
      not yet CRDT-merged — see Collaboration gaps.)
