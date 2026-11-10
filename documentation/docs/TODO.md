# Spindle Docs — remaining work & known limitations

Gaps between the current document editor and a full-featured word processor.
Verified against the source (`packages/docs-core`, `packages/docs-react`); the
package READMEs' "Limitations" sections are stale — trust this file.

> **Already shipped** (the READMEs under-report these): true page layout with
> **line-level pagination**, ProseMirror rich text, headers/footers with dynamic
> fields, **tables** (with column *and* row resize, cell background),
> **block + inline images** (upload / drag-drop / URL), links, lists
> (bullet/numbered with nesting), page setup (size/orientation/margins), rulers
> with draggable margins, zoom + auto-fit-to-width, **real-time collaboration**
> (Yjs + offline persistence), and **comments** (threads, replies, resolve,
> @mentions). None of these are TODOs.

## Remaining features

Verified absent in code.

- [ ] **Export — PDF / HTML / Markdown** — none. (DOCX export shipped in 0.4.0:
      `exportToDocx`, plus a Download button in the toolbar.)
- [ ] **Import — HTML / Markdown** — none. (DOCX import shipped in 0.4.0:
      `importFromDocx`, plus an Open button in the toolbar.)
- [ ] **DOCX fidelity gaps** (work lands in `@b32nio/spindle-io`) — `data:` URI images only (remote images are
      skipped with a warning on export), floating/anchored images are skipped
      on import, and unsupported field codes keep no value.
- [ ] **Find & replace** — none.
- [ ] **Word / character count** — none.
- [ ] **Multi-section documents in the UI** — the model supports `sections[]`
      (add/insert/delete), but the editor renders only `sections[0]`
      (`DocumentEditor.tsx`), and collaboration hard-rejects >1 section
      (`collab/y-schema.ts`). Section breaks are effectively unavailable.
- [ ] **Table cell merge / split** — `colspan`/`rowspan` exist in the type and
      schema (parsed from pasted DOM), but there's no merge/split command or UI
      — only column/row resize.
- [ ] **Multi-column (newspaper) layout** — `PageConfig` has no column count;
      layout is single-column only.
- [ ] **Named paragraph styles / style gallery** — a heading dropdown
      (Normal/H1–H3) exists, but no user-facing named-style system despite the
      style-pool infrastructure.
- [ ] **List style options** — bullet vs numbered only; no lettered / roman /
      checkbox markers or per-level marker customization beyond nesting.
- [ ] **Table of contents** — headings exist but nothing consumes them.
- [ ] **Footnotes / endnotes** — none.
- [ ] **Bookmarks / cross-references** — none.
- [ ] **Track changes / suggestions mode** — comments exist; redline/suggestion
      editing does not.
- [ ] **Spell-check** — browser spellcheck is deliberately disabled
      (`input-bridge.ts`); no custom engine.
- [ ] **Equation editor, drawing / shapes / canvas** — none.
- [ ] **Word / line delete gestures** — deleteWord / deleteSoftLine /
      deleteHardLine currently fall back to deleting the selection or a single
      char (`input-bridge.ts`).
- [ ] **Editor context menu** — the true-layout path has none (right-click just
      suppresses the native menu in the legacy path).

## Collaboration gaps

- [ ] **Collaborative undo** — `DocumentHistory` is local-only; no Yjs
      `UndoManager` is wired (`document.ts` notes it's "shadowed by Yjs
      UndoManager once we wire that in"). Undo/redo semantics under
      collaboration are unresolved.
- [ ] **Single-section only under collab** — `setData` is disallowed while
      attached, and IndexedDB persistence is browser-only (`collab/binding.ts`,
      `collab/y-schema.ts`).
- [ ] **Inline images & dynamic fields lost on the collab round-trip** — the
      block↔ProseMirror conversion the collab hydrate/serialize path uses
      (`prosemirror/sync.ts`) has no inline PM node type for an inline image run
      or a `DynamicFieldRun`, so `inlineContentToPm` drops them. Needs new inline
      schema nodes (+ their render paths). Non-collab `getData/setData` is
      unaffected (it passes blocks by reference, never through PM).

## Known limitations (from code)

- **Oversized blocks overflow** — a table or image taller than one page
  overflows rather than paginating; no intra-block splitting
  (`core/true-layout-engine.ts`).
- **Caret / selection geometry** relies on character-ratio estimation fallbacks
  when DOM line data is unavailable (numerous sites in `selection-overlay.ts`,
  `input-bridge.ts`) — a source of caret drift on complex content.
- **Two layout engines coexist** — the recommended `TrueLayoutEditor` and a
  legacy, heuristic engine (`LayoutEngine.ts`, `measurer.ts` — "charsPerLine =
  80", "simplified implementation") are both still exported. Two non-equivalent
  code paths to maintain; the legacy path is a candidate for removal.
- **Comment mark edges** — typing at either edge of a commented range does not
  extend the comment (`prosemirror/schema.ts`); intended, but a UX constraint.
- **Deprecated `HeaderFooter` type** is still exported (`types.ts`) — superseded
  by `HeaderFooterContent`; candidate for removal.
- **Inline images can't be split** at an arbitrary offset — placed before/after
  based on offset (`blocks/utils.ts`).

## Cleanup / tech debt

- [ ] Remove the legacy layout engine and its approximate measurer once the true
      engine is at parity, collapsing to one code path.
- [ ] Drop the deprecated `HeaderFooter` interface.
