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
      on import, tables of contents keep no field, legacy form fields keep their text only, and content controls that wrap table rows or cells, repeating sections and data-bound controls keep only their content (other fields are kept with their cached result, never recalculated). `inspectDocx(bytes, { detail: 'full' })` and the
      fixture goldens in `fixtures/docx/` list what else is flattened or ignored (sections…). DOCX table formatting (borders, shading, header rows, widths, table styles), table structure (cells of several blocks, nested tables) and
      comments round-trip with replies, resolved state and anchors.
- [ ] **Footnotes and endnotes: remaining gaps** — notes are imported, inserted and edited in the editor, footnotes are
      laid out at the bottom of their page and endnotes collected at the end of each section, and the notes are written
      back. Not done: notes of several paragraphs (they become one paragraph with line breaks), the numbering format and
      restart settings of the file, the endnote position setting (end of document versus end of section), a
      continuation notice on a note that continues, notes in table cells that split across pages (tables are not
      split), and references in headers and footers.
- [ ] **Table of contents: remaining gaps** — a table of contents can be inserted, updated and removed in the editor,
      read from a file as a block with its cached entries and written back as a field Word can refresh. Not done:
      `\t` custom styles, paragraphs with an outline level as sources, levels 7 to 9, updating on its own when
      headings change, the tab stops of the entries (a leader at the margin is used), a table of contents in a table
      cell or header, and indexes and tables of authorities (static text).
- [ ] **Find & replace** — none.
- [ ] **Word / character count** — none.
- [ ] **Sections: remaining gaps** — the editor shows and edits every section
      (Insert → Section break next page/continuous, remove, page setup for
      this section or the whole document, link-to-previous headers/footers).
      Not done: columns and vertical alignment are not laid out, a continuous
      section only shares a page when page size and margins match, there is no
      visible marker for a break, undoing the deletion of a break restores the
      section with the setup of the next one, and a section's setup changes are
      not synced to peers under collaboration.
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
- [ ] **Bookmarks / cross-references** — imported bookmarks and internal links are kept (a bookmark is a point, clicking a
      `#name` link scrolls to it); there is no UI to insert bookmarks or cross-references, and `REF` fields are not recalculated.
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
- [ ] **Multi-section under collab** — sections travel as `section_break`
      nodes and `sectionsMeta`; the shared document is shown by section.
      `setData` is disallowed while attached, and IndexedDB persistence is
      browser-only (`collab/binding.ts`, `collab/y-schema.ts`).

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
