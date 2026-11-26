# Changelog

All notable changes to Spindle are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project
follows [Semantic Versioning](https://semver.org/) (pre-1.0: minor versions may
contain breaking changes). All `@b32nio/spindle-*` packages are versioned in
lockstep.

## [Unreleased]

### Added

- **`@b32nio/spindle-io`: new optional package for XLSX and DOCX import/export.** The readers/writers
  (`exportXlsx`/`importXlsx`/`detectXlsxLoss`, `exportDocx`/`importDocx`/`detectDocxLoss`) and their
  dependencies (`exceljs`, `docx`, `jszip`, `fast-xml-parser`) moved out of `@b32nio/spindle-shared`, so
  Phase 1 fidelity work grows there instead of in shared's always-on path. Subpaths `@b32nio/spindle-io/xlsx`
  and `/docx` load one format only; the root re-exports both plus the CSV helpers. See `packages/io/README.md`.
- **Full DOCX loss report.** `detectDocxLoss(bytes, { detail: 'full' })` (`@b32nio/spindle-io/docx`) and
  `inspectDocx(bytes, { detail: 'full' })` (`@b32nio/spindle-docs-core`) additionally report ten new ids for
  content the importer flattens or ignores: `docx.sections`, `docx.namedStyles`, `docx.tableFormatting`,
  `docx.tableStructure`, `docx.paragraphFormatting`, `docx.listFormatting`, `docx.bookmarks`,
  `docx.headerFooterVariants` and `docx.sectionProperties`, plus a long tail (`docx.trackedMoves`,
  `docx.formattingRevisions`, `docx.dropCaps`, `docx.watermarks`, `docx.pageBorders`, `docx.lineNumbers`,
  `docx.hyphenation`, `docx.tabStops`, `docx.rtl`, `docx.ruby`, `docx.embeddedFonts`,
  `docx.altChunk`, `docx.legacyForms`), and `docx.fields` lists its field kinds in `where`.
  The default report (`detail: 'summary'`), and so the import-confirmation dialog, is unchanged apart from the
  count fixes below. The new detectors parse the XML instead of matching text.

- **Section breaks in the ProseMirror document and in collaboration.** `docsSchema` gains a top-level
  `section_break` node (attribute `sectionId`) that ends a section; the last section has none, so a document without
  breaks is one section as before. `documentToProseMirror` and the new `sectionsToPmDoc` write the breaks, and
  `proseMirrorToDocument` reads them back into `sections` (a document without breaks still gives
  `existingDoc.sections[0]`). The shared Yjs document now carries every section: `hydrateYDocFromData` no longer
  rejects more than one section, and `serializeYDocToData`/`getData` under collaboration return them with their page
  setup, header, footer, numbering and vertical alignment. Existing single-section Y documents need no migration.
  The editor still draws one flow, so a multi-section document attached to a room shows as one run until the section
  editing work lands.
- **`getCollabBlockers(document)`** (`@b32nio/spindle-docs-core`) lists content the shared layout cannot hold (no
  sections, duplicate section ids, unknown block or inline types), and `attachCollab` now rejects such a document
  with a `CollabBlockedError` that carries the list, so nothing is dropped silently. It is empty for the documents
  the importers produce today.
- **Collaboration schema version.** The shared document records `meta.schemaVersion` (`COLLAB_SCHEMA_VERSION`, now 2;
  a missing value means 1). `attachCollab` rejects with a `CollabSchemaVersionError` when the room was written with a
  newer version than the client supports, before seeding or editing anything.
- **Layout across sections (engine).** `computeSectionedLayout` (`@b32nio/spindle-docs-react`) paginates a document of
  several sections: each section is laid out with its own page size, orientation and margins, starts on a new page, on
  the same page (continuous, when the page setup is the same) or on the next even or odd page (a blank page is added
  when needed), restarts or continues the page number and carries its number format, and each page records which header
  and footer variant (first page, even page, default) it shows. `PageLayout` gains the optional `sectionIndex`,
  `pageConfig`, `top`, `pageNumber`, `pageNumberFormat`, `headerFooterVariant` and `isBlank`, `getPageY` and
  `findPageAtY` follow pages of different heights, and `formatPageNumber`, `resolveInheritedContent` (a section with no
  header or footer shows the previous section's) and `proseMirrorToSectionedFlowBlocks` are exported. A layout of one
  section is the same as before. The editor does not use it yet.
- **DOCX: how a section starts is kept.** `w:type` of a section (continuous, even page, odd page; a new page is the
  default and is not kept) is imported into the additive `DocSection.breakType` / `Section.breakType` and written back,
  and travels through `getData`/`setData` and the shared document.
- **Editor: every section of a document is shown.** `DocumentEditor` passes all sections to `TrueLayoutEditor`
  (new `sections` and `onSectionsChange` props, `TrueLayoutSection`), which lays them out with
  `computeSectionedLayout`: each page has its own size, orientation and margins (pages of different widths are centred),
  each section is measured at its own content width, and headers and footers follow the section (a section with none
  shows the previous one's) including first-page and even-page variants and the page number restart and format. Edits in
  any section are written back to that section, and deleting a section break merges the section into the next one. The
  rulers and page setup act on the section of the active page, and editing a header or footer edits the section that owns
  it (`onHeaderChange`/`onFooterChange` get the section id as a second argument). Not laid out yet: columns, vertical
  alignment, and per-fragment margins for continuous sections.
- **Sections can be inserted, removed and set up in the editor.** The toolbar has a Section break menu (next page,
  continuous, remove): a break ends the section before it (the content before it becomes a new section with a copy of the
  setup, the section after keeps its own and starts the way the break says and inherits the header and footer). New
  `insertSectionBreak(sectionId)` and `removeSectionBreak()` commands (`createCommands`) and
  `DocumentImpl.splitSection(sectionId, newSectionId, breakType?)`, which also adds the new section to the shared document
  while collaboration is attached. Page setup asks whether it applies to this section or the whole document (the modal's
  `onConfirm` gets the scope as a second argument), and the header/footer editor of a later section has "Link to
  previous" (`onHeaderFooterLinkChange`). `docx.sections` is removed from the full loss report, since sections are now
  shown and editable.
- **DOCX comments are imported.** `importDocx` reads `word/comments.xml`, `commentsExtended.xml` (replies and the
  resolved state) and `people.xml` (author ids), anchors each thread to its commented text across runs, links,
  paragraphs and table cells, and returns the threads in the additive `DocumentImportResult.comments`
  (`DocCommentThread`, `DocComment` in `@b32nio/spindle-shared/export/docx`); the anchored runs carry the additive
  `commentThreadId` on `DocRunBase`. `importFromDocx` fills the document's comment store, so threads appear in the
  existing comments sidebar with author, initials (additive `Comment.authorInitials`), date, replies and resolved
  state. Overlapping comments are tagged with the innermost thread (all threads keep their anchor quote), a comment
  with no surviving anchor stays as a thread without one, and damaged comment parts import what they hold plus a
  warning.
- **DOCX comments are exported.** `exportDocx` accepts the additive `DocumentExportInput.comments` and writes
  `word/comments.xml`, the range markers and reference runs, `commentsExtended.xml` (reply parentage and the
  resolved state, on every comment of a resolved thread) and `people.xml`, with content types and relationships.
  `exportToDocx` passes the document's threads and the runs' `commentThreadId`. Author, initials, date and reply order
  survive a round trip, and a thread with no surviving anchor is written as a collapsed range at the start of the
  first body paragraph so it is not lost. Comment dates are written without fractional seconds; the placeholder date of
  an undated comment is omitted. LibreOffice reads the result (comments, replies, resolved state).

- **DOCX styles are resolved into the formatting you see.** `importDocx` now reads `word/styles.xml` fully:
  `docDefaults`, the default paragraph style, `basedOn` chains (loops and missing bases are survived), paragraph and
  character styles (`pStyle`/`rStyle`, including the `Hyperlink` style on links), toggle properties (bold, italic,
  small caps, strike: overridden inside one chain, XOR-ed between the paragraph and character style, set outright by
  direct formatting), numbering carried by a paragraph style (`numId` 0 switches it off), and paragraph alignment,
  spacing and indents through the chain. The effective values are written into the existing run and paragraph fields,
  so runs and paragraphs gain formatting they previously lost (paragraph spacing from the defaults, heading sizes and
  colours, link colour and underline, callout styles). A heading is recognised by its `Heading N` name, else by an
  outline level in its style chain (level 10 means body text), else through a `Heading N` ancestor; a heading numbered
  through its style stays a heading.
- **DOCX theme fonts and colours are resolved.** `importDocx` reads `word/theme/theme1.xml` (major/minor Latin fonts and
  the colour scheme) and resolves `asciiTheme` fonts, `themeColor` text colours and `themeFill` shading with their
  `themeTint`/`themeShade` modifiers, in `docDefaults`, styles and direct formatting. A theme reference takes precedence
  over the literal value next to it; the literal is kept when there is no theme or the slot is missing, and when Word's
  recorded colour is within a few steps of the computed tint/shade it is used as the exact value.

- **DOCX paragraph, run and list properties.** Import and export now carry line spacing (single, 1.5, double,
  any multiple, at least, exactly), left, right, first-line and hanging indents, keep-with-next, keep-lines-together
  and page-break-before, all capitals (`w:caps`), hidden text (`w:vanish`), the high-ANSI font as a fallback for the
  Latin font, and East Asian and complex-script fonts (kept only where they differ from the Latin font). Numbered lists
  keep their first number, continue counting after an interruption and restart where Word restarts them (the
  additive `listItem.start`); every separate numbered list in an exported file now restarts instead of carrying on
  counting from the previous list. Interchange additions (all optional): `DocParagraphStyle.rightIndentPt`,
  `lineSpacingCustom`, `keepWithNext`, `keepTogether`, `pageBreakBefore`, `DocRunBase.allCaps`, `hidden`,
  `fontFamilyEastAsia`, `fontFamilyComplex`, `listItem.start`.
- **Editor: paragraph formatting is part of the editor document.** `ParagraphBlock`, `HeadingBlock` and
  `ListItemBlock` share the new `ParagraphFormat` fields (`lineSpacing`, `spaceBefore`, `spaceAfter`, `leftIndent`,
  `rightIndent`, `firstLineIndent`, `keepWithNext`, `keepTogether`, `pageBreakBefore`), `TextStyle`/`TextRun` gain
  `allCaps`, `hidden`, `fontFamilyEastAsia` and `fontFamilyComplex`. The ProseMirror schema carries them as optional
  attrs on `paragraph` and `heading` and as new `allCaps` and `hidden` marks, so they survive editing and collaboration
  (they are shared through the Yjs document like every other attr and mark). The layout honours left, right and
  first-line indents, spacing, multiple line spacing, small caps, all caps and numbered-list start values; keep-with-next,
  keep-lines, page-break-before and `atLeast`/`exactly` line heights are stored and exported but not laid out yet, and hidden
  text is shown dimmed rather than removed (the caret and selection mapping needs every character painted).

- **DOCX named styles are kept and exported.** Import keeps the ids, names, `basedOn`/`next`/`link`, default flag,
  outline level and own run/paragraph properties of the paragraph and character styles a document uses (and the styles
  they refer to) in the additive `DocumentImportResult.namedStyles` (`DocNamedStyle`), and tags blocks and runs with
  the additive `namedStyleId` (`DocParagraphStyle`, `DocRunBase`). `exportDocx` accepts
  `DocumentExportInput.namedStyles` and writes the styles into `styles.xml` (a built-in style of the same id, such as
  `Heading1` or `Hyperlink`, is replaced by the document's definition) with `pStyle`/`rStyle` references, so Word shows
  "Callout", "Accent Word" and the like in its style list. A reference to a style the document no longer defines is
  left out of the file. Word's implicit `ListParagraph` on list items and `Hyperlink` on links are not kept as named
  styles (the writer applies them again).
- **Editor: named styles in the document model.** Additive `DocumentData.namedStyles` (`NamedStyle`),
  `DocumentImpl.getNamedStyles()` / `setNamedStyles()`, `namedStyleId` on `ParagraphFormat` (paragraph, heading and
  list-item blocks), `TextRun` and `InlineLink` (distinct from the existing style-pool `styleId`). In ProseMirror the
  paragraph and heading nodes gain a `namedStyleId` attr and there is a new `namedStyle` mark, and in the Yjs
  document a new `namedStyles` map, so styles and references sync in collaboration (a late joiner receives the room's
  styles; no migration is needed). There is no style gallery or style editing UI yet.

- **DOCX table formatting is imported and exported.** Import keeps table and cell borders (line style, width, colour), cell
  shading, vertical alignment, repeating header rows (`w:tblHeader`), `w:cantSplit`, column widths (`w:tblGrid` and `w:tcW`),
  the table style id and `w:tblLook`. A table style's conditional formatting (header and total rows, first and last column,
  banded rows and columns, corner cells, resolved through the style's `basedOn` chain and the table's `tblLook`) is applied to
  the cells it reaches: shading, borders, vertical alignment and text properties land on the cells and their runs, below the
  paragraph and run formatting. Table styles are kept in `namedStyles` (`kind: 'table'`, with the additive `table` member holding
  borders, whole-table and per-`tblStylePr` formatting) and written back to `styles.xml`; the export writes borders, shading,
  header rows, `cantSplit`, vertical alignment, widths, the table style and `tblLook`. `docx.tableFormatting` is removed from the
  full loss report.
- **Editor: table formatting in the document model.** Additive `TableBorders`, `TableLook` and `CellVerticalAlign` types;
  `borders`, `look` and `namedStyleId` on `TableBlock`, `borders` and `verticalAlign` on `TableCell`, `repeatHeader` and
  `cantSplit` on `TableRow` (matching attributes on the ProseMirror table nodes, so they sync over Yjs), and
  `NamedStyle.table`. The page painter and the measurer draw cell borders, vertical alignment and cell widths; a table with no
  `borders` keeps the default grid and a table whose `borders` is `{}` shows no lines. Repeating header rows and `cantSplit` are
  stored and exported; tables are not split across pages yet, so the editor does not act on them.
- **DOCX bookmarks and internal links are imported and exported.** A `w:bookmarkStart` becomes the additive `DocBookmarkRun`
  interchange member (`{ type: 'bookmark', name }`; `_Toc` and `_Ref` names kept, Word's `_GoBack` marker dropped), a
  `w:hyperlink` with `w:anchor` becomes a link whose `href` is `#name` (`url#name` with a target too), and the export writes
  `w:bookmarkStart`/`w:bookmarkEnd` and an internal hyperlink. A bookmark is kept as a point where it starts (its range is
  not preserved). Clicking an internal link in the editor scrolls to its bookmark. `docx.bookmarks` is removed from the full
  loss report.
- **Editor: bookmark in the document model.** Additive `InlineBookmark` in `InlineContent` and a `bookmark` inline atom in the
  ProseMirror schema (attr `name`) that syncs over Yjs. It takes no space on the page but counts as one position. Code that
  switches exhaustively over `InlineContent` or `DocInline` needs a `bookmark` case.
- **DOCX fields are kept with their cached result.** Import keeps every field other than `PAGE`/`NUMPAGES` (still page-number
  fields), `HYPERLINK`, tables of contents and legacy form fields as the additive `DocFieldRun` interchange member
  (`{ type: 'field', instruction, kind, result, dirty? }`): `REF`, `PAGEREF`, `SEQ`, `STYLEREF`, `DATE`, `TIME`, `TITLE`,
  `AUTHOR`, `NUMWORDS`, `IF`, `MERGEFIELD`, `SECTIONPAGES` and the rest, complex and `fldSimple`. The result is the text (or
  formatted runs) Word last computed; it is displayed as is and never recalculated, and the export writes it back as a
  complex field (begin, instruction, separate, result, end, `w:dirty` kept) so Word can refresh it. A field inside another
  field's result is part of that result. `HYPERLINK "url"` fields become links. A table of contents or index keeps its text as
  static paragraphs (the field itself is not kept), legacy form fields (`FORMTEXT`, `FORMCHECKBOX`, `FORMDROPDOWN`) still lose
  their value, and fields in headers and footers other than the page-number ones are still dropped with a warning.
- **Editor: generic field in the document model.** Additive `InlineField` in `InlineContent` (`DocumentData` blocks and table
  cells), a new `field` inline atom in the ProseMirror schema (attrs `instruction`, `kind`, `result`, `dirty`) that syncs
  over Yjs, and `docx.fields` now counts only what is not kept as a field (tables of contents, indexes, legacy form
  fields). Code that switches exhaustively over `InlineContent` or `DocInline` needs a `field` case.

- **DOCX table structure is imported and exported.** A table cell keeps all of its content: several paragraphs, lists,
  headings, images and tables nested in the cell (to a depth of 20; the corpus goes three deep) are read into the additive
  `DocTableCell.blocks`, with `content` still holding the cell's first paragraph that has text for code that does not read
  `blocks` (a cell of one plain paragraph has no `blocks`). The export writes the blocks back, ends a cell that finishes with a
  nested table with an empty paragraph as the file format requires, and numbers a list in a cell on its own. Merged cells are read
  more carefully: a vertical merge is extended once per row however many continuation cells sit under it, ends in a row that
  does not continue it, and a continuation cell with no merge above it keeps its content as an ordinary cell; a row that begins
  after skipped grid columns (`w:gridBefore`) gets an empty cell without lines in front so its cells stay in their columns. A
  rotated cell (`w:textDirection` `btLr` or `tbRl`) is recorded as the additive `textDirection` and written back (the editor
  still shows the text horizontally). The conditional text formatting of a table style no longer reaches a table nested in one
  of its cells. `docx.tableStructure` is removed from the full loss report.
- **Editor: cells of several blocks and nested tables.** Additive `TableCell.blocks` and `TableCell.textDirection` and a
  `textDirection` attribute on the ProseMirror cell nodes; they sync over Yjs like the other cell attributes (the cell nodes
  already held any block). The page painter and the measurer draw the paragraphs, list items (with a bullet or number),
  images and nested tables of a cell, with the document positions of each so clicking and selecting work inside a nested table.
  The table resize handles and the context menu act on the table under the pointer, not on the tables nested in it.

- **DOCX header/footer variants and section properties are imported and exported.** A section's headers and footers are
  chosen by `w:type` (default, first, even) instead of taking whichever reference comes first: `w:titlePg` becomes
  `differentFirstPage` with `firstPageParagraphs`, and the document's `w:evenAndOddHeaders` setting becomes the additive
  `DocHeaderFooter.differentOddEven` with `evenPageParagraphs`. Each section also keeps its page number restart and format
  (`w:pgNumType`, the additive `DocSection.pageNumbering`), vertical alignment (`w:vAlign`, `verticalAlign`) and text columns
  (`w:cols`, `columns`: count, spacing, separator, unequal widths). The export writes all of these, including first-page and
  even-page header and footer parts and the even/odd setting (a document written with a first-page header used to get no first-page
  part). `docx.headerFooterVariants` is removed from the full loss report, `docx.sectionProperties` now only counts chapter
  numbers in page numbers, and the effect of `docx.columns` reads "kept for export; laid out as a single column".
- **Editor: header/footer variants and section properties in the document model.** Additive `HeaderFooterContent.differentOddEven`
  and `evenPageBlocks`, `PageConfig.columns` (`PageColumns`), `Section.pageNumbering` (`PageNumbering`) and
  `Section.verticalAlign`, kept by `getData`/`setData` and by the shared Yjs document (`sectionsMeta`), so collaboration carries them.
  Nothing is laid out from them yet: the editor draws the first-page header and footer as before, and even-page headers, page
  number restarts and formats, vertical alignment and columns are stored and exported only.

- **DOCX content controls are kept, edited (check boxes) and written back.** A Word content control (`w:sdt`) of type rich
  text, plain text, drop-down list, combo box, date, check box, picture or building block (cover page, table of contents
  wrapper) is imported as a wrapper around the content it holds, at the block level (around paragraphs, lists, tables and
  other controls) and inside a paragraph, nested to any depth, with its title (`w:alias`), tag, lock, placeholder, list
  items, date format and value, and check box state. The export writes it back as a `w:sdt` with its properties. Additive
  interchange types `DocContentControl`, `DocContentControlRun` (a `DocInline` member, `type: 'contentControl'`) and a
  `DocBlock` member of the same type, and in docs-core `ContentControl`, `InlineContentControl` and `ContentControlBlock`
  (block type `'content-control'`). `getPlainText` reads through inline controls. Widening these unions can break a host's
  exhaustive `switch` at compile time.
- **Content controls in the ProseMirror document and in collaboration.** `docsSchema` gains a block node `content_control`
  (top level, table cells and other controls) and an inline node `content_control_inline`, with the control's properties as
  attributes; both are plain ProseMirror nodes, so they sync over Yjs like any other content and
  `getCollabBlockers` knows them (it looks inside them for types it does not know). The collaboration `schemaVersion`
  (`COLLAB_SCHEMA_VERSION`) is now 3: a client from before this change refuses to join a room written with it, and rooms
  written before it need no migration. Pasted HTML that carries a control (`data-content-control`) is validated, not trusted.
  New commands `toggleCheckboxControl(pos?)` and `findCheckboxControl`.
- **Editor: content controls.** The editor draws the content of a control as before, with a title chip in the left margin
  next to the first line of a block control (its title, tag or kind) and a faint tint and tooltip on the content of an inline
  one; placeholder text is grey. Clicking the glyph of a check box control toggles it (its state and glyph). Other control
  types are shown but have no picker yet, and a control's lock is kept and written back but not enforced.
- **DOCX legacy form fields keep their text.** `FORMTEXT`, `FORMCHECKBOX` and `FORMDROPDOWN` fields now import the text they
  display as plain text (it was dropped, with a warning). The form behaviour is still lost and is reported as
  `docx.legacyForms`.

- **DOCX footnotes and endnotes are imported, shown and written back.** A footnote or endnote reference becomes
  an inline reference that holds the text of its note (formatting, links, pictures, line breaks and fields), with the
  note's custom mark if it has one; the note parts (`footnotes.xml`, `endnotes.xml`, their relationships and the
  separator settings) are written on export, with Word's separator notes. Additive interchange type `DocNoteRefRun`
  (a `DocInline` member, `type: 'noteRef'`) and, in docs-core, `InlineNoteRef` (inline content `'noteRef'`). A note of
  several paragraphs is read as one paragraph with line breaks between them, and a table, list or picture-only structure
  in a note is not kept (reported as a warning for tables). Widening these unions can break a host's exhaustive `switch`
  at compile time. The loss report no longer counts the notes it imports: `docx.footnotes` and `docx.endnotes` now only
  count notes that nothing in the body refers to (or whose only reference is in a text box), and notes in a part that
  cannot be read.
- **Note references in the ProseMirror document and in collaboration.** `docsSchema` gains an inline atom `note_ref`
  (attributes `kind`, `noteId`, `customMark`) whose content is the note's text (only content a note can hold, so notes
  do not nest). It is a plain ProseMirror node, so it syncs over Yjs and `getCollabBlockers` knows it. The collaboration
  `schemaVersion` (`COLLAB_SCHEMA_VERSION`) is now 4: a client from before this change refuses to join a room written
  with it, and rooms written before it need no migration.
- **Editor: notes.** A reference is drawn as its number, a superscript (digits for footnotes, lower-case roman
  numerals for endnotes, numbered per kind in document order, or the custom mark). Footnotes are laid out at the
  bottom of the page their reference is on, under a separator; endnotes are collected in an area after the last block of
  each section. Clicking a reference scrolls to its note and clicking a note's number goes back to its reference.
- **Footnotes at the bottom of the page (layout).** The True Layout engine places a line (or a block that is not split,
  such as a table) on a page only if it fits together with the footnotes the page then carries; the line moves to the next
  page with its notes when it does not. A page gives at most half of its content height to notes: a note that does not fit
  continues on the next page, under a full-width separator, and notes still waiting after the last page get pages of
  their own. New in `@b32nio/spindle-docs-react`: `LayoutOptions.footnotes` / `SectionedLayoutOptions.footnotes`
  (`FootnoteHooks`) and `initialNotes`, `PageLayout.notes` / `.footnotes` / `.notesOnly`, `createFootnoteHooks`,
  `footnoteRefs`, `applyPageNotes` and `FootnotePlan`; a layout made without the hooks is unchanged. A continuous section
  that shares a page also shares its notes. `TrueLayoutEditor` has a `footnotesAtPageBottom` prop (default true; false
  puts the footnotes in the notes area at the end of each section as before).
- **Editor: insert and edit notes.** The toolbar has a Footnote menu (Insert footnote, Insert endnote, shown when the host
  `Toolbar` gets `onInsertNote`; `DocumentEditor` wires it) that puts a reference at the caret and opens a small editor
  over the new note's entry in the notes area; clicking a note's text edits it the same way, keeping its formatting, links
  and line breaks (Enter is a line break; Escape or clicking away ends the edit, as one undoable step). Deleting a
  reference removes its note and renumbers the rest, since numbers follow from where the references are. New
  `TrueLayoutEditorHandle.insertNote(kind)` and, in docs-core, `insertNoteReference`, `findNoteRef`, `noteOrdinalAt`,
  `nextNoteId` and `noteTextFragment`.

- **Table of contents.** The editor has a Table of contents menu in the toolbar (Insert, Update, Remove; shown when
  `Toolbar` gets `onTableOfContents`, which `DocumentEditor` wires). Insert puts a table of contents after the block the caret is
  in; it lists the headings of the levels its instruction asks for (`TOC \o "1-3" \h \z \u` by default: levels 1 to 3, with
  links, page numbers shown) with a dotted leader and the page number at the right, an entry per heading, each linking to a
  `_Toc` bookmark that the heading gets when it has none (an imported heading keeps the bookmark it has). The page numbers are
  the ones the layout gives, in the format of the page (so a section with roman page numbers lists roman numerals). Update writes
  every table again from the headings and the pages they are on, as one undoable step; nothing changes a table on its own, so a
  table in a file keeps its numbers until it is updated. Remove deletes the table at the caret (the nearest one). Clicking an
  entry goes to its heading (the internal link of the bookmark). Not acted on yet: switches other than `\o`, `\h` and `\n`
  (`\t` styles, `\b`, `\f`, `\p`) stay in the instruction, headings are the only source (not paragraphs with an outline level
  or custom styles), and nothing lists headings of levels 7 to 9.
- **Table of contents in the model, in the ProseMirror document and in collaboration.** Additive block type `TocBlock`
  (`type: 'toc'`: `instruction` and the entries as `blocks`, one paragraph each, with the style `TOC1` to `TOC6`) and `docsSchema`
  gains a `toc` node (an atom holding its entry paragraphs; at the top level and in content controls, not in cells or list
  items), plain ProseMirror content that syncs over Yjs; `getCollabBlockers` knows it. The collaboration `schemaVersion` is now 5.
  New commands `insertTableOfContents`, `updateTableOfContents` and `removeTableOfContents`, `parseTocInstruction`, and
  `TrueLayoutEditorHandle.tableOfContents(action)`.
- **DOCX tables of contents are read and written back as fields.** A Word table of contents (a `TOC` field that starts in the first
  entry paragraph and ends after the last, in the body or inside a content control such as Word's Table of Contents building
  block) is imported as a table of contents block with the entries it showed, unchanged: links to the `_Toc` bookmarks, tabs, the
  `PAGEREF` fields with their cached page numbers, and the entry paragraph styles. Nothing recomputes them until the user updates
  the table. The export writes the field again: begin, instruction and separator in the first entry, each entry with a right tab
  and dotted leader at the margin, and the end in a paragraph of its own; an entry made in the editor (a link to a heading, a tab
  and a number) gets its number as a `PAGEREF` field of the heading's `_Toc` bookmark so Word can refresh the table (F9, or
  Update Field). No `updateFields` setting is written (it makes Word ask on open). Additive interchange block member
  `DocBlock` `{ type: 'toc'; instruction; blocks }`. A table of contents whose field does not end in the same body or control,
  has a table or a section break inside it, or is a simple field stays static text and is still reported. Entry paragraphs are read
  as plain paragraphs whatever their style; a link around the tab and the number is split at the tab.
- **Layout: a tab with a leader, and measuring a tab.** A paragraph can ask for a tab to be drawn as a dotted leader up to the text
  after it at the right (`ParagraphAttrs.tabLeader`; used by the entries of a table of contents), and `LinkRun.plain` draws a link
  like the text around it. The measurer treated a tab and the text after it as two pieces on the same line and gave the text twice;
  it is one piece now (this also affects tabs in imported text).

### Changed

- **DOCX loss report: `docx.contentControls` counts only what is still lost** (repeating sections, data-bound, equation,
  citation, bibliography and group controls, and controls wrapping table rows or cells, whose content is kept without the
  control), by reason in `where`, in both detail levels; it counted every `w:sdt` before. `docx.fields` no longer counts the
  legacy form fields (see `docx.legacyForms`), and no longer counts a table of contents the importer reads as a block (it counts one that stays static text, and indexes). The DOCX fidelity harness now counts the text inside content controls.
- `@b32nio/spindle-sheets-core` and `@b32nio/spindle-docs-core` now depend on `@b32nio/spindle-io` and load it
  lazily on first XLSX/DOCX open or save (public APIs unchanged).
- `@b32nio/spindle-shared` no longer depends on `exceljs`, `docx`, `jszip` or `fast-xml-parser`.
- All packages bumped to 0.4.1 (lockstep; 0.4.1 is the first release containing `@b32nio/spindle-io`).

### Deprecated

- `@b32nio/spindle-shared/export/xlsx` and `/export/docx` functions: they keep working as async shims that
  load `@b32nio/spindle-io` on demand (now an optional peer dependency, so install it if you use them).
  Import from `@b32nio/spindle-io/xlsx` / `/docx` instead. Types and import-loss helpers stay in shared;
  `@b32nio/spindle-shared/export/csv` is unchanged.

### Fixed

- **DOCX loss report: `docx.namedStyles` is gone** (the style names are now imported and exported).
- **A paragraph's `indent` is a level of 24 px, not pixels.** The DOCX importer stored pixels in `indent` (the editor's
  indent command stores levels), so an imported indent of 48 px was laid out as 48 levels. The importer now sets
  `leftIndent` (pixels), and the exporter writes `leftIndent` plus 24 px per `indent` level.
- **Spacing of 0 survives editing.** An explicit zero space before/after a paragraph was dropped when the editor wrote
  its document back (and in collaboration), so it came back as the default.
- **DOCX loss report: `docx.paragraphFormatting` is gone and `docx.listFormatting` reports number formats only** (line
  spacing, indents, keep-with-next, page-break-before, hidden text, custom list starts and restarts are imported).
- **DOCX export keeps list item spacing and indents** (they were written for paragraphs and headings only).
- **DOCX import: a paragraph with `numId` 0 is not a list item** (it used to be treated as a numbered one).
- **DOCX loss report: `docx.namedStyles` effect narrowed.** The style's formatting is now applied to the text; only the
  style name is not kept.
- **DOCX loss report: `docx.comments` is reported only for comments that could not be read** (a damaged comments
  part), since readable comments are now imported.
- **DOCX import no longer drops text.** Tabs (`w:tab`, `w:ptab`) are imported as `\t`, line breaks (`w:br`
  other than page breaks, `w:cr`) as `\n`, soft and non-breaking hyphens as U+00AD and U+2011, and the text inside
  `w:smartTag`, `w:customXml`, content controls (`w:sdt`: inline, block, in table rows, cells and cell content, and
  in headers and footers), tracked insertions (`w:ins`) and `w:ruby` base text is kept. The wrappers themselves
  are still lost: tracked insertions show as ordinary text (as if accepted) while tracked deletions are still
  applied, content controls lose their type, tag, placeholder and lock, and ruby reading marks are dropped.
  `exportDocx` writes `\t` as `w:tab` and `\n` as `w:br`. Placeholder text of content controls is now imported
  as ordinary text. The editor measures text with `white-space: pre-wrap`, like the pages it paints, so tabs get
  a visible width (default tab stops; custom tab stops are still not imported), and header/footer text keeps tabs
  and line breaks.
- **DOCX loss report: effects narrowed.** `docx.trackedInsertions` and `docx.contentControls` keep their counts but
  now say the text is kept and what is lost, and `docx.ruby` says only the reading marks are lost. Tabs, line
  breaks and smart tags are no longer reported because nothing of them is dropped.
- **DOCX loss report: `docx.charts` and `docx.floatingImages` count what their labels say.** `docx.charts`
  counted XML parts under `word/charts` and `word/diagrams` (one SmartArt is several parts) and now counts chart
  and SmartArt objects referenced from the document; `docx.floatingImages` counted every `wp:anchor`, including
  anchored shapes and text boxes, and now counts anchored pictures only.
- **DOCX loss report: `docx.fields` counts a complex field once** even when its instruction is split across
  several runs (for example `IF 1 > 0 ...`, which was counted twice).
- **DOCX: landscape sections no longer lose their orientation on a round trip.** `exportDocx` handed the `docx`
  library an already-swapped page size and the library swapped it again, so a landscape section was written
  as 8.5x11in with `w:orient="landscape"` and re-imported as portrait-shaped. The writer now emits the on-page
  size (width larger than height for landscape), and the reader also accepts landscape written with
  `w:w` < `w:h` or with `w:w` > `w:h` and no `w:orient`. `PageConfig.size` keeps its natural-size convention.
- **DOCX import: accurate warnings for drawings that are not pictures.** `importDocx` reported every `w:drawing`
  without a picture as "An embedded image could not be resolved". Drawings are now classified by their
  `graphicData` URI: charts, SmartArt diagrams, shapes/text boxes/groups and embedded-object previews each get
  their own warning, linked (`r:link`) pictures get "A linked picture was skipped" and are never fetched, and the
  "could not be resolved" warning is kept only for real pictures. Warning text is observable by hosts; no content
  or loss-report change. Anchored non-picture drawings now carry the kind-specific warning instead of the
  floating-image one.

### Sheets and slides fidelity

- **XLSX fidelity corpus, first part (`fixtures/xlsx/`).** 26 committed files: eight hand-authored construct
  workbooks (cell types, 1900 and 1904 dates, basic styles, layout, sheet properties, names, conditional
  formatting, data validation), the 159-formula, 33-code number-format and 23-form formula-syntax probe
  workbooks, producer-variety files from openpyxl, xlsxwriter and LibreOffice (including three valid openpyxl
  files that make the importer throw), ten edge files (truncated zip, no workbook, broken relationships, XML
  entities, malformed XML, encrypted and legacy `.xls` OLE stubs, duplicate sheet names, indexes out of range,
  empty workbook) and three never-committed run-time generators (zip bomb, huge dimension, 100,000 x 10 rows).
  A manifest records the producer, licence and claimed constructs of each file; `scripts/xlsx-fidelity/`
  holds the construct probes and manifest checks; the Jest suite in `packages/sheets-core/src/fidelity/`
  verifies every claim against the package XML. No importer or runtime behaviour changes; goldens come with
  the XLSX runner.
- **XLSX fidelity corpus, second part.** 17 more files complete the planned construct list: theme colours, named
  styles, gradient and pattern fills and all 13 border styles; 60 number formats; modern formulas (`_xlfn`,
  LET and LAMBDA, dynamic and legacy arrays, shared formulas); references (whole row and column, 3D, circular,
  external link); comments (notes and threads); links (including a `javascript:` target); images (PNG, JPEG, GIF,
  SVG with fallback, EMF, three anchor kinds, crop); charts (bar, line, pie, scatter, combo); tables; filters and
  sort; a pivot table with a slicer; sparklines; x14 conditional formatting; sheet and workbook protection with
  password hashes; print setup; rich-text cells; and an `.xlsm` with an inert dummy macro part. The construct
  registry grows from 85 to 128 probes.
- **Generic fidelity core and the XLSX runner.** `scripts/fidelity-core/` holds the format-neutral modules (CLI
  arguments and runner, snapshot compare with orphan detection, structured diff with a cap, report model with
  markdown and JSON and a launch-bar scoreboard, manifest schema, claims, coverage matrix); the DOCX harness is
  untouched. `npm run sheets:xlsx-fidelity[:update]` and `npm run fidelity` (DOCX, then XLSX) run the 43 committed
  files against the built sheets-core and record five goldens each in `fixtures/xlsx/goldens/`: `model`, `loss`,
  `roundtrip`, `ledger` (every package part as mapped, opaque, dropped or unaccounted) and `values` (engine
  recalculation against the file's cached values). The goldens record today's behaviour, including the ten files
  whose import throws. CI gains an "XLSX fidelity corpus" step in the `packages` job (job summary plus a
  `xlsx-fidelity-report-node<version>` artifact). The cached values of `formulas-core.xlsx` are now the
  LibreOffice-computed results instead of zeros. No importer, exporter or runtime behaviour changes.
- **XLSX loss detector, round 2 (XML-based, `full` detail).** `detectXlsxLoss(bytes, options)`
  (`@b32nio/spindle-io/xlsx`) and `inspectXlsx(bytes, options)` (`@b32nio/spindle-sheets-core`) now parse the
  package XML instead of matching text and file names. `options.detail` is `'summary'` (default, the import
  dialog) or `'full'` (the fidelity harness); `options.knownFunctions` and `options.isNumberFormatExact` are
  optional (additive `DetectXlsxLossOptions` in `@b32nio/spindle-shared/export/xlsx`), and `inspectXlsx` passes the
  formula catalog and the number-format mapping. The summary report gains three ids: `xlsx.formulas.unsupported`
  (functions the editor does not have, `_xlfn.`/`_xlws.` removed before matching), `xlsx.formulas.structuredRefs`
  and `xlsx.numberFormats.approximated`, so the confirmation dialog now also appears for files with those. The
  full report adds 46 ids (formulas, styles, sheet, print, conditional formatting, validation, comments, links,
  objects, workbook and package: see `documentation/import-loss.md`), including `xlsx.package.encrypted` and
  `xlsx.package.legacyBinary` for OLE containers. Counts now count objects, not parts (a chart is one chart; a
  picture is one picture; a sheet with notes and threaded comments is one sheet), so some existing counts change
  (`xlsx.images`, `xlsx.outline`, `xlsx.comments`) and `xlsx.richText` finds rich strings whose first run is plain.
  Effect texts start with `lost:`, `approximated:` or `kept:`. The detector never expands XML entities, skips parts
  over a size limit (reported as `xlsx.package.sizeLimit`) and never throws. The XLSX fidelity loss goldens and the
  package-ledger rules are regenerated.
- **PPTX fidelity corpus, first part (`fixtures/pptx/`).** 22 committed files made by a stdlib `zipfile`
  PresentationML builder (fixed timestamps, normalised core properties, no third-party content): basic and rich text
  (levels, bullets, numbering, links, spacing, tab stops, columns, vertical text, `normAutofit` with `fontScale`),
  two masters with every placeholder type, themes with colour maps, font schemes and effect styles, all 187
  `ST_ShapeType` presets (generated from the list), custom geometry, fills (gradient, pattern, picture, alpha),
  lines (dashes, arrowheads, caps, joins, compound), effects (shadow, glow, reflection, soft edge, blur, 3-D), five
  slide sizes, and nine edge files (truncated zip, no `presentation.xml`, missing slide part, relationships to missing
  parts, malformed slide XML, 60 nested groups, relationship cycles, encrypted and legacy `.ppt` OLE stubs); a
  1000-slide deck is generated at run time and never committed. A manifest records producer, licence and claimed
  constructs; `scripts/pptx-fidelity/` holds the 83-construct registry and its probes on the generic
  `scripts/fidelity-core` modules; the Jest suite in `packages/slides-core/src/fidelity/` verifies every claim
  against the package XML, the structure of each deck and the byte-for-byte rebuild of the generated files
  (`npm run test:fidelity:coverage -w @b32nio/spindle-slides-core` reports coverage of the scripts). No importer or
  runtime behaviour changes; there is no PPTX importer yet, so goldens come with the PPTX loss detector and runner.
- **PPTX fidelity corpus, second part.** 17 more files complete the planned construct list: images (PNG, JPEG,
  GIF, SVG with a PNG fallback, EMF, WMF, crops, flips, recolouring, a linked picture), audio and video stubs, tables
  (merges, banding, borders, a custom table style), groups and connectors (nested, scaled, attached through
  connection sites), charts with embedded workbooks, SmartArt, transitions and animations (including the PowerPoint
  2010 forms and builds), notes, comments (legacy and modern with a reply), hidden slides, sections and custom shows,
  right-to-left and CJK text, WordArt, ink, an OLE object, an equation and action hyperlinks, a python-pptx deck and
  its LibreOffice conversion (tool and version recorded, normalised so that they rebuild byte for byte), a
  kitchen-sink deck and an `.pptm` with an inert placeholder macro part. The construct registry grows from 83 to
  148 probes.
- **PPTX loss detector and fidelity runner (`@b32nio/spindle-io/pptx`, `scripts/pptx-fidelity`).** New subpath
  `@b32nio/spindle-io/pptx` (also re-exported from the package root) with `detectPptxLoss(bytes, { detail })`: an
  XML-based, never-throwing scan of a `.pptx` with the same safety measures as the XLSX detector (no entity
  expansion, DOCTYPE cut out, nesting capped, per-part and total limits counted on the bytes that actually inflate,
  a limit on package entries) that returns 16 summary ids (`pptx.charts`, `smartArt`, `media`, `embeddedObjects`,
  `transitions`, `animations`, `comments`, `macros`, `hiddenSlides`, `notes`, `fonts.unavailable`,
  `shapes.unsupportedGeometry`, `images.unsupportedFormat`, `encrypted`, `legacyBinary`, `sizeLimit`) and 43 finer
  ones at `detail: 'full'` (masters, theme, text, shapes, images, tables, sections, zoom, ink, equations,
  signatures, links, actions, unaccounted parts; see `documentation/import-loss.md`). Counts are of objects, not
  parts. **There is no PPTX importer yet**: every effect says `lost: not imported` and that no `.pptx` can be opened
  today. `@b32nio/spindle-shared` gains the `'pptx'` member of `ImportLossFormat` (additive at runtime; a host with an
  exhaustive `switch` over it needs a new case) and a `./export/pptx` subpath with the `DetectPptxLossOptions`
  type. `npm run slides:pptx-fidelity[:update]` runs the 39 committed decks against the built package and records
  four goldens each in `fixtures/pptx/goldens/` (`model` and `roundtrip` record `importError: "no importer"` and a
  skipped round trip until the importer lands; `loss` and `ledger` are real, today 1,113 of 1,156 parts
  unaccounted because nothing is mapped), `npm run fidelity` now runs DOCX, XLSX and PPTX, and CI gains a "PPTX
  fidelity corpus" step in the `packages` job (job summary plus a `pptx-fidelity-report-node<version>` artifact).
  The Jest suite in `packages/slides-core/src/fidelity/` also checks the detector against the manifest claims. No
  runtime behaviour of any editor changes.
- **Formula engine: whole-column and whole-row references, `%`, array constants, 3D references, range operators
  over computed references, union, intersection, `@`, `#` and `#NULL!` (`@b32nio/spindle-sheets-core`).** The
  formula parser is now a tokenizer plus a precedence-climbing parser (`formula-parser/tokenizer.ts`,
  `syntax.ts`) with Excel's operator order, replacing the text-splitting one. New forms: `A:A`, `A:C`, `1:1`, `1:3`
  (stored unbounded in the AST, bounded to the sheet's used range when evaluated so that `A:A*B:B` lines up and a
  span of over 1,024 columns or rows is trimmed too, one dependency rectangle instead of a
  cell per row, so `SUM(A:A)` over a million rows takes about two seconds and an edit recomputes only the formulas
  of that column), postfix `%` (`50%`, `A1+50%`, `10%^2`), array constants (`{1,2;3,4}`, `,` across and `;` down, a
  ragged array is a parse error), 3D references (`Sheet1:Sheet3!A1`, `!A1:B2`, `!A:A`; sheet order from the workbook;
  valid as a function argument), ranges ending in a computed reference (`A1:INDEX(A1:A5,3)`, `OFFSET`, `INDIRECT`,
  `INDEX(..):INDEX(..)`; `INDEX` is now a reference function and `ROW`, `COLUMN`, `ROWS`, `COLUMNS`, `ISREF`, `CELL`
  accept all of these), union `(A1:A2,C1)`, intersection (a space), `@` implicit intersection, the spill operator
  `A1#` and the `#NULL!` error literal (all error literals are now literals). The sheet-edit rules are documented in
  `documentation/sheets/core/formulas.md`: insert/delete rows never touch a whole-column reference, a whole-row
  reference follows its row and is `#REF!` once the row is deleted, a 3D reference is positional, follows renames
  and the sheet order, and shrinks inward when an end sheet is deleted. A formula that does not parse now shows
  `#ERROR! <what and where>` (and stops being the previous formula) instead of being read as text (a chain deeper than 500 levels is `#ERROR! Formula is too complex`); an imported
  formula the engine cannot parse keeps its saved value and text through recalculation and export never writes an
  error message as a result; unquoted sheet names of words or digits (`Q1 Data!A1`, `2024!A1`) are still read and
  printed quoted; array-constant errors are tagged (`{#N/A}` is not `{"#N/A"}`); a bare name is
  `#NAME?` but keeps its text; a deleted reference prints as `#REF!` and parses back. `MAX`, `MIN`, `SUBTOTAL(4|5)`
  and every range function no longer overflow the stack on a column of hundreds of thousands of values. Public
  types are widened, additively: `ParsedFormulaNode.type`/`StableFormulaNode.type` gain `error`, `array`,
  `percent`, `implicit`, `spillref`, `ref3d` (new `FormulaNodeType`), `RangeReference`/`StableRangeRef` gain
  `whole`, `RangeDependency` gains `axis`, `StableFormulaNode`/`ParsedFormulaNode` gain `sheetEnd`,
  `EvaluationContext` gains the optional `getSheetNames` and `getSpillRegion`, and `MAX_SHEET_ROWS`,
  `MAX_SHEET_COLS`, `WHOLE_AXIS_ID`, `WorkbookImpl.getSheetNames()` are exported. The XLSX loss id
  `xlsx.formulas.wholeColumn` is gone (the engine now calculates these), the 23-form formula-syntax probe goes from
  11 to 20 matching values, and the XLSX `values` goldens from 134 to 147 of 214 engine-evaluated values matching (the harness now counts
  the 11 formulas the engine cannot parse, which keep the file's value, separately as `passthrough`); the 4 left in that probe
  are the `_xlfn.` prefixed forms (`SINGLE`, `ANCHORARRAY`, `SEQUENCE`) that wait for the prefix mapping.

## [0.4.0] - 2026-10-26

### Added

- **CSV, XLSX and DOCX import/export.**
  - `@b32nio/spindle-shared` gains the subpaths `./export/csv`, `./export/xlsx`
    and `./export/docx`: format-level readers/writers (`exportCsv`/`importCsv`,
    `exportXlsx`/`importXlsx`, `exportDocx`/`importDocx`) independent of any
    editor model (#36, #37, #38, #39, #40).
  - `@b32nio/spindle-sheets-core`: `exportToXlsx` / `importFromXlsx`; CSV
    import/export now runs on the shared CSV module (#36, #37, #38).
  - `@b32nio/spindle-docs-core`: `exportToDocx` / `importFromDocx` (#39, #40).
- **Open / Download in the editor toolbars** (opt-in, files are read from a local
  file input only; no network):
  - `WorkbookCanvas` (`sheets-react`): `fileActions`, `exportFormats`,
    `fileName`, `onFileOpened`, `onFileError`. Open accepts `.xlsx` and `.csv`;
    Download offers XLSX and CSV. Also exported: `exportWorkbookFile`,
    `importWorkbookFile`, `SHEETS_EXPORT_FORMATS`, and the `Toolbar` props
    `onOpenFile` / `onDownload` / `downloadFormats`.
  - `DocumentEditor` (`docs-react`): the same props for `.docx`. Also exported:
    `exportDocumentFile`, `importDocumentFile`, `DOCS_EXPORT_FORMATS`.
  - `@b32nio/spindle-shared/react`: `downloadFile`, `fileExtension`,
    `stripExtension`, `sanitizeFileName`.
  - The sheets and docs demos in `examples/` enable the new buttons.
- Dependabot configuration (weekly, grouped updates) for npm and GitHub Actions.
- `.nvmrc` (Node 22).
- **Import loss report.** `ImportLossReport` (in `@b32nio/spindle-shared`) lists what an .xlsx / .docx / .csv import cannot
  preserve (conditional formats, validation, named ranges, charts, comments, tracked changes, footnotes, floating images…).
  `inspectWorkbookFile` / `inspectDocumentFile` (react packages) and `inspectXlsx` / `inspectCsv` / `inspectDocx` (core) check a file
  without importing it; `importWorkbookFile` / `importDocumentFile` / `importFromXlsx` / `importFromDocx` now also return `loss`, and
  `onFileOpened` receives it. See `documentation/import-loss.md`.
- **Import confirmation.** Opening a file that would lose something now shows an *Import anyway / Cancel* dialog
  (`ImportConfirmDialog`) in the sheets and docs editors; `importConfirmation` selects `'dialog'` (default), `'never'` or a host callback.

### Changed

- **Minimum Node.js is now 22** (`engines.node` is `>=22.0.0` in every package);
  Node 20 reached end of life. CI runs the package job on Node 22 and 24; the
  website, deploy and publish jobs run on Node 22.
- Internal dependency ranges bumped to `^0.4.0`.
- **Privacy: image hooks on the docs and slides editors.** The editors no longer fetch typed image URLs themselves.
  `resolveImage(url) => Promise<Blob | dataURL>` enables insert-by-URL (otherwise the By-URL UI is hidden and only
  local files can be inserted); `uploadImage(file) => Promise<url>` stores local files with the host;
  `maxImageBytes` bounds embedded images. See `documentation/images.md`.
- **Privacy: `fontSource` option on the sheets, docs and slides editors.** The hard-coded Google Fonts
  loader in the sheets toolbar is gone. Web fonts now come from the Spindle font mirror
  (`https://fonts.b32n.io/css2`) by default and fall back to system fonts if it is unreachable.
  `fontSource={{ kind: 'local' }}` makes zero font requests; `{ kind: 'css', url }` and
  `{ kind: 'custom', load }` let hosts bring their own source. See `documentation/fonts.md`.

### Security

- `prosemirror-view` updated to 1.42.x, fixing an XSS in paste handling
  (GHSA-c8x8-7fp4-3x9w).
- `brace-expansion` updated, fixing a recursion DoS (GHSA-qhr7-859c-m2p7,
  GHSA-6j4f-fj2g-mc7p).
- `uuid` forced to `^11.1.1` through npm `overrides` (used only by `exceljs`,
  which only calls `v4()`), resolving GHSA-w5hq-g745-h8pq.
- `npm audit --omit=dev` reports no vulnerabilities for the library packages.

### Documentation

- Removed stale claims that XLSX/DOCX import/export were missing from the
  READMEs and TODO lists; documented the new APIs and subpaths.

## [0.3.1] - 2026-09-17

### Added

- `readOnly` mode for the sheets and docs editors, matching the existing
  `SlidesEditor` prop (#31).
- Filmstrip mutations are gated behind `readOnly` in the slides editor (#34).

### Changed

- Slides title bar folded into the toolbar; sheets toolbar edge-bleed fixed;
  demos renamed to "Spindle" (#32).
- Publishing switched to `npm stage publish` with a manual 2FA approval step
  (#33).
- Fixed Rollup build warnings across all packages (#30); `@b32nio/spindle-shared`
  added to the docs-core / sheets-core Rollup externals (#35).
- Repository and npm scope moved to `b32n` / `@b32nio` (#29).

> Note: 0.3.1 was published to npm but has no `v0.3.1` git tag.

## [0.3.0] - 2026-07-14

### Added

- Responsive (overflowing) toolbars for docs, sheets and slides, via
  `ResponsiveToolbar` in `@b32nio/spindle-shared/react` (#11).
- Slides on the website; npm Trusted Publishing (OIDC) workflow (#14);
  Cloudflare Pages deploy for the website (#13).

### Fixed

- A series of review rounds (#18–#28) hardening formula-engine, collaboration,
  rendering, packaging and security issues (including ReDoS and data-loss
  fixes), plus docs-core and pagination test coverage.

## [0.2.0] - 2026-05-20

### Added

- Comments (threads, replies, resolve, @mentions) for docs and sheets, with the
  shared comment types in the `shared` package.

### Changed

- Manifests fixed for publishing: `shared` declared as a dependency of the
  cores and React layers, `publishConfig.access: public`, and internal ranges
  pinned.

## [0.1.0] - 2026-01-07

### Added

- Initial release: docs, sheets and slides editors (React layers over
  framework-free cores) with Yjs-based real-time collaboration primitives.

[Unreleased]: https://github.com/b32n/spindle/compare/v0.4.0...HEAD
[0.4.0]: https://github.com/b32n/spindle/compare/v0.3.1...v0.4.0
[0.3.1]: https://github.com/b32n/spindle/commit/d04923d
[0.3.0]: https://github.com/b32n/spindle/releases/tag/v0.3.0
[0.2.0]: https://github.com/b32n/spindle/commits/cfaefac
[0.1.0]: https://github.com/b32n/spindle/commit/012b8a9
