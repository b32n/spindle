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

- **Tracked changes (insertions and deletions) are kept, shown and written back.** `w:ins` and `w:del` around runs and links
  (also inside table cells, content controls, smart tags and hyperlinks) are imported as the same text carrying the change, not
  applied: inserted text stays in the text, deleted text (`w:delText`) stays too, struck through, until the change is accepted
  or rejected (the review panel and suggesting mode are the next steps of this milestone). A run can be both inserted by one
  author and deleted by another. The export writes `w:ins`, `w:del` and `w:delText` with the author, date and id (the runs of a
  change share one element; ids that are not numbers, or are taken, get new ones), the link of a changed link inside the
  hyperlink. Additive: `DocRevision` / `DocRevisionMeta` and `revision?` on the interchange runs (`DocRunBase`), `RunRevision` /
  `RevisionMeta` and `revision?` on `TextRun` and `InlineLink`, and `docsSchema` gains the `insertion` and `deletion` marks
  (attributes `revisionId`, `author`, `date`; not inclusive), plain ProseMirror marks that sync over Yjs, so the collaboration
  `schemaVersion` is now 6; `getCollabBlockers` is unchanged (runs with a change are ordinary text). The editor draws insertions
  underlined and deletions struck through in the author's colour (like Word's All Markup; deleted text takes part in the
  layout) and shows a legend of the authors above the page (`RevisionLegend`). New helpers in docs-core: `authorColor`,
  `assignAuthorColors` (a colour per author of the document, the same on every client, nothing stored), `collectRevisions`,
  `revisionAuthors`, `documentRevisionAuthors`. Marked text takes the author's colour over its own (a link, a content-control
  placeholder), as in All Markup. Editing around changes: copying tracked text and pasting it gives the copy change ids of its own
  (author and date stay), so accepting or rejecting one place never touches the other (`freshRevisionIds`); text typed in the
  middle of a deleted run is live text, not part of the deletion (also through the editor's own input path and IME composition, dead keys and other browser-reported input, which never take the deletion; text typed inside someone else's insertion in Editing mode becomes part of that insertion, as it grows the change; typing over a selection that lies wholly inside a deletion replaces the struck-through text, as any edit in Editing mode that is not recorded; pasting plain text into a deleted run is live text too, while pasted HTML keeps its own marks; cut and paste, and drag and drop, also count as a paste, so a moved tracked range gets a new id); the plain text copied to the clipboard, `getPlainText` and `getPmPlainText` leave deleted text out (the final text; pass `includeDeleted` to keep it); a line break typed in tracked
  text (Shift+Enter) is a newline in the text, so its change survives collaboration. The DOCX writer no longer uses a regular
  expression look-behind (older Safari cannot parse one), writes a `w:date` without seconds with `:00` added (the file format requires the seconds) and a deleted run counts as content
  when a page break follows it. Copied tracked text now keeps its changes when pasted in the browser (the clipboard HTML was parsed as plain text style before: the parse rules of the change marks have a priority of their own). Updating a table of contents lists the final text of a heading (words in a tracked
  deletion are left out). The authors of the changes are collected when typing pauses, not on every key. The paragraph break between the paragraphs of a note no longer becomes part of a tracked change at the end of the paragraph before it, and the same change id from two authors counts as two changes in the list of changes.

- **Tracked moves, formatting changes, paragraph-mark and table-row changes are kept, shown and written back.** The two ends
  of a move (`w:moveFrom` / `w:moveTo` with their range markers) are read as a deletion and an insertion that share the name of
  the move (`moveId`) and written back with the range markers; a change of a run's formatting (`w:rPrChange`) keeps what the
  formatting was before (`formatChange`: the effective formatting, in the same fields as a run) next to the formatting now;
  likewise `w:pPrChange` for a paragraph (with the heading level it had); an inserted or deleted paragraph mark
  (`w:pPr/w:rPr/w:ins|w:del|w:moveFrom|w:moveTo`; a move of a paragraph mark is written as an insertion or deletion, as the writer's
  library has no move there) and an inserted or deleted table row (`w:trPr`). Additive: `DocRevisionMeta.moveId`,
  `DocFormatChange`, `DocRunFormat`, `DocParagraphFormat`, `DocBlockRevision`, `DocRevision.formatChange`,
  `DocParagraphStyle.revision` and `DocTableRow.revision` in the interchange types; `RevisionMeta.moveId`, `FormatChange`,
  `RunFormat`, `PreviousParagraphFormat`, `BlockRevision`, `RunRevision.formatChange`, `ParagraphFormat.revision` and
  `TableRow.revision` in docs-core. `docsSchema` gains the `formatChange` mark (attributes `revisionId`, `author`, `date`,
  `previous`; the formatting read back from pasted HTML is limited to known fields), a `moveId` attribute on the `insertion` and
  `deletion` marks, and the attributes `revisionInsertion`, `revisionDeletion` and `revisionFormat` on paragraphs, headings and table rows (one per kind of change, so two people suggesting different things on the same paragraph mark or row at once both keep theirs; the same kind from two people is last-writer-wins; after a concurrent split a record made on the whole paragraph stays with the first half, next to the insertion the split recorded there, not on the half that holds the original mark; a record read from another client is rebuilt from the fields it should have, and one without a usable id is ignored); all are plain ProseMirror content that
  syncs over Yjs; the collaboration `schemaVersion` is now 7 (a client at 6 would drop the `formatChange` mark and the new attributes, so it refuses to join such a room). A move that spans several paragraphs (what Word writes for a moved block) loses its shared name on import: its two ends are read as an ordinary deletion and insertion, with a warning, and are no longer paired in the editor. Pressing Shift+Enter in text that only has a formatting change inserts a newline character in the text (a `hard_break` node would not sync its change). The font names kept as the formatting before a change may contain any letter, digit or space (CJK names included). The editor draws the two ends of a move with a
  double line, a change of formatting alone with a dotted underline (the text keeps its colour), a bar in the margin beside a
  paragraph with a change, and a tint on the cells of an inserted or deleted row. `documentRevisionAuthors` also lists the authors of changes recorded on paragraphs and rows and of formatting changes, so they share out the palette like the others.

- **Review panel: accept and reject tracked changes.** `DocumentEditor` has a Review changes button in the toolbar that opens a
  panel listing the changes (author, date, what it did, the text) with Accept and Reject for each, for all of them and for all of
  one author's (a filter), a next and a previous change (the caret goes there and the page scrolls to it), and a choice of how
  changes show: **All markup**, **Final** (as if every change were accepted: deleted text left out, no markup) and **Original**
  (as if every change were rejected: inserted text left out, the formatting before a change back). The panel has no Accept or
  Reject in a read-only editor, and the Review button is on the toolbar, which a read-only editor does not show (open the panel with the `ReviewPanel` component there). **The Final and Original views are read-only**: a banner ("Showing final text: switch to All markup to edit" / "Showing original text: ...") offers a one-click switch back to All markup (`RevisionViewBanner`), and the toolbar and rulers are hidden while a view is showing, as in a read-only editor; Accept and Reject in the panel still work there (they are explicit commands). The colours in the panel and the legend are shared out among every author of the document, as on the page (also for the author of the other end of a move). New prop `DocumentEditor.revisionView` (the view to begin with, of the type `RevisionView`, exported) and `TrueLayoutEditor.revisionView`;
  `proseMirrorToFlowBlocks` takes `revisionView` in its options (text a view leaves out keeps its length as zero-width spaces,
  so positions stay in step; in the Final and Original views rows and paragraph formatting are not shown as they were, only text and
  character formatting, a paragraph whose text a view hides stays as an empty line, and paragraph breaks are not hidden: a deleted mark still splits paragraphs in Final, an inserted one in Original). New optional `TrueLayoutEditorHandle.goTo(pos)` (optional, so code that implements or mocks the handle keeps compiling). The Next and Previous buttons of the panel step through the changes in document order, also through changes that start at one position or a move whose old place is above its new one; a move with several ends covers all of them; a settle that changes nothing no longer puts an empty step on the undo history. Accepting the deletion of the mark of an empty paragraph keeps the first paragraph's type.
  `ReviewPanel` and `RevisionViewBanner` (with `ReviewPanelProps`, `RevisionViewBannerProps`). In docs-core: `listChanges` (the changes of a document, runs of one change joined, the two
  ends of a move as one), `changeUnits`, and the commands `resolveChange`, `resolveAll` (all, or one author's) and
  `resolveSelection`, plus `settle`. Accepting keeps an inserted run, drops a deleted one, keeps new formatting, keeps an inserted
  paragraph mark or row and joins away a deleted one; rejecting does the reverse (the old formatting of a run, and the old
  formatting and heading level of a paragraph, come back; in a list item a former heading comes back as a paragraph). Joining two
  paragraphs (a paragraph mark that goes) keeps the first one's formatting and the second one's mark with its changes (a deletion of somebody else on the first mark moves to the joined mark, so rejecting a concurrent split keeps the other person's suggestion; the same happens when somebody deletes a mark that somebody else inserted earlier and the insertion is rejected: the deletion stays on the joined mark). A move
  is settled whole, also for one author (it belongs to the author of the end it went to). Everything one action does is one
  transaction, so one undo takes it back. Known limits: the Original view restores character formatting only (not links or the
  character style from before a change); a change mark on an inline object (an image, a field, a note reference) is still drawn in the Final and Original views, while the placeholder of a hidden link is plain text; rejecting all with an inserted paragraph mark before a table whose only row is also rejected joins the paragraph with the one after the table, while rejecting that mark alone leaves them apart; deleting text with a comment on it removes the comment's anchor and leaves the thread
  in the document; the legend counts changes the way the panel lists them (a move once, the runs of one change as one), not runs.
  The list of changes is kept only while the panel is open. Switching between editable and read-only (these views, a host's `readOnly`) no longer rebuilds the editor from the document it was mounted with: the editor reads `editable` live, so edits and accepted or rejected changes made before the switch stay.

- **Suggesting mode: typing creates tracked changes.** `DocumentEditor` has an Editing / Suggesting / Viewing mode switch
  above the page (props `mode`, `onModeChange`, `showModeSwitch`; uncontrolled by default; Viewing is read-only; never shown with
  `readOnly`). In Suggesting mode the author is the current user (`currentUser`). The new `suggestingPlugin` of docs-core (also
  `suggestChanges`, `suggestingKey`) puts what the user does into the state a suggestion would have left it in, in a transaction
  that joins the user's in the undo history: typed or pasted text gets an insertion mark (consecutive typing is one change), deleted
  text is put back with a deletion mark (text the same author inserted themselves is just removed; backspaces and deletes move
  the caret over the kept text), Enter marks the paragraph break inserted, joining paragraphs keeps the break marked deleted
  (a pasted paragraph does not bring the formatting-change record of its source; Enter then Backspace leaves no trace, also after typing in the new paragraph, and the records another person had on the paragraph mark are kept; the first Backspace
  at the start of a paragraph moves the caret before the deleted break, so the second one deletes the character before),
  text typed in the middle of another author's deletion is an insertion of its own, never part of the deletion, a change of character formatting records the formatting before, on text only (`formatChange`; an image, field, bookmark or other object in the range keeps the formatting without a record and the notice says so; none
  when it is put back, or on the author's own insertion), a change of a paragraph's formatting or type records the old formatting.
  What cannot be recorded (a list, table, image or other structure inserted or deleted, wrapping blocks in a list, an inline image, field, bookmark, note reference or content control, which would lose a mark in collaboration because marks sync on text only; a paragraph pasted with such an object is inserted directly) is left as
  the user made it and `TrueLayoutEditor.onSuggestionNotice` / a notice above the page says so. (A transaction that a host plugin appends is recorded like the user's own edit.) Changes from collaborators, undo
  and redo, and accepting or rejecting in the review panel are never suggested. New `TrueLayoutEditor` props `suggesting`,
  `suggestAuthor` and `onSuggestionNotice`; new `ModeSwitch` component and `EditorMode` type. Formatting several marks in one
  transaction (clear formatting) records the formatting before the first of them, the selection stays a range after a format,
  a line break typed or deleted is kept as text of the tracked run (so it syncs with its change), and notices go out after the
  transaction was applied. Limits: only the body text is suggested (an edit in a header or footer is made directly, and a notice
  says so once per mode; a note edited in place is recorded as a replacement of its whole text: the old text deleted, the new
  text inserted; a change to the properties of a list, table, image, form control or other structure, and Increase / Decrease indent, are made directly, with a notice; Enter in a list item is recorded as an inserted paragraph mark, lifting an item out of its list with Backspace is made directly; a link or
  a comment added inside someone's change is not recorded; text composed with an IME is recorded like typing but has not been tried with a real input method), and a change is continued by the same author name, with no time limit, so two people with one display name
  share a change. `RevisionLegend` keeps its own band unless drawn with the new `bare` prop (which `DocumentEditor` uses). Formatting text that carries another person's formatting change is applied directly, with a notice (rejecting their change brings back the formatting from before it). The mode switch group is labelled "Editor mode", and the notice is announced through a live region that is always present.

- **Tracked changes between collaborators.** In a shared document the author of a suggestion is the name the user joined under
  (`CollabIdentity.displayName`, before the provider's `currentUser`), and people who are in the room are drawn in the colour of
  their cursor: new `collabAuthorColors(handle)` (docs-core, also from `@b32nio/spindle-docs-core/collab`) reads it from the
  presence (awareness), `assignAuthorColors`, `revisionAuthors` take a `preferred` map of such colours (the others share out the
  palette, never a taken colour), `ConversionOptions.authorColors`, `TrueLayoutEditor.authorColors` and `ReviewPanel.authorColors`
  carry it, and the legend, the panel and the markup agree. The collaboration `schemaVersion` is not raised again (7, for the
  tracked-change marks and attributes): a client from before them refuses to join such a room. Suggestions, and accepting and
  rejecting them, are ordinary edits of marks and attributes, so they merge, converge and survive being made offline like other
  edits (tests with three and four clients, concurrent accepts and rejects of one change, and a person who was never part of
  the suggestions settling them all; two people splitting and deleting the mark of one paragraph at once keep both suggestions
  and settle them one by one, also in the browser, with the two offline). Limits: two people suggesting the *same* kind of change on one
  paragraph mark or row at once leave one of the two (the same everywhere), and when one splits a paragraph while another marks
  its mark, the other's record lands on the first half; text a person adds to a change while someone else settles it is settled
  with it; two people with one display name are one author, in the smallest of their cursor colours (and a blank joined name falls
  back to the current user). Comments still carry the provider's `currentUser`, not the name the user joined under.

### Changed

- **DOCX loss report: `docx.trackedInsertions`, `docx.trackedDeletions` and `docx.trackedMoves` are gone, and `docx.formattingRevisions`
  counts only what is still lost**: insertions, deletions and moves of text, of paragraph marks (also in a table cell of one
  paragraph, which keeps its block when its paragraph has a change, like the export does) and of table rows are kept. What
  is left: the formatting change of a paragraph mark's own formatting, a paragraph change of numbering, tab stops, borders or
  shading (the old value is not kept, so rejecting cannot restore it) and changes of tables, table rows, cells and sections (`w:tblPrChange`,
  `w:trPrChange`, `w:tcPrChange`, `w:sectPrChange`, and the cell changes `w:cellIns`, `w:cellDel`, `w:cellMerge`, whose cell is imported as an ordinary cell), and `docx.trackedBlockChanges` / `docx.trackedHeaderFooterChanges`. A change inside a field result is not kept, and neither is a formatting change on a run with no text of its own (a field result, an image, a note reference); each adds a warning. A paragraph whose only content is an image keeps its paragraph (and the change of its mark); a tracked change of the mark of a page-break or rule paragraph is reported as lost (`docx.trackedStructureChanges`, plus a warning). A formatting change (`w:pPrChange`) in a header or footer is counted with the other changes there. Text moved away (`w:moveFrom`) is skipped in a comment's text and counts as content before a page break. The `w:id` of a change is no longer renumbered when a move's range markers are written, and pasting renews a numeric record id too.
- **Tracked insertions and deletions are no longer applied on import.** Before, `w:ins` text came in as ordinary text and
  `w:del` text was dropped. Now both are in the document, as marked text (see above), so the text of a document with
  deletions is longer than before. A change around a field is not kept as a change (an inserted field stays as it is, a deleted
  one is dropped, with a warning), as are a deleted image and a deleted note reference (dropped, with a warning); an inserted
  image or note reference stays as ordinary content, with a warning. Headers and footers keep no tracked changes: deleted text is
  dropped, inserted text stays as ordinary text (warnings, and `docx.trackedHeaderFooterChanges` in the loss report). Changes
  around whole paragraphs or tables are counted as `docx.trackedBlockChanges` (their content is not imported). A change without
  an id gets a number that is unique in the import. A line break inside changed text keeps the change (it stays in the text node
  in the editor, because the collaboration binding keeps marks on text only), and the export drops characters XML cannot carry
  from an author and a date that is not a date (a date out of range is dropped too, and two authors' changes with the same id get two `w:id`s). The still-lost changes around a field, an
  image or a note reference, and deleted text in a field result, are counted as `docx.trackedChangesNotKept` in the loss report (a `w:moveFrom` or `w:moveTo` around them too, and a change inside another one counts once); a deleted image in `mc:AlternateContent` is dropped like a plain one.
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

- **Editing no longer lays the document out with the settings the editor first rendered with.** After a keystroke the editor used
  the layout function from its first render, so later changes of `footnotesAtPageBottom`, `revisionView` or `authorColors` were
  undone as soon as the user typed (the Final view showed its markup again). It uses the current one now.
- **Switching an editor to read-only and back no longer resets the document.** `TrueLayoutEditor` rebuilt its ProseMirror view (and
  seeded it again from the initial content) whenever `editable` changed, so the edits made before the switch were lost on screen.
  It now reads `editable` live. (Needed by the Viewing mode.)
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
  are still lost (tracked changes are kept as such since the tracked-changes entry above), content controls lose their type, tag, placeholder and lock, and ruby reading marks are dropped.
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

- **The nine remaining values-golden mismatches are gone (SA-gap).** (1) A circular reference met while a workbook is calculated on load (`recalculateOnLoad`, the
  values harness) is settled like one met by a full recalculation: the cells on the cycle take the circular value, 0 by default, instead of `#VALUE!`.
  (2) A formula an .xlsx file saved without the array mark is intersected, as Excel does, where it uses a multi-cell range as an operand of its own operators
  (`=A1:A3*2` in row 152 is `#VALUE!`, in row 2 it is `A2*2`); it is shown with the `@` Excel 365 shows and saved back as the bare range the file had (`A1:A3*2`, not `_xlfn.SINGLE(A1:A3)*2`, which Excel 2016 reads as `#NAME?`); an `@` before a call, a name or a spill reference, and in a dynamic array, still writes `_xlfn.SINGLE`. A range inside a call's arguments is still not intersected (`LEN(A1:A3)`, `IF(A1:A3>1,...)`, `SUM(A1:A3*2)` can return a value where Excel gives `#VALUE!`; documented and pinned as known divergences). A formula whose result is an error now keeps its reference text in step with row and column insert and delete (it stayed stale, and was exported so). (3) The
  `formulas-core` and `xlsxwriter-rich` fixtures were wrong, not the engine: the dynamic-array probes sat one under the other, so every spill was blocked
  by the next probe (Excel recalculating the file gives `#SPILL!` too), and a SORT spill carried xlsxwriter's placeholder 0. The generators now give each spill room,
  its true range and its results. Values golden: 195 of 204 matched before, 204 of 204 now (cached-only 21, unchanged).
- **References inside cached-only formulas follow inserted and deleted rows and columns (S4b).** A formula the engine cannot evaluate
  (`cachedOnly`) keeps no parsed form, so inserting or deleting rows or columns used to leave the cell references in its text where they were.
  They are now rewritten in the text, using the formula tokenizer, with the outcome of a normal formula's: `A1`, `$A$1`, `A1:B2`, `A:A`, `1:1`,
  `Sheet2!A1` and `'My Sheet'!A1` move with their cells, a reference into deleted rows or columns becomes
  `#REF!`, a range keeps the rows or columns of it that survive (and is `#REF!` only when all of them are deleted), a reference to another sheet only changes when that sheet is edited, and everything else in the text (whitespace, case, string
  literals, function names, structured references, fragments the grammar rejects) is kept byte for byte. 3D references, references to sheets
  that do not exist and external workbook references are left as they are, as for a normal formula. Undo restores the text, and with
  collaboration the shifted text is the cell's new value. The saved value stays, and the note's "does not recalculate when the cells it uses
  change" caveat stays. The wording everywhere now says: the saved value and the formula are kept; the formula's cell references follow row and column insert and delete (they do not follow anything else, and the value does not update when the cells it refers to change); the S4 phrases "written back unchanged" and "saved unchanged" for these cells are gone from the note, the loss effect, the docs and the goldens. New exports of `@b32nio/spindle-sheets-core`: `shiftFormulaRefs`, `shiftIndex`, `shiftSpan`, `shiftPair` and `StructuralEdit`. The
  `xlsx.formulas.cachedOnly` effect text and `documentation/import-loss.md` changed accordingly.
- **A range keeps its surviving rows and columns when part of it is deleted (fix, found in the S4b review).** Deleting the first or last row or
  column of a range, or a column or row that is a boundary of it, made the whole range `#REF!` (`=SUM(A2:A5)` without row 5), although deleting
  one in the middle shrank it: the engine disagreed with itself and with Excel. A range now keeps the rows and columns of it that survive
  (`A2:A5` without row 5 or without row 2 is `A2:A4`, `2:4` without row 2 is `2:3`, `B:D` without `B` is `B:C`) and is `#REF!` only when all of
  them are deleted; a single cell into deleted rows or columns is still `#REF!`. Inserting before the first row shifts a range, after its last
  row leaves it, inside it extends it. It applies to normal formulas (stable ids: the range boundary moves onto the nearest surviving row or
  column before the delete; undo restores it) and to the text of cached-only formulas alike. Sheet names in references are case-insensitive
  (`sheet2!A1` points at `Sheet2` and follows its rename and edits). A workbook saved with such a `#REF!` keeps it; only new deletes change.
- **Slides: deck <-> PPTX (K15a).** New additive exports from `@b32nio/spindle-slides-core`: `importFromPptx(bytes, deck, options?)`
  (replaces the deck with the file: slide size, theme, layouts with their text levels, shapes, text with marks and paragraph
  attributes, pictures, tables, connectors glued to their shapes, groups, notes, comments, and charts, diagrams, objects and media
  kept byte for byte), `exportToPptx(deck, options?)` and `inspectPptx(bytes, { detail? })` (the loss report, with the "kept, not
  editable" group). `@b32nio/spindle-io/pptx` loads on first use, like the sheets and docs cores, and `@b32nio/spindle-io` becomes a
  dependency of `slides-core` (a workspace package, as in sheets and docs). `@b32nio/spindle-shared/export/io-loader` gains
  `loadIoPptx`. Type sizes and spacing in points become px at 4/3 (18 pt is 24 px) and back. Slide-level comments are kept as threads
  with no element (listed as "no longer attached", written back as slide comments). No model fields were added, so the slides
  collaboration schema version stays at 10. The Open and Download buttons follow in K15b.
- **Slides: Open and Download .pptx in the toolbar (K15b).** `SlidesEditor` gets the same host contract as the docs and sheets editors: `fileActions` (opt-in, default off, so the default toolbar is unchanged) shows Open (a `.pptx` file picker) and Download (`Microsoft PowerPoint (.pptx)`) in the toolbar, with `importConfirmation` (`'dialog'` default, `'never'` or a host function), `fileName`, `onFileOpened({ fileName, warnings, loss })`, `onFileError(error, 'open' | 'download')` and `onExportWarning(warnings)`. A file whose import loses or only keeps something is confirmed first by the import dialog, which now has its "Kept, not editable (saved back unchanged)" group filled by charts, diagrams, objects, media, transitions and animations. New exports of `@b32nio/spindle-slides-react`: `exportDeckFile`, `importDeckFile`, `inspectDeckFile`, `openDeckFile`, `downloadDeckFile`, `SLIDES_EXPORT_FORMATS`, `SLIDES_IMPORT_ACCEPT`, `FileActions` and the types `SlidesExportFormat`, `ImportConfirmation`. The website demo has the buttons. The slides demo shows them with `?fileActions=1`. E2E `SL-FIO-01` to `SL-FIO-06` replace the `SL-GAP-01` gap row.
- **Sheets collaboration gate.** Sheets rooms now record `meta.schemaVersion` (missing key = 1; no migration), and
  `attachCollab` refuses a room written by a newer client with a `CollabSchemaVersionError` instead of editing it.
  New additive exports from `@b32nio/spindle-sheets-core`: `getCollabBlockers(workbook)` (empty for now) and
  `CollabBlockedError`; the `/collab` subpath also exports `COLLAB_SCHEMA_VERSION`, `CollabSchemaVersionError`,
  `readSchemaVersion` and `assertSchemaVersion`.
- **Opaque part store and loss state (X5).** New additive subpath `@b32nio/spindle-shared/opaque`: a
  content-addressed, size-capped store for package parts kept byte for byte (`createOpaqueStore`, `hashBytes`
  (SHA-256 via SubtleCrypto), `OpaquePart`, `OpaqueRef`, defaults 25 MiB per part and 200 MiB per store), a
  Yjs-free helper for a per-document `Y.Map<hash, Uint8Array>` (`writeOpaqueParts`, `readOpaqueParts`, which
  re-hashes every entry from a peer) and loss-item builders (`keptLossItem`, `droppedLossItem`). A part over a cap
  is refused and reported as a `lost: too large` item, never dropped silently. `ImportLossItem` gains an optional
  `state: 'lost' | 'approximated' | 'kept'` (absent = lost; new `ImportLossState` type and `groupImportLoss`),
  and `ImportConfirmDialog` lists "Kept, not editable" apart from "Will be lost" when any item has a state. A
  report without states renders exactly as before. No importer emits `state` yet.
- **Own XLSX package reader, foundation (S2a).** New internal reader layer in `packages/io/src/xlsx/read/` on `jszip` +
  `fast-xml-parser` through the safe-import modules: package open, `[Content_Types]`, relationships (absolute and
  relative targets, backslash and percent-encoded forms, case-insensitive types), `workbook.xml` (sheets with state,
  `definedNames`, `calcPr`, `date1904`), shared strings (rich runs kept), styles (fonts, pattern and gradient fills,
  borders with colours, `numFmts`, `cellXfs`, `cellStyleXfs`, `dxfs`) with theme, indexed and tint colours resolved, and
  comment parts located. Additive types in `@b32nio/spindle-shared/export/xlsx` (`XlsxWorkbookInfo`, `XlsxStyles`, ...) and
  optional `workbook`, `styles`, `sharedStrings` and `theme` fields on `SpreadsheetImportResult`; nothing populates them yet. Not wired into
  `importXlsx` (the `exceljs` path is unchanged and import output is identical); the sheet reader and the switch are S2b.
- **XLSX import uses the own reader (S2b).** `importXlsx` now reads a workbook with the own package reader (`jszip` and its own
  forward-only XML scanner, `sax.ts`, for the workbook-level parts, `sheetData` and the small sheet elements; `fast-xml-parser` is left to
  the loss detector; no new dependency) instead of `exceljs`; export is unchanged (still `exceljs`). New sheet reader:
  `sheetData` read as forward-only events with no tree (the part's text is still held as one string; shared formulas
  expanded, array formulas kept with their range, inline and shared strings, ISO dates), `sheetPr`, `dimension`, `sheetViews`, `sheetFormatPr`, `cols` (with default styles), row
  attributes, `mergeCells`, `hyperlinks` and `extLst` (raw). Theme, indexed and tinted colours now reach cell styles as
  `#rrggbb` (plain black text is left to the editor), so the loss id `xlsx.styles.themeColors` is removed. Workbooks
  from openpyxl with an image, a chart or a comment now import (they made the old reader throw), as do files with a
  shared-string or format index past the table, a damaged worksheet part (read as far as the damage, with a warning) and
  sheets whose names clash (later ones are renamed `Name (2)` with a warning). A shared formula the old reader read as a
  plain value now has its formula, and a saved result of 0 or FALSE is kept. A binary `.xlsb` is refused with a message
  that says to save as `.xlsx`. Additive fields on the result: `workbook`, `styles`, `sharedStrings`, `theme`,
  per-sheet `details` and per-cell `styleIndex`, `sharedString` and `formulaInfo`. Limits against hostile files (what each
  costs is in `documentation/import-loss.md`): a formula over 8192 characters is dropped and shared formulas stop expanding
  past `ImportLimits.maxSheetFormulaChars` (32 Mi per sheet) and `maxFormulaChars` (64 Mi for the whole workbook; the cells
  keep their value, one warning), a workbook is cut off past `maxTotalCells` (5,000,000) and `maxSheets` (1,000, with the
  `<sheet>` entries read capped at four times that), a worksheet part is read once, `<col>` ranges are applied in one sweep
  and capped, merges and hyperlinks are read as events and applied in one sweep (thousands of huge ranges stay fast), at
  most 1,048,576 merges and 65,530 hyperlinks per sheet and `maxRanges` (1,048,576) over the workbook, the per-item
  warnings (a link without its relationship, a repeated sheet) are given for the first twenty and then counted, and the
  small elements cut out of a sheet part are read up to a size (64 KiB each, 8 MiB for the raw `extLst`). The workbook-level
  parts are bounded by what they keep: shared strings, relationships and content types are read as events, and the styles,
  theme and workbook are element trees under a budget (`maxPartElements`, 1,000,000 elements: a part over it is not read, with
  one warning); rows with settings are bounded by `maxRows` (2,097,152); entity-dense text is decoded in windows; a sheet name
  is cut at 255 characters and warnings quote the file's text at 64. A cell or row
  without an address follows the previous one; only a position past the sheet's last row or column is clamped, with a
  warning. A number text over 64 characters is not a number, and the 1900 date quirk (serial 60) is honoured for `t="d"`
  dates. Number format codes a file defines lose the backslash only before a character Excel shows without quotes
  (a LibreOffice `\$#,##0.00` is `$#,##0.00` in the editor); an escape that keeps a letter, digit or `. , @ * _ % # ? " \`
  literal stays (`0\ \d`, `0.0\"`), as does the one that keeps `/` from being a fraction or an AM/PM token (`0\/0`,
  `AM\/PM`) and `+`/`-` from being an exponent sign (`0.00E\+00`); nothing in quotes is touched, an escaped letter is no date
  token, and a date format is recognised in either case (`DD`, `YYYY`). The `exceljs`
  reader dropped every escape; the differential names each code where the two now differ. Numbers are plain decimals only (no
  `Infinity` or hex), `t="d"` dates are validated, `_xHHHH_` escapes are decoded in inline and formula strings, `<f/>`
  without text is no formula, shared-formula names may use letters of any script, `frozenSplit` panes freeze, and
  `activeSheetIndex` follows the sheets that were read. The `exceljs` reader stays in the tree,
  unused, until S14; `reader-differential.test.ts` runs both over the corpus and over generated workbooks. S2 is
  complete with this entry. Review round 5: what a limit cuts is also a loss item (`xlsx.package.sizeLimit`, in `SpreadsheetImportResult.limitLoss`
  and appended by `importFromXlsx` to the `loss` it returns: cells past `maxTotalCells`, sheets, rows, ranges, columns, extension lists,
  formulas, parts and defined names past `maxPartElements`, relationships; each names the option that raises it); new `ImportLimits`
  `maxColumns`, `maxExtensionChars`, `maxRelationships` and `maxScannedElements` (package-wide budgets for `<col>`, `extLst` text,
  relationship entries read and cell, row and column elements scanned); only the relationships the reader uses are kept; defined names are read
  as a list under `maxPartElements` instead of refusing the workbook (a workbook over the element budget says it has too many
  elements); a number format code over 255 characters is kept as written and never scanned (it was rescanned for every cell
  format); every string the result keeps is a copy of its own, so a short `dimension` no longer keeps a 100 MB sheet part alive; the
  workbook part is read once; the cells-over-limit warning names the limit and the `limits` option; a sheet after the cell budget is
  spent is kept if it is empty. The hostile-shape tests run in a child process with a limited heap (skipped below 6 GB or with
  `XLSX_SKIP_HOSTILE`) instead of reading the test process's heap. Review round 6: an element with more than 1,024 attributes
  (`MAX_ATTRIBUTES` in `sax.ts`; Excel writes fewer than 30) is damage like any other (a sheet is read as far as it, a workbook-level part
  skipped with a warning, the workbook part refused as corrupt) where nine million attributes made the engine stop answering; sheets-core
  `importFromXlsx` applies the merges of a sheet in one pass (`WorkbookImpl.mergeCellsBulk` / `SheetImpl.mergeCellsBulk`, the same regions
  and cleared cells as `mergeCells` one by one, one notification, and no undo step per merge: opening a file leaves nothing to undo), 20,000
  merges in under a second where one by one took 51 s and 200,000 in a few seconds; merges that overlap (Excel writes none) are resolved as one by one
  would resolve them, however many (100,000 blocks and one overlap lose only the two it overlaps), also over regions that already overlap each other;
  the loss detector
  finds the relationship of each hyperlink by id instead of by a search (120,000 external links took 45 s); the React editor can raise the
  limits: `importWorkbookFile(workbook, file, { limits })` and the `importLimits` prop of `WorkbookCanvas`.
  Known limitation: the reader is bounded, but loading a result with very many formulas into the editor's model is not (20,000 array formulas over
  blocks took 110 s, 1,000,000 formulas 82 s and 2.1 GB, 1,000 chained sheets 33 s; the reader 0.1 to 0.2 s of each, and a table of two million shared
  strings peaked at 2.8 GB): hosts should lower `maxTotalCells` and/or import in a worker; follow-up in `documentation/sheets/TODO.md`.
- **DOCX reader adopts the safe-import limits (X8).** `importDocx(bytes, { limits })` opens the package through
  the shared `openSafeZip` (entry, size and compression-ratio limits, unsafe part names, encrypted, legacy and
  other-office files recognised) and throws only an `ImportError` (`@b32nio/spindle-io`, also exported from
  `/docx`); `detectDocxLoss` takes the same `limits` and no longer inflates a package over them. The XML parser is
  the shared one (depth cap, DOCTYPE removed, so declared and external entities never expand; the five predefined
  entities are decoded as before). A `.docx` without `[Content_Types].xml` is now refused as `not-a-package`. All
  DOCX goldens are unchanged; two new edge fixtures (`edge-zip-bomb`, `edge-entity-bomb`) cover the bomb and the
  entity expansion. Known gap: docs-core `importFromDocx` does not forward `limits` yet (follow-up).
- **DOCX fidelity harness re-pointed to the shared core.** `scripts/docx-fidelity/` now keeps only the DOCX adapter
  and imports the generic runner, snapshot, manifest, matrix, diff and report modules from `scripts/fidelity-core/`;
  the duplicated copies are deleted. The DOCX goldens and README coverage matrix are byte-identical. The markdown
  report lists the block counts as part of one line ("1 sections, 1 heading, 10 paragraph, 589 textChars") and a
  per-file run is limited to 60 seconds (`--timeout`). No library or runtime behaviour changes.
- **Slides fills (K1).** `Fill` gains `gradient` (linear, radial and path gradients, stops with per-stop alpha),
  `picture` (stretch or tile, alpha, data URLs only; any other source paints nothing) and `pattern` (kept in the
  model, painted as the solid blend of its two colours) on shapes, text boxes, table cells and slide backgrounds
  (layout backgrounds are carried the same way). They are stored in the existing element and slide JSON, so
  collaboration carries them with no Y change. New slides-core exports: `GradientFill`, `GradientStop`, `PictureFill`,
  `PatternFill`, `colorChannels`, `fillToPaint`, `fillToCss`, `paintToCss`, `linearGradientVector`,
  `cssGradientAngle`, `patternBlend`, `patternCoverage`, `isSafePictureSrc`, `GRADIENT_PRESETS`. The shape,
  table-cell and slide-background controls get gradient presets and an image fill. **Compile-time note:** widening
  `Fill` can break a host's exhaustive `switch (fill.kind)`; `resolveFill` now returns the first stop for a gradient,
  the blend for a pattern and `null` for a picture.
- **Slides image crop (K2).** `ImageElement` gains `crop { l, t, r, b }`: the percent (0 to 100) of the source trimmed
  from each edge, kept in the element JSON so collaboration carries it with no Y change. A cropped picture is clipped
  by the element box, and `fit` (stretch, fit, crop to fill) applies to the cropped region; an omitted or all-zero crop
  renders exactly as before. An axis can never be cropped to less than 1 percent. The image format bar gets a "Crop
  image" popover with four percent fields and a reset. New slides-core exports: `ImageCrop`, `normalizeCrop`,
  `isCropped`, `setCropEdge`, `cropLayout`, `NO_CROP`, `MIN_CROP_VISIBLE`, `CropLayout`, `PctRect`. Additive.
- **Slides text body fidelity, part 1 (K3a).** `BodyStyle` gains `insets` (px per side; `padding` still works and
  `insets` wins), `anchorCtr` (centre the text block), `autofit { mode: none | shrink | resize, fontScale, lnSpcReduction }`,
  `vert: horz | vert | eaVert` and `columns { count, gap }`, all in the element JSON so collaboration carries them with no
  Y change. A `shrink` body applies the **stored** scales (as PowerPoint's `normAutofit` does) and recomputes them only
  when the user edits the text (a 2.5 percent step search against the live layout); a `resize` body grows or shrinks its
  height to the text on commit; a body that declares autofit is no longer clipped (an unset body keeps clipping, so
  existing decks render as before). Text and shape bodies render all of it, and the live editor lays text out the same
  way. New "Text box options" toolbar control (popover, or a row and modal in the overflow menu). New slides-core
  exports: `Autofit`, `BodyInsets`, `TextVert`, `resolveInsets`, `insetsToCss`, `autofitMode`, `autofitScales`,
  `isScaled`, `hasAutofit`, `textVert`, `writingModeCss`, `resolveColumns`, `fitShrinkScale`, `lineReductionFor`,
  `mergeBodyStyle`. The editor schema's font sizes and line heights are now `calc(... * var(--spindle-fs / --spindle-ls))`
  expressions (same computed values at scale 1). Additive; paragraph, run and field attributes follow in K3b and K3c.
- **Slides paragraph fidelity, part 2 (K3b).** Paragraphs gain `bullet { kind: none | char | auto, char, startAt, scheme,
  color, sizePct }` (overrides `listType`; schemes arabic, alpha and roman in both cases), `lineHeightPx` (exact line
  height; `lineHeight` stays the multiple), `spaceBeforePct` / `spaceAfterPct` (percent of the text size; `spaceBefore` /
  `spaceAfter` stay px), `marginLeft` and `firstLineIndent` (px; a negative first-line indent hangs the marker) and
  `tabStops` (stored and synced, laid out with default tabs). Editor and static render share one resolver, so edit mode
  and idle mode agree; untrusted values are coerced or allowlisted before they reach CSS, and the bullet glyph is only
  ever an attribute. The "Text box options" panel gets a "Paragraph options" section (bullet, numbering, size, colour,
  line height, spacing, margins). **Collaboration:** the new attributes ride on the paragraph elements of the rich-text
  Y fragment, so slides `COLLAB_SCHEMA_VERSION` is now 2 and an older client refuses a room written by this one
  (`CollabSchemaVersionError`) instead of dropping them. New slides-core exports: `Bullet`, `BulletKind`, `BulletScheme`,
  `TabStop`, `LevelStyles`, `resolveParagraph`, `flattenLevelStyles`, `formatOrdinal`, `normalizeTabStops`,
  `safeBulletChar`, `INDENT_PX`, `DEFAULT_BULLET_CHAR`, `SCHEME_COUNTER_STYLE`. Additive.
- **Slides run marks and fields (K3c).** Text runs gain four marks: `baseline` (percent raised or lowered, drawn as a
  smaller shifted run that does not change the line height), `caps` (`all` or `small`), `charSpacing` (px) and `lang`
  (a BCP 47 tag). A new inline node `field` (slide number, date and time in three layouts, footer text) holds its last
  shown text, which carries the run's formatting; the renderer shows the live slide number (the slide's place in the
  deck, also in the filmstrip and present mode) or the current date. The "Text box options" panel gets a "Run options"
  section (baseline, capitals, character spacing, language, insert field while editing). Every value is coerced,
  clamped or allowlisted before it reaches CSS or the DOM. **Collaboration:** the marks and the field node ride on the
  rich-text Y fragment, so slides `COLLAB_SCHEMA_VERSION` is now 3 and an older client refuses the room
  (`CollabSchemaVersionError`). New slides-core exports: `RichTextText`, `RichTextField`, `FieldKind`, `FieldAttrs`,
  `DateFieldFormat`, `FieldContext`, `ResolvedRun`, `CapsKind`, `inlineText`, `resolveRun`, `runCss`, `baselineCss`,
  `resolveFieldText`, `initialFieldText`, `fieldMarks`, `safeLang`, `safeFieldKind`, `BASELINE_SCALE`,
  `SLIDENUM_PLACEHOLDER`. `RichTextInline` is now the union `RichTextText | RichTextField` (hosts that read
  `inline.text` should narrow on `type`). Additive otherwise.
- **Slides layouts, slide size and hidden slides (K4).** `LayoutData` gains `elements` (static, non-placeholder content
  from the layout and its master: logos, rules, footers, drawn under the slide's own elements and never selectable),
  `textStyles` (per placeholder type, per paragraph level 0 to 8, with level-to-level inheritance via
  `resolveLevelStyle`; level 0 is baked into a placeholder when a slide is created from the layout) and `masterName`.
  `Slide.hidden` (PPTX `show="0"`): a hidden slide stays in the deck, is dimmed with a badge in the filmstrip (right-click
  menu: Hide slide / Show slide) and is skipped by present mode (`PresentMode` takes `includeHidden`). The "Theme"
  popover (and its overflow modal) gains a **Slide size** section: presets 16:9, 16:10, 4:3, A4 landscape and portrait,
  portrait 9:16, plus custom width and height; `DeckImpl.setSlideSize(size, { scaleContent: true })` scales every
  element, line end point and layout per axis (one undo step) so content keeps its relative position. Importers still
  flatten master to layout to slide inheritance into effective values (decision SD6); there is no master editor.
  **Collaboration:** `meta.layouts` (one LWW JSON blob, so two users editing layouts at once keep one edit) now carries the
  new fields and slide maps carry `hidden`, so slides `COLLAB_SCHEMA_VERSION` is now 4 and an older client refuses the
  room. New slides-core exports: `LevelTextStyle`, `LayoutTextStyles`, `resolveLevelStyle`, `SLIDE_SIZE_PRESETS`,
  `SlideSizePreset`, `SlideSizePresetId`, `detectSlideSizePreset`, `clampSlideSize`, `scaleElementPatch`, `scaleLayout`,
  `MIN_SLIDE_SIDE`, `MAX_SLIDE_SIDE`. Additive.
- **Slides collaboration gate.** Slides rooms now record `meta.schemaVersion` (missing key = 1; no migration), and
  `attachCollab` refuses a room written by a newer client with a `CollabSchemaVersionError` instead of editing it.
  New additive exports from `@b32nio/spindle-slides-core`: `getCollabBlockers(deck)` (empty for now) and
  `CollabBlockedError`; the `/collab` subpath also exports `COLLAB_SCHEMA_VERSION`, `CollabSchemaVersionError`,
  `readSchemaVersion` and `assertSchemaVersion`.
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
- **Excel number-format engine and `TEXT()` date codes (S3, engine part).** New module
  `packages/sheets-core/src/number-format/` implements the ECMA-376 format-code grammar: sections and conditions,
  colours, locale and currency tokens, digit placeholders, thousands and scaling commas, percent, exponent,
  fractions (variable and fixed denominators), literals, `_x` padding and `*x` fill, `@`, Excel's General
  (11 characters, E notation), dates and times (`[h]` elapsed, fractional seconds, minutes versus months,
  `[$-F800]`) in the 1900 system (with its 29 February 1900) and the 1904 system, negative dates as `####`, rounding
  half away from zero on 15 digits, Excel's 255-character, five-denominator-digit and three-fractional-second limits. `formatValue`, `formatDisplay`, `formatGeneral`,
  `compileFormat`, `isValidFormatCode`, `isDateTimeFormatCode`, `serialToDate` and `dateToSerial` are exported
  (additive); codes are compiled once and cached. It replaces the `formatCustomPattern` stub: the custom, scientific,
  fraction, date, time, datetime and duration format types and unformatted cells now render through it (unformatted
  numbers use General, so `0.1+0.2` shows `0.3`; the duration types show `[hh]:mm:ss`, `[mm]:ss`, `[ss]` and
  `[ss].000` of a value in days instead of `5h`; scientific shows `1.23E+03` instead of `1.23e+3`).
  `TEXT(value, code)` uses it: `TEXT(date,"yyyy-mm-dd")`, `"dddd"`, `"mmm d, yyyy"`, `[h]:mm`, sections, conditions,
  scaling and fractions work; an empty code gives `""` and an invalid code `#VALUE!`. A conformance table of 1,023
  (code, value) rows with LibreOffice's display strings (`fixtures/xlsx/generators/build_numfmt_oracle.py`, the
  33-code probe first, then the 60-code corpus, ECMA-style examples, conditions, literals, General, dates, times and 1904
  dates) is checked by Jest. The engine agrees with LibreOffice alone on 802 rows (78.4%); the other 221 rows carry a
  correction where Excel is believed to differ (negative dates `####`, the 1900 calendar before March 1900, `-0.00`
  keeping its sign, General's width, ...) with the reason, and 51 of them (time rounding, `am/pm` case) plus 9 condition
  rows are tagged `unverified` pending real Excel files; the Excel strings are self-authored, not an independent oracle.
  `TEXT(x,"General")` gives 15 significant digits and `TEXT` reads `"$5"`, `"1,234"`, `"50%"`, `"(5)"` and date text like Excel.
  The `formulas-core` XLSX goldens record the two `TEXT` formulas that now match (probe values 110 to 112 of 159).
  Mapping imported codes onto the model follows in the next change.
- **Number formats survive XLSX import and export exactly (S3, mapping part).** An imported number-format code that maps
  exactly onto a structured type (the type writes the very same code back) keeps it; every other code is stored verbatim as
  `custom` with the original text, so `d-mmm-yy`, `dddd, mmmm d, yyyy`, `0.0,,"M"`, `# ?/?`, `[h]:mm:ss`, `[$€-407] #,##0.00`,
  `[>100]"big";"small"`, `0.0 "kg"` and the rest of the 60-code corpus now round-trip unchanged (probe 5: 16 of 33 codes survived,
  now 33 of 33; corpus 60 of 60). The old fallback of any date-looking code to `MM/DD/YYYY` (and of `#,##0.0" kg"` to a plain number)
  is removed. `fromExcelNumberFormatCode`, `toExcelNumberFormatCode`, `normalizeExcelNumberFormatCode` and the en-US table of the
  built-in ids (`BUILTIN_NUMBER_FORMAT_CODES`, 0 to 22, 27 to 58) are exported; `CellFormat` and the collab format pool are unchanged.
  The loss predicate behind `xlsx.numberFormats.approximated` now means "the engine cannot read the code or would show another
  language's month names", so the id drops from 66 occurrences in 3 corpus files to 0 (`number-formats` 46 to 0, the probe 19 to 0,
  `cell-types` 1 to 0). The XLSX model goldens of those three files show `custom` formats with the original codes, and their
  round-trip goldens lose the `currencyCode` differences. The Format Cells dialog gains accessible markup, test ids,
  Excel codes as common patterns and an inline message plus a disabled Apply for a code Excel would reject; autofill treats a custom
  code with a year, month or day as a date (`isDateFormat`; not `h:mm` or `[h]:mm`), and the dialog applies the code untrimmed. The
  loss label and effect for `xlsx.numberFormats.approximated` now say the code is kept and saved unchanged but shown in General or
  with English names (`packages/io/src/xlsx/loss.ts`, text only). E2E `S-FMC-06` opens a workbook with currency, date, time, elapsed, scaled and conditional
  formats and reads the dialog's sample; `S-FMC-04b` (custom suffix, issue 83) is no longer `fixme`, and `S-FIO-01` now expects the
  verbatim `$#,##0.00` and `#,##0.0" kg"` codes.
- **Formula conformance suite and function batch A (S11).** `formula-conformance-oracle.json` (generated by
  `fixtures/xlsx/generators/build_conformance_oracle.py` from LibreOffice 25.2 recalculation) lists the 159-formula
  probe and 279 new batch-A rows; `formula-conformance-oracle.test.ts` runs them all through the engine and keeps an
  explicit allow-list of the rows that still fail, each with a reason. The engine gains 47 functions (146 to 193):
  `SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, the hyperbolic set, `PI`, `RADIANS`, `DEGREES`, `FACT`,
  `FACTDOUBLE`, `COMBIN`, `PERMUT`, `SUMSQ`, `EVEN`, `ODD`, `ISEVEN`, `ISODD`, `CEILING.MATH`, `FLOOR.MATH`,
  `MAXIFS`, `MINIFS`, `AGGREGATE` (functions 1-19, reference and array forms, error-ignoring options),
  `STDEV.S/.P`, `VAR.S/.P`, `PERCENTILE.INC/.EXC`, `QUARTILE.INC/.EXC`, `RANK.EQ`, `MODE.SNGL`, `SLOPE`,
  `INTERCEPT`, `RSQ`, `FORECAST`, `FORECAST.LINEAR`, `GEOMEAN` and `FREQUENCY`. Scalar functions of the batch lift
  over ranges and arrays. `MODE` now returns the first of equally frequent values and `CORREL`/`COVAR` propagate
  error values, as Excel does. The XLSX `values` golden for `formulas-core` improves from 112 to 125 of 159
  matching and `xlsx.formulas.unsupported` shrinks (34 to 15 formulas there); `_xlfn.`-prefixed spellings still
  wait for S4.
- **Function batch B, text and date (S12a).** 18 new functions (193 to 211): `CLEAN`, `NUMBERVALUE`, `UNICHAR`, `UNICODE`, `T`,
  `FIXED`, `DOLLAR`, `BASE`, `DECIMAL`, `ROMAN` (five forms), `ARABIC`, `ISNONTEXT`, `DAYS`, `DAYS360`
  (US and European), `YEARFRAC` (all five bases), `ISOWEEKNUM`, `WORKDAY.INTL` and `NETWORKDAYS.INTL` (weekend codes,
  seven-character strings, holidays, constant time). `TEXTBEFORE`, `TEXTAFTER` and `TEXTSPLIT` are deferred until Excel-computed
  files exist (no trustworthy oracle; plan D8). `INDEX(range,0,col)`, `INDEX(range,row,0)`, `INDEX(range,,col)` and
  `INDEX(table,row)` return a whole column or row (also as a range endpoint); a negative index is `#VALUE!`. The parser keeps an empty
  argument in place for the functions that take one (`INDEX(a,,2)`, `NUMBERVALUE(a,,",")`). Date functions now follow Excel's 1900
  calendar (29 February 1900 is serial 60, `DATE(1900,1,1)` is 1, overflow such as `DATE(1900,3,0)` and `EDATE`/`EOMONTH`/`DATEDIF` follow it). 220 new conformance rows
  (LibreOffice 25.2; the rows where it differs from Excel carry Excel's documented value, tagged unverified until real Excel files confirm it). The published list is `documentation/sheets/core/functions.md`. The XLSX `values` golden for
  `formulas-core` improves from 125 to 127 of 159 matching and `xlsx.formulas.unsupported` shrinks (15 to 10 formulas there; the
  `_xlfn.`-prefixed ones wait for S4).
- **Function batch B, finance, math and TRANSPOSE (S12b).** 18 new functions (211 to 229): `TRANSPOSE` (spilled; a blank cell that passes
  through reads 0), `XNPV`, `XIRR` (Newton from the guess, bracketed search as the fallback: a deliberate superset of Excel), `MIRR`, `SLN`,
  `SYD`, `DDB`, `EFFECT`, `NOMINAL`, `RRI`, `PDURATION`, `ISPMT`, `FVSCHEDULE`, `QUOTIENT`, `SQRTPI`, `SUMX2MY2`, `SUMX2PY2` and `SUMXMY2`.
  118 new conformance rows (LibreOffice 25.2 computes them; where it reports `#VALUE!` for an argument Excel calls `#NUM!` or `#DIV/0!`, the
  row carries Excel's documented value, tagged unverified). `LET` and `LAMBDA` stay post-launch. The XLSX `values` golden for `formulas-core`
  improves from 127 to 130 of 159 matching and `xlsx.formulas.unsupported` shrinks again (10 to 6 formulas there); `=TRANSPOSE(A1:B2)` in that
  fixture evaluates but its spill range is blocked by the next probe row (`#SPILL!`).
- **Safe import: limits and a named error (X4).** `importXlsx` (`@b32nio/spindle-io/xlsx`) now opens a workbook through
  `common/safe-zip.ts` and only ever throws an `ImportError` with a stable `code`: `encrypted` (an OLE container
  holding an encrypted package, or an encrypted zip entry), `legacy-binary` (`.xls`, `.ppt`, `.doc`, `.xlsb`),
  `not-a-package` (empty, another format, OpenDocument, Apple iWork), `size-limit` or `corrupt` (truncated, unreadable,
  unsafe part name). Limits are checked from the zip's metadata before anything is inflated and again while every part
  is inflated (a forged header cannot hide a bomb): package 100 MB, 20,000 entries, 1 GB in total, 500 MB per part
  (256 MB for XML), compression ratio 200:1 for anything over 16 MB; all are options (`importXlsx(bytes, { limits })`,
  `importFromXlsx(bytes, workbook, { limits })`). The XML parser settings (entities never expanded, DOCTYPE removed,
  depth cap 256; the loss detector's nesting cap rises from 100 to 256 with it) moved to `common/safe-xml.ts`. CSV import drops a leading byte order mark and refuses text over 100 MB
  (`importFromCSV(csv, sheet, { maxChars })`; the Open action checks the file size before reading it); The CSV refusal is an error with `name` `ImportError` and `code` `size-limit`
  but not an instance of the io package's class (CSV is synchronous and core does not load it): use `isImportError(e)`
  (`@b32nio/spindle-sheets-core`) or match on `code`. `.xls` handed
  to Open now gets the legacy-binary message instead of "unsupported type". Loss ids are the ones that already exist
  (`xlsx.package.encrypted`, `.legacyBinary`, `.sizeLimit`, the last now also for a file over the package limit). The
  `importError` text of six edge fixtures and three openpyxl fixtures in the XLSX goldens changes (nothing else does).
  **Behavioural:** `importXlsx` / `importFromXlsx` used to fail with whatever error JSZip or exceljs raised; hosts that
  matched on those messages should read `error.code` instead. The DOCX reader is not changed here.
- **Slides: preset geometry library (K5).** `ShapePreset` grows from 18 to 76 presets named after OOXML `prst`
  values: snipped, rounded and folded rectangles, frame, plaque, bevel, can, cube, donut, no-symbol, teardrop, pie,
  chord, block arc, moon, smiley, lightning bolt, heptagon to dodecagon, block arrows, 4/6/8/12-point stars, three
  callouts, eleven flowchart shapes, brackets and braces, plus/minus and the inverted line. Adjustments are fractions
  (OOXML `adj` / 100000), defaulted and clamped per preset (`resolveAdjustments`). A preset outside the library is
  kept as an `unknown:<prst>` shape (`ShapeKind`, `unknownShape`, `isUnknownShape`, `shapePrst`, `shapeFromPrst`) with
  its adjustments, drawn as a rounded rectangle and left intact for re-export. The Shapes menu groups the presets into
  categories with labelled buttons. Additive exports from `@b32nio/spindle-slides-core`; `ShapeElement.shape` widens
  from `ShapePreset` to `ShapeKind`. Slides collaboration schema version 5 (a v4 client refuses a v5 room). The PPTX
  loss id `pptx.shapes.unsupportedGeometry` now counts only presets outside the library.
- **Slides shape adjustment handles (K5b).** A single selected shape with adjustable guides (rounded and snipped
  corners, arrows, stars, callouts, brackets, pie, chord, block arc, moon, can, cube and more) shows yellow diamond
  handles on the canvas; dragging one changes the guides live and commits one undo entry. The handles are drawn in
  the selection overlay (labelled "Adjust shape"), so the toolbar is unchanged. Guides are already stored in
  `props.adjustments`, so there is no new collaboration key and the schema version stays 5. New additive slides-core
  exports `adjustmentHandles` and `adjustmentsAfterDrag`; presets without a handle for a guide (the block arc
  thickness) keep editing it through the data only.
- **Slides nested groups (K6a).** A group can sit inside another group: Group on a selection that includes a group
  wraps the whole group, Ungroup dissolves the outermost group and leaves the inner ones, and a click selects the
  outermost group. Elements carry `parentGroupId` beside `groupId` (additive; `groupChain`, `groupParents`,
  `rootGroupOf` in `scene/group-tree.ts`, `DeckImpl.getGroupChain` / `getRootGroupId`, and `getGroupMembers` now
  includes nested members). Duplicating a slide or elements and copy/paste keep the nesting under fresh group ids.
  `childFrameInParent` and `composeGroupChain` compose DrawingML group transforms (scale, flips, rotation) for the
  PPTX reader. Slides collaboration schema version 6 (a v5 client refuses a v6 room).
- **Slides elbow and curved connectors (K6b).** A line can be routed as an elbow (right-angle legs, middle leg at
  `adjustments[0]`, default 0.5) or a curve: new optional `LineElement.connector` (`'straight' | 'elbow' | 'curved'`,
  type `ConnectorKind`) and `adjustments`. Bound ends leave their shape perpendicular to the bound side (`anchorDirection`,
  `routeConnector`, `lineRoute`, `routePathD`). Connection sites follow the shape: an ellipse's diagonal sites sit on
  the outline, a diamond offers its four edge midpoints and a triangle its apex, slanted-side midpoints and base
  points (`connectionAnchors`, `connectionFrame`, `ConnectionFrame`; the eight anchor ids are unchanged, so existing
  bindings keep working; sites mirror with a shape's or picture's flip, so a connector bound to a flipped rectangle's `nw` now attaches at its `ne` corner, matching how it is drawn and OOXML). Choose the routing from the right-click menu of a selected line. Slides collaboration schema
  version 7 (a v6 client refuses a v7 room).
- **Slides custom geometry (K6c).** A shape can carry an arbitrary outline: `ShapeElement.shape` may be `path` (type
  `PathShape`, part of `ShapeKind`) with `ShapeElement.path` (`CustomGeometry`: one or more `CustomPath`s of
  move/line/cubic/quad/arc/close commands, per-path size, `norm`/`none`/`lighten`/`darken` fill modes and a stroke
  flag). It is drawn scaled onto the box (arcs, several paths, shaded fills) and survives collab and copy/paste.
  New slides-core `scene/custom-geometry.ts`: `parseCustGeom` and `serializeCustGeom` read and write `a:custGeom`
  on a plain element tree (for the PPTX reader and writer), `normalizeCustomGeometry` validates untrusted data with caps
  on paths and commands. Guide formulas are not evaluated: a geometry using them stays an `unknown:custGeom`
  passthrough. Slides collaboration schema version 8 (a v7 client refuses a v8 room).
- **Slides tables: merge, per-cell borders, style flags (K7).** `TableCell` gains `gridSpan`, `rowSpan`, `hMerge`,
  `vMerge` (OOXML `a:tc` semantics: the grid stays rows x cols, the top-left cell of a merged region is its origin, the
  cells it covers are not drawn) and `borders` (`l`/`t`/`r`/`b`, a `Stroke`, `null` for no line, absent to inherit the
  grid line); `TableElement` gains `styleFlags` (first/last row and column, banded rows and columns, kept for re-export;
  the resolved fill and borders live on each cell). Cell margins and vertical anchor are `bodyStyle.insets` and
  `bodyStyle.vAlign`, now honoured in the cell. New slides-core `scene/table.ts` exports `mergeCells`, `splitCell`,
  `setCellBorders`, `resolveCellEdges`, `spanRects`, `spanAt`, `expandRange`, `normalizeSpans`; the deck gains
  `mergeTableCells`, `splitTableCell` and `setTableCellBorders`, and row/column insert and delete keep merged regions
  intact. The cell right-click menu gains Merge cells, Split cell and border commands. Slides collaboration schema
  version 9 (a v8 client refuses a v9 room); table cell text stays last-writer-wins, not CRDT-merged.
- **PPTX reader, step 1 (K9a): package, presentation and theme.** New `importPptx(bytes, { limits })` in
  `@b32nio/spindle-io/pptx` (also from the package root): opens the package through `openSafeZip`, finds the main part
  through the package relationships (falling back to the content type and to `ppt/presentation.xml`), and returns a
  `PresentationImportResult` (new additive types in `@b32nio/spindle-shared/export/pptx`: `PptxSlide`, `PptxTheme`,
  `PptxColorMap`, `PptxColor`, `PptxSection`). It reads the slide size (EMU to px, 1 px = 9525 EMU), the slide list in
  presentation order (`pptxId`, name, hidden flag, layout part), sections, the first master's theme (12 colours, major and
  minor fonts) and colour map. A slide that is missing or unreadable stays in the list as an error placeholder with a
  warning; the only error thrown is an `ImportError`. Slide content, layouts and the editor's `importFromPptx` follow
  (K9b, K9c, K15), so nothing in the editor opens a `.pptx` yet and the loss report is unchanged. The XML tree
  (`xlsx/xml.ts`) gains `rattrs` (attributes in the relationships namespace) so `r:id` is no longer hidden by `id`. The
  fidelity harness (`npm run slides:pptx-fidelity`) now records the imported model and `ledger` mapped counts; the four
  files that are not presentations carry `expectImportError`, and the run is strict about it.
- **PPTX reader, step 2 (K9b): masters and layouts.** `importPptx` now also returns `layouts` (`PptxLayout`, new additive
  types `PptxLayout`, `PptxPlaceholder`, `PptxFrame`, `PptxLevelStyle`, `PptxLevelStyles` and `PptxBullet` in
  `@b32nio/spindle-shared/export/pptx`): every layout of every master, flattened per decision SD6. Each layout placeholder
  carries its effective frame (the layout's own, else the matching master placeholder's), vertical anchor and the nine
  paragraph levels merged down the chain (presentation default text style, master `titleStyle` / `bodyStyle` /
  `otherStyle`, master placeholder `lstStyle`, layout `lstStyle`: size, bold, italic, colour, font, alignment, margins,
  line height, spacing, bullets). Also read: layout and master names, the layout type, a solid or referenced background
  (layout, else master, with the layout's colour map override applied), `showMasterSp`, and a count of the static shapes
  the layout and master draw (their content comes with the shape reader, K9c). A second master with a different theme is
  reported. An internal `matchLayoutPlaceholder` resolves a slide placeholder to its layout placeholder for K9c. No editor change and
  the loss report is unchanged; only the `model` goldens grew.
- **PPTX reader, step 3 (K9c): slide shapes and text.** `importPptx` now reads the `p:sp` shapes of every slide
  (`PptxSlide.elements`, new additive types `PptxShape`, `PptxFill`, `PptxLine`, `PptxGeometry`, `PptxBodyProps`,
  `PptxTextBody`, `PptxParagraph`, `PptxRun`, `PptxFonts`): element id (`pptxId`), name, alt text, hidden flag, frame
  (px, rotation, flips; a placeholder without its own position takes the layout's), preset geometry with its adjustments
  (a preset outside the library is marked `known: false`, per K5) and `custGeom` marked as custom, fills (none, solid,
  gradient, pattern, picture reference) and outlines (width, dash, cap, join, line ends), `p:style` fill, line and font
  colour references, the slide background, and text bodies with all K3 attributes (insets, anchor, vertical text, columns,
  stored `fontScale` / `lnSpcReduction`, bullets with colour, size and font, tab stops, spacing, baseline, caps, character
  spacing, language, fields and breaks), each paragraph carrying its effective level style (layout placeholder or default
  text style, shape list style, own properties). Layouts gain `elements`: the static shapes of master and layout in drawing
  order. The result reports its fonts (`fonts`, with a warning for those outside the editor's list: `pptx.fonts.unavailable`),
  and a slide with an unclosed element is now a damaged slide instead of a partly read one (`OoxmlPackage.xmlChecked`).
  The fidelity harness gains the automated text parity check (every source `a:t` run of the shapes read so far appears
  in the model, as a multiset; 536 of 536 in the corpus) in the `model` goldens. Pictures, tables, groups, connectors and
  custom geometry paths follow (K10). No editor change; the loss report effects of the ids now read say `kept` or `approximated`
  and the rest stay `lost: not imported` until the editor can open a `.pptx` (K15), so the `loss` goldens changed
  along with the `model` goldens.
- **PPTX reader, step 4a (K10a): pictures, tables and hyperlinks.** `importPptx` now reads `p:pic` pictures
  (`PptxPicture`: the embedded image as a data URL typed from its bytes, crop in percent, transparency, frame, outline,
  placeholder frame; an SVG with a PNG fallback reads as the PNG; EMF/WMF and linked pictures become labelled
  placeholders, never fetched) and table `p:graphicFrame`s (`PptxTable`, `PptxTableCell`: spans and merge markers,
  per-cell borders, margins and anchors, with the table style of `ppt/tableStyles.xml` resolved into each cell), and
  hyperlinks on runs, shapes and pictures (`PptxLink`): only `http`, `https`, `mailto` and `tel` addresses are kept,
  `ppaction://` and `file:` targets are kept inert, any other scheme is dropped with a warning. New additive types
  `PptxElement`, `PptxPicture`, `PptxTable`, `PptxTableCell`, `PptxCellBorders` and `PptxLink` in
  `@b32nio/spindle-shared/export/pptx`; `PptxSlide.elements` widens to `PptxElement[]`; `OoxmlPackage.bytes`. The text
  parity check now covers pictures and table cells (603 of 603 runs in the corpus). The loss effects of
  `pptx.images.crop`, `.svg`, `.linked`, `.unsupportedFormat`, `pptx.tables.merge`, `.styles`, `.banding`, `pptx.actions`
  and `pptx.externalLinks` now say what the reader does. Groups, connectors, custom geometry paths, notes and comments
  follow (K10b, K10c). No editor change.
- **PPTX reader, step 4b (K10b): groups, connectors and custom geometry.** `importPptx` now reads `p:grpSp` groups
  (`PptxGroup`: a `children` tree whose frames are already composed through the group transforms (`a:chOff` and
  `a:chExt` scale, rotation, flips, nested groups) into slide space), `p:cxnSp` connectors (`PptxConnector`: the
  straight, bent or curved route with its adjustments, outline and arrow heads, and the shapes the ends are glued to
  with `stCxn` and `endCxn`) and `a:custGeom` paths (`PptxGeometry` custom gains `paths`: move, line, cubic and quad
  curves, arcs, close, several paths, fill modes; a path that needs guide formulas stays without paths). A `grpFill`
  shape now takes the fill of its group. The group composition (`childFrameInParent`, `composeGroupChain`,
  `GroupXfrm`) moved to `@b32nio/spindle-shared` and `slides-core` re-exports it unchanged. New additive types
  `PptxGroup`, `PptxConnector`, `PptxConnection`, `PptxCustomPath` and `PptxPathCommand`; `PptxElement` widens. The text
  parity check now covers groups at any depth (614 of 614 runs in the corpus; one `mc:Fallback`
  copy and the two runs of the damaged slide are set aside). The loss effects of `pptx.shapes.custGeom`, `.groups.nested`
  and `.connectors.elbowCurved` now say what the reader does. The box guides `ss` and `ls` and the angle constants
  (`cd2`, `cd4`, `cd8`, `3cd4`...) are read in custom geometry. Notes and comments follow (K10c). No editor change.
- **PPTX reader, step 4c (K10c): speaker notes and comments.** `importPptx` now reads the body placeholder of each slide's
  notes page into `PptxSlide.notes` (a text body with the notes' run and paragraph formatting, its styles inherited from the
  notes master's `p:notesStyle` and body placeholder) and comments into `PptxSlide.threads` (`PptxCommentThread` and
  `PptxComment`, new additive types): legacy `ppt/comments/*.xml` (author list, position, replies by `p15:parentCm`) and
  modern comments (authors part, replies, `resolved` and `closed` status, anchored to the slide or to the shape a
  `*MkLst` marks). `pptx.comments` now counts only comments parts no slide points at (`lost`), `pptx.notes` narrows
  to notes with formatting in the notes body (`approximated`) and `pptx.masters.notesMaster` is `approximated`; the summary ids
  stay 16. Notes or comments of a slide that cannot be read are dropped with a warning, not the whole slide. The text parity
  check covers the notes body runs (a notes page shared by several slides once for each). No editor change.
- **PPTX reader C: opaque content (K11a).** `importPptx` now keeps charts (with their embedded workbook and style parts), SmartArt
  diagrams, OLE objects and ActiveX controls, audio and video clips, and ink as opaque elements, and a slide's transition and timing XML,
  and stores the parts byte for byte in the opaque part store (`PresentationImportResult.opaqueParts`; new `PptxOpaque`, `PptxOpaqueKind`,
  `PptxSlideOpaque`; `PptxElement` gains `PptxOpaque` and `PptxSlide` gains `opaque`; `ImportPptxOptions.opaqueLimits`). A part over the
  cap is dropped with a warning and reported as a `lost` item under its own id (`pptx.media` for a clip, likewise `pptx.charts`, `.smartArt`, `.embeddedObjects`, `.ink`); the loss ids `pptx.charts`, `.smartArt`, `.media`,
  `.embeddedObjects`, `.transitions`, `.animations` and `.ink` now say `kept`. An equation is read from its `mc:Fallback` when there is one (the text it holds, or a picture). In
  `@b32nio/spindle-slides-core`: new `OpaqueElement` (`type: 'opaque'`, added to `SlideElement`), `ElementBase.pptxId`, `Slide.opaque`,
  `DeckData.opaqueParts`, `DeckImpl.getOpaqueParts()` / `getOpaquePart(hash)`, `createOpaqueElement`, `opaqueLabel`, `pruneTiming`;
  deleting an element prunes the timing that animates it. `@b32nio/spindle-slides-react` draws an opaque element as a labelled placeholder
  ("Chart: <title> (kept, not editable)"). Slides collaboration schema version 10 (`opaque` slide key, `opaqueParts` map, opaque elements):
  older clients refuse such rooms. Hosts that switch exhaustively on `SlideElement['type']` need an `'opaque'` case.
- **XLSX formula interchange and cached values (S4).** Behaviour change: an opened .xlsx now shows the results the file saved and
  recalculates a formula when a cell it refers to changes or it is edited, not wholesale on load (formulas without a saved result,
  volatile ones and array formulas that spill are still evaluated on load); the host option `recalculateOnLoad` (default off; an
  `importFromXlsx` option, an `importWorkbookFile` option and a `WorkbookCanvas` prop) evaluates every formula the engine can on load.
  `_xlfn.`, `_xlfn._xlws.` and `_xlpm.` prefixes are read as the plain function names the engine knows (`_xlfn.IFS`, `_xlfn.XLOOKUP`,
  `_xlfn.STDEV.S`...) and written back on export; `_xlfn.SINGLE(x)` and `_xlfn.ANCHORARRAY(A1)` are read as `@x` and `A1#` and
  written back as the calls (`fromFileFormula`, `toFileFormula` and `FUTURE_FUNCTIONS` are new exports of `@b32nio/spindle-io/xlsx`).
  A formula the engine cannot evaluate (a function it lacks, text it cannot parse, a table, defined-name or external reference) is
  never shown as `#NAME?`: the cell keeps the file's formula text and the saved result (`cachedOnly`), shows a small "Saved value" note
  in the formula bar (the cell keeps its saved value and its formula text: the value does not follow the cells it refers to; Excel recalculates it
  when the file is reopened, possibly to a different result; the note, its tooltip, the loss effect and `documentation/import-loss.md` say so; the
  references in the text follow inserted and deleted rows and columns since S4b below), and is written back with the saved value and the formula kept (S4 wrote it back unchanged; since S4b the references in its text follow insert and delete). The note's colours are the CSS variables `--spindle-note-fg`, `--spindle-note-bg` and `--spindle-note-border` (defaults: at least 4.5:1 contrast, about 6.4:1 for the text and background pair, and 4.5:1 for the border), so a dark host theme can set its own. Array formulas keep their range: a legacy array formula keeps its extent, a dynamic-array one
  spills, a one-cell one stays an array formula, and export writes `t="array"` with the range and the cell metadata for dynamic arrays.
  A legacy (CSE) array formula whose engine result differs from the results the file saved (a blank source cell reads 0 in the result, as in Excel, so it is the same as a saved 0 and only that; numbers are held to a relative 1e-9) becomes cached-only and keeps the file's values, counted in
  `xlsx.formulas.cachedOnly`; with `recalculateOnLoad` the formulas that read it are recalculated from those values. A dynamic array is not
  compared: the engine's spill replaces the saved one (known divergence: it spills a blank source cell as blank where Excel shows 0, so `=C2&"x"` over it gives `"x"` instead of `"0x"`; recorded in the sheets TODO). Array results are cut by the export's budget in a fixed order (sheet, then row, then
  column; a cut array keeps its anchor), an error in a spilled result is written as an error cell (`t="e"`), and so is an error saved with a
  legacy array's members, found within the same `maxTotalArrayCells` budget (past it they stay the text of the error).
  New loss id `xlsx.formulas.cachedOnly` (state `kept`); `xlsx.formulas.unsupported`, `.structuredRefs` and `.externalRef` now say the
  value is kept. Additive: `Cell.cachedOnly` and `Cell.arraySpan`, `SpreadsheetCell.cachedOnly`, `.arrayRef` and `.dynamicArray`,
  `XlsxFormulaInfo.dynamic`, `DetectXlsxLossOptions.isFormulaEvaluable`, `ImportFromXlsxOptions`, `WorkbookImpl.getSpillRegion`. The sheets
  collaboration layout is now version 2 (a client at version 1 refuses a version 2 room); `documentation/collaboration.md` has the rules.
  The values golden counts `cachedOnly` formulas where it counted `passthrough` ones, and measures the engine with a recalculating import.
  A formula whose result is an error (`#DIV/0!`, `#N/A`, `#NAME?`, `#SPILL!`, `#CALC!`...) is written as an error value (`t="e"`), not as text,
  for cached-only and ordinary formulas alike. A legacy array formula whose range does not start at its own cell is read as a one-cell
  formula, with a warning; a one-row or one-column result is repeated across a bigger array-entered range, as Excel does. The `_xlfn.` table gains
  `ENCODEURL`, `VALUETOTEXT`, `IMAGE`, `STOCKHISTORY`, `GROUPBY`, `PIVOTBY` and `PERCENTOF` and loses `QUERYSTRING` (not an Excel function).
  The array formulas of one sheet may cover 1,000,000 cells on load (`maxArrayCells` option of `importFromXlsx`; past it a formula is kept as a
  saved value and reported as `xlsx.package.sizeLimit`), and `exportToXlsx(workbook, { maxArrayCells, onWarning })` writes at most 1,000,000 cells
  of one spilled result (the first rows, the range cut to match, reported to `onWarning`): a file of 1.7 KB with `ref="A1:J400000"` took 44 s and 3 GB to write back.
  A formula whose text result is an error literal (`="#N/A"`) is written as an error value: the model cannot tell them apart.
  The array budget of an import is settable where a host opens files: `maxArrayCells` in `SheetsFileImportOptions` (`importWorkbookFile`) and a
  `maxArrayCells` prop of `WorkbookCanvas`; the `xlsx.package.sizeLimit` effect names them.
  An imported legacy (Ctrl+Shift+Enter) array formula is no longer written back as a dynamic array when the engine spills it: it stays `{=...}` in
  Excel (no `cm`, no metadata part); only a dynamic array and a formula typed in the editor that spills are written as dynamic arrays.
  `exportToXlsx` also has a budget for the whole workbook, `maxTotalArrayCells` (default 1,000,000 cells of array formula results over all formulas
  and sheets), and reports the first 20 cuts one by one and the others as a count. The in-app download hands them to the host: the
  `onExportWarning` prop of `WorkbookCanvas` and the `onWarning` option (`SheetsFileExportOptions`, fourth argument) of `exportWorkbookFile`.
- **PPTX writer A: package, theme, master and layouts (K12a).** New `exportPptx(presentation, options?)` in `@b32nio/spindle-io/pptx`
  (`jszip` and string builders, no new dependency) writes what `importPptx` returns as a .pptx: `[Content_Types].xml`, relationships,
  `presentation.xml` (slide size, slide list with kept or allocated ids, hidden slides, sections), the deck theme (twelve colours,
  major and minor fonts with east Asian and complex-script faces), one master with the standard placeholders and the deck's colour map,
  the deck's layouts with their placeholders stated in full, empty slide shells and `docProps` (the application is "Spindle"; no clock
  unless `options.metadata.modified` is given). The same deck gives the same bytes. It returns `{ bytes, warnings }`: the warnings say
  what is not exported yet (slide shapes and text, K12b; pictures, tables, groups, connectors, notes and comments, K13; opaque parts,
  transitions and timing, K14), so nothing is dropped silently. New `lintPptxPackage(bytes)`: the structural lints of the plan (content
  types, relationship ids and targets, unique shape ids, `p:spTree` child order, slide, master and layout id ranges). The PPTX fidelity
  harness now runs import, export, import for every readable deck (`roundtrip` goldens hold real differences; the scoreboard row L14 is measured),
  and `scripts/fidelity-core/lo-check.mjs` is a local LibreOffice open-and-render check (not run in CI).
- **PPTX writer A: slides with text boxes and shapes (K12b).** `exportPptx` now writes each slide's `p:sp` shapes and text boxes: preset geometry
  by its name with its adjustments (`unknown:<prst>` written as the preset name), placeholders, solid, gradient and pattern fills,
  outlines (width, dash, cap, join, arrow heads), text bodies with the body properties, paragraph levels and spacing, bullets, tab stops,
  run marks, fields and line breaks, hyperlinks (shape and run, `http`, `https`, `mailto` and `tel` only, plus the slide-jump and file actions without a target; slide jumps point at the slide's new
  number), slide backgrounds and the layouts' static shapes. `pptxId`s are kept when usable and unique, else allocated. The level style
  every paragraph of a level shares goes into the shape's list style, so PowerPoint (which ignores run defaults on a paragraph) shows the same text.
  Custom geometry, picture fills and custom dash patterns are not written yet and the warnings say so. The PPTX round-trip report now aligns
  elements by kind and name and allows half a pixel on frames; decks of shapes and text read back with no differences, and the
  round-trip differences over the corpus fall from 846 to 140, all of them content still to come (K13, K14) or macro actions dropped on purpose.
- **PPTX writer B: pictures, tables, groups, connectors, notes, comments, custom geometry (K13).** `exportPptx` now writes `p:pic` pictures
  (media parts with their content types, shared between pictures with the same image and keeping their part name when that is a usable
  `ppt/media` name; crop, transparency, outline, geometry and link), tables (`a:tbl` with column and row sizes, spans, per-cell borders,
  margins, anchors, fills and the style flags, the style named in `tableStyles.xml`), groups (the group's box is its child space, so
  each child's slide-space frame is mapped back through the group's rotation and flips, nested groups included), connectors with
  `a:stCxn` and `a:endCxn` glued to the shapes' new ids, custom geometry paths, speaker notes (a notes master and a notes slide per slide that has notes)
  and comments (legacy format with authors and replies; modern comments are written as legacy ones and the warnings say so). Pictures with no
  embedded image PowerPoint can show (linked, missing, EMF, WMF, WebP, AVIF, an SVG without a PNG) are left out and reported. Opaque parts, transitions and
  timing stay K14. The PPTX round-trip differences over the corpus fall further: pictures, tables, groups, connectors, notes, comments and custom
  geometry now read back with no differences; what is left is opaque content (K14), picture fills, dropped pictures and the modern-comment conversion.
  The outer shadow of a shape is now a model property (`PptxShape.shadow`, new additive type `PptxShadow`: colour with alpha, blur, distance,
  direction, anchor, rotation flag): the reader reads an effect list that holds one `a:outerShdw` and the writer writes it back. `pptx.shapes.effects` no longer
  counts a plain shape whose effect list is only that shadow; glow, reflection, soft edges, inner and preset shadows, two shadows, and
  shadows with a scale or skew, the effects of groups, pictures, connectors and runs are still reported.
- **PPTX writer C: opaque content, transitions and timing (K14).** `exportPptx` now re-emits what the reader keeps in the opaque part store:
  charts with their embedded workbook and style and colour parts, SmartArt (data, layout, quick style, colours and the cached drawing, which keeps
  the relationship id the data part names), OLE objects with their preview picture, audio and video with their poster frame (under both the
  clip's and the `p14:media` relationships) and ink, each as the frame that names its parts and each part byte for byte (hash equal). A part keeps
  its name unless the package has another use for it (then it moves under `ppt/opaque/` or gets a number) and the relationship parts are
  rewritten only where a target changed or points at a part that was not kept. Transitions (with their `mc:AlternateContent` fallback) and the
  timing tree are written as read, after `p:clrMapOvr`; a timing that names a shape the slide does not hold under the same id, a malformed
  fragment and a sound the file does not carry are left out and reported. Linked media and objects, ActiveX controls and extended
  charts (`chartEx`) are not written and are reported (no link to linked media is written on the slide; the external relationships inside a
  kept part's own relationship part, such as a chart's linked workbook, are kept as they were, inert). The lints of the writer now also report a part no relationship targets, and accept the id 0
  of an OLE preview picture and an empty relationship id (`r:blip=""`). The PPTX round-trip differences for opaque content, transitions and
  animations fall to zero over the corpus. No collab schema change (no new model fields).

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
