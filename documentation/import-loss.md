# Import loss report

Spindle models a subset of Excel, Word and PowerPoint. Instead of dropping the rest silently, every importer can tell you **what it cannot preserve**
before the open document is replaced.

```ts
import { inspectWorkbookFile, importWorkbookFile, hasImportLoss, describeImportLoss } from '@b32nio/spindle-sheets-react';

const report = await inspectWorkbookFile(file);   // read-only: nothing is imported
if (hasImportLoss(report)) {
  console.log(describeImportLoss(report));
  // ['2 × conditional formatting rule(s) (Sales) — not imported', '1 × chart(s) — not imported']
}
const { warnings, loss } = await importWorkbookFile(workbook, file); // `loss` is the same report
```

Docs: `inspectDocumentFile(file)` / `importDocumentFile(doc, file)` from `@b32nio/spindle-docs-react`. Core (no React):
`inspectXlsx(bytes)`, `inspectCsv(text)` in sheets-core, `inspectDocx(bytes)` in docs-core, or `detectXlsxLoss` / `detectDocxLoss` / `detectCsvLoss`
from `@b32nio/spindle-io/{xlsx,docx}` (XLSX/DOCX) and `@b32nio/spindle-shared/export/csv`. PPTX: `detectPptxLoss(bytes, { detail })` from
`@b32nio/spindle-io/pptx` (slides has no importer yet, so there is no `inspectPptx` or dialog flow; see the PPTX table below).

`ImportLossReport` is `{ format, items: ImportLossItem[] }`; an item is `{ id, label, count, where?, effect, state? }` with a stable `id` such as
`xlsx.conditionalFormatting`. `state` (`'lost'` when absent, `'approximated'`, or `'kept'`: stored and saved back unchanged but not editable) lets the dialog list
"Kept, not editable" apart from "Will be lost"; ids stay stable and only counts, effect text and state change as constructs become preserved.
Parts an importer keeps opaquely live in the `@b32nio/spindle-shared/opaque` store (content-addressed, size-capped; a part over the cap is reported as `lost: too large`). The store returns copies from `get`/`list`, and `put` reports `contentTypeConflict` when equal bytes arrive under a different content type (the first type is kept; importers should surface it). Parts are hashed with SubtleCrypto (SHA-256), which browsers expose only in secure contexts (https or localhost): on a plain-http page `put`/`hashBytes` reject with `SubtleCrypto unavailable`, so hosts must serve over https. An empty `items` list means nothing detectable is lost. The scan is a cheap pre-pass over the zip (no workbook
model is built) and never throws on a malformed file; the import itself reports that error. The XLSX fidelity harness (`fixtures/xlsx/`,
`npm run sheets:xlsx-fidelity`) records this report for every corpus file next to a package ledger that accounts for each part of the file.

## What is reported

| Format | Reported |
|---|---|
| XLSX | conditional formatting, data validation, filters, protected sheets / workbook protection, grouped rows/columns, sparklines, hidden sheets (become visible), named ranges, charts, images, comments/notes, tables, pivot tables, macros, external links, slicers, rich text inside cells, formulas that use functions the editor does not have (`xlsx.formulas.unsupported`), formulas with table (structured) references (`xlsx.formulas.structuredRefs`), number formats the editor cannot show like Excel: kept as written and saved unchanged, but shown in General or with English month and day names (`xlsx.numberFormats.approximated`) |
| DOCX | tracked changes that cannot be kept (formatting changes of a paragraph mark's own formatting or of a table or section, a change inside a field result, around an image or a note reference; insertions, deletions, moves and formatting changes of text, paragraph marks and table rows are kept), content controls that cannot be kept (repeating sections, data-bound, group, row- or cell-level controls: their content is kept, the control lost), footnotes and endnotes that nothing in the body refers to, floating images, text boxes/shapes, equations, embedded objects, charts/diagrams, multi-column sections (kept for export, laid out as one column), tables of contents that cannot be read as a block and indexes (text kept, field lost), legacy form fields (text kept as plain text) |
| PPTX | **nothing imports yet**: there is no PPTX importer, so no `.pptx` can be opened today and every id below is `lost: not imported`. Detected: charts, SmartArt, audio and video, embedded objects, transitions, animations, comments, macros, hidden slides, speaker notes, fonts the editor may not have, shapes with a geometry outside its library of 76 presets or custom geometry, EMF/WMF pictures, encrypted and legacy `.ppt` files, size limits (detector only; the full level adds 43 finer ids) |
| CSV | semicolon- or tab-separated files (they import into one column) |

### XLSX loss ids

`inspectXlsx(bytes, { detail })` (sheets-core) and `detectXlsxLoss(bytes, { detail, knownFunctions, isNumberFormatExact })` (`@b32nio/spindle-io/xlsx`)
parse the package XML (entities are never expanded, oversized parts are skipped and named, nothing throws). The default `detail: 'summary'` is what the
confirmation dialog shows; `'full'` adds the finer ids below and is what the XLSX fidelity harness records. Every `effect` starts with a state word:
`lost:` (not kept at all), `approximated:` (kept in a weaker form) or `kept:` (kept, listed for information). Formula counts are per formula cell (every cell of a shared formula counts with its master's functions and references). Counts are of objects, not of package
parts: a chart is one chart however many parts it has, a picture is one picture however many times its image file is used.

Summary ids (dialog): `xlsx.conditionalFormatting`, `.dataValidation`, `.autoFilter`, `.sheetProtection`, `.workbookProtection`, `.outline`, `.sparklines`,
`.hiddenSheets`, `.definedNames`, `.charts`, `.images` (pictures, sheet backgrounds and pictures in cells), `.comments` (sheets with notes or threads),
`.tables`, `.pivotTables`, `.macros`, `.externalLinks`, `.slicers`, `.richText`, and three new ones: `xlsx.formulas.unsupported` (`where` = function names;
a function counts as supported when the editor's formula catalog has it after the `_xlfn.`/`_xlws.` prefix is removed; only reported when the caller passes
`knownFunctions`, as `inspectXlsx` does), `xlsx.formulas.structuredRefs` and `xlsx.numberFormats.approximated` (`where` = the codes; only reported when the caller passes `isNumberFormatExact`, as
`inspectXlsx` does, which since plan S3 means: the number-format engine reads the code and shows it as Excel does. Every code is kept verbatim, as a structured
format when that writes the very same code back and as `custom` otherwise, so a save never changes it; what is still reported is a code the engine cannot read
(not valid in Excel, a non-Gregorian calendar, `[DBNum1]`), a date or time tagged with another language (`[$-407]dddd`: month and day names would be English),
and a built-in id with no code, such as the era dates 27 to 36 of the Asian locales. Before S3 the check was "maps to a structured format", which flagged
every date, time, fraction, elapsed, scaling and literal code: 66 occurrences in 3 of the 43 corpus files, now 0). The effect text says so: the code is kept as written and saved unchanged, but shown in General or with English month and day names.
`xlsx.formulas.cachedOnly` is not reported yet: it needs the cached-value fallback of the formula work (plan S4).
`xlsx.formulas.wholeColumn` (whole-column and whole-row references, `A:A`, `1:1`) is no longer reported: the engine evaluates them (plan S5), so they are not a
loss. A formula using one still appears under `xlsx.formulas.unsupported` if it calls a function the engine lacks.

Full-only ids:

| Group | Ids (what is counted) |
|---|---|
| Formulas | `xlsx.formulas.array` (`t="array"`), `.dynamicArray` (array formula with `cm` metadata), `.externalRef` (`[1]Sheet!A1`), `.volatile` (`NOW`, `TODAY`, `RAND*`, `OFFSET`, `INDIRECT`, `INFO`) |
| Styles | `xlsx.styles.themeColors` (cell formats in use with a theme or indexed colour; plain black text excluded), `.patternFill`, `.gradientFill`, `.diagonalBorder`, `.namedStyles` (named styles in use), `.columnRowStyles` (column and row default formats) |
| Sheet | `xlsx.sheet.tabColor`, `.zoom` (other than 100), `.rtl`, `.gridlines` (turned off), `.veryHidden`, `.customViews` |
| Print | `xlsx.print.setup` (a sheet whose orientation, paper size, scale, fit to page, print options or margins differ from Excel's defaults), `.areas`, `.titles`, `.breaks`, `.headersFooters` |
| Conditional formatting | `xlsx.cf.dataBarExt`, `.iconSetExt` (Excel 2010 extensions), `.formulaRule`, `.stopIfTrue` |
| Validation | `xlsx.dv.customFormula`, `.messages`, `.crossSheet` (references to another sheet, including Excel 2010 `x14` validations) |
| Comments, links | `xlsx.comments.threaded`, `.richFormatting` (notes with bold, italic, underline, strike or script runs, or runs of differing font, size or colour; the bold author line and plain run properties every note has are not counted); `xlsx.links.internal` (to a cell or name), `.unsafeScheme` (`javascript:`, `vbscript:`, `data:`, `file:` targets; state `kept:`, the target is stored as written and never opened or run by the library) |
| Objects | `xlsx.objects.shapes` (shapes and connectors), `.textBoxes`, `.formControls` (form and ActiveX controls), `.embeddedObjects` (OLE), `.cellImages` |
| Workbook | `xlsx.workbook.date1904` (dates are converted to the 1900 system), `.calcSettings` (manual calculation, iteration, precision as displayed, R1C1; the `calcPr` Excel writes in every file is not reported), `.externalData` (connections, query tables, data model), `.richData` (linked data types) |
| Package | `xlsx.package.unaccountedParts` (parts nothing reads or reports, named in `where`), `.encrypted` and `.legacyBinary` (an OLE container: password-protected package, or a legacy `.xls`; also a binary `.xlsb`), `.sizeLimit` (parts over the size limit, 64 MiB each or 256 MiB in total, and cell data past 48 MiB of one sheet, are not scanned; a file over the importers' 100 MB package limit is not scanned at all) |

The XLSX loss counts grow with each fidelity change that keeps a construct: a feature that becomes preserved leaves the table or changes state. The package ledger
in `fixtures/xlsx/` maps every part of a file to a mapped, dropped or unaccounted status using these ids.

`inspectDocx(bytes)` / `detectDocxLoss(bytes)` return the table above (the report the confirmation dialog uses). Pass `{ detail: 'full' }` for the
complete measurement, which adds ids for things the importer flattens or ignores rather than drops: `docx.listFormatting` (other number formats) and `docx.sectionProperties` (chapter numbers in page numbers), and a `where` breakdown by field kind on
`docx.fields`. The full report also lists a long tail: `docx.formattingRevisions` (`rPrChange` of a paragraph mark, `tblPrChange`, `trPrChange`, `tcPrChange`,
`sectPrChange`), `docx.dropCaps`, `docx.watermarks`, `docx.pageBorders`, `docx.lineNumbers`, `docx.hyphenation`, `docx.tabStops` (custom stops and
leaders), `docx.rtl`, `docx.ruby` (base text kept, reading marks lost), `docx.embeddedFonts`, `docx.altChunk` and `docx.legacyForms` (`FORMTEXT`/`FORMCHECKBOX`/`FORMDROPDOWN`, whose displayed text is kept as plain text; they are no longer also counted in `docx.fields`). In both modes `docx.charts` counts chart and SmartArt objects referenced from the body (not the XML parts behind them) and
`docx.floatingImages` counts anchored pictures only; anchored shapes and text boxes are not floating images. Tabs, line breaks, soft and non-breaking hyphens, and the text inside smart tags, custom XML and ruby base text are imported, so they are no longer reported as lost. The DOCX fidelity harness (`fixtures/docx/`) records the full report.

Tracked changes: `w:ins` and `w:del` around runs and links are kept as the `insertion` and `deletion` marks (author, date, id), with the deleted text kept (`w:delText`), shown struck through, and written back as `w:ins` / `w:del`. They are not applied on import, so a document with deletions has more text than it did before. A change around a field is not kept as a change (an inserted field stays as it is, a deleted field is dropped), a deleted image or note reference is dropped and an inserted one stays as ordinary content; each adds a warning and is counted as `docx.trackedChangesNotKept` (changes around a field, image or note reference, and deleted text inside a field result). Headers and footers keep no tracked changes (deleted text is dropped, inserted text kept as ordinary text: `docx.trackedHeaderFooterChanges`, which also counts a formatting change there), and a change around whole paragraphs or tables is counted as `docx.trackedBlockChanges` (its content is not imported). Moves are kept as a deletion and an insertion that share the name of the move, formatting changes of runs and paragraphs (`w:rPrChange`, `w:pPrChange`) with the formatting before, and inserted or deleted paragraph marks and table rows as attributes of the paragraph or row. What stays in the loss report: `docx.formattingRevisions` (the change of a paragraph mark's own formatting, a paragraph change of numbering, tab stops, borders or shading, whose old value the change cannot carry, and changes of tables, table rows, cells and sections: `tblPrChange`, `trPrChange`, `tcPrChange`, `sectPrChange`), `docx.trackedBlockChanges`, `docx.trackedStructureChanges` (a tracked change of the paragraph mark of a paragraph that is read as a page break or a horizontal rule, which have no place for it; an image paragraph keeps its change as a paragraph with an inline image) and `docx.trackedHeaderFooterChanges`; a change in a field's result is not kept, and a formatting change on a run with no text of its own (a field result, an image, a note reference) is not kept either; each adds a warning instead of a count. A table cell of one paragraph keeps that paragraph's block (not just its text) when the paragraph has a tracked change of its mark, so the change is read and written back like any other. In the file, the change of a paragraph mark and the changes of its text can share one `w:id`; they are read as separate changes (one per kind) and written with ids of their own. A move of a paragraph mark is written back as an insertion or deletion.

Paragraph and run properties the importer keeps (line spacing, left/right/first-line/hanging indents, keep-with-next, keep-lines-together, page-break-before, all caps, hidden text, East Asian and complex-script fonts, list start numbers and restarts) are not reported as lost. The editor does not lay out keep-with-next, keep-lines-together, page-break-before or at-least/exact line heights yet, and shows hidden text dimmed instead of removing it; all of them are kept in the document and written back on export.

Named styles are kept: the names, ids and own properties of the paragraph and character styles the document uses (with the styles they are based on, linked to or followed by) are stored in `DocumentData.namedStyles`, blocks and runs keep a `namedStyleId`, and the export writes them back into `styles.xml` with `pStyle`/`rStyle` references, so `docx.namedStyles` is no longer reported. The editor does not show a style gallery yet; the formatting of a styled paragraph or run is also applied directly to it, as before. Table styles are kept as well (see below). The style definitions of styles nothing uses are not kept.

Fields are kept with their cached result: `REF`, `PAGEREF`, `SEQ`, `STYLEREF`, `DATE`, `TIME`, `TITLE`, `AUTHOR`, `IF`, `MERGEFIELD` and the other kinds are imported as a generic field (instruction, kind and the result Word last computed, with its formatting), shown as that result and written back as a field so Word can refresh it. They are never recalculated in the editor, so a stale `DATE` or `PAGEREF` stays stale until Word updates it. `HYPERLINK` fields become links, and a field inside another field's result is part of that result. `docx.fields` now counts only what is not kept as a field: an index (and a table of contents that cannot be read as a block) keeps its text as static paragraphs. Two cases are still reported as warnings: fields in headers and footers other than `PAGE`/`NUMPAGES` (a single cached value would be wrong on most pages) and a field nested in another field's instruction (`IF { PAGE } > 1 ...`), which is kept as its cached value. A field whose result spans several paragraphs (other than a table of contents or index) keeps its text only.

Table formatting is kept: borders (line style, width and colour, per table and per cell), cell shading, vertical alignment, repeating header rows, `cantSplit`, column widths and the table style (with its `tblLook`) are imported and written back, so `docx.tableFormatting` is no longer reported. A table style's conditional formatting (header and total rows, first and last column, banding, corner cells) is applied to the cells it reaches, the way Word resolves it: shading, borders and vertical alignment become the cell's own, text properties (bold, colour...) become the formatting of its runs, and the table style itself is kept in `namedStyles` and written to `styles.xml`. Known limits: a table style's paragraph properties (spacing, alignment) and cell margins are not applied; the table's overall width, indentation and alignment are not kept (columns use the grid widths); the editor does not split tables across pages, so repeating header rows and `cantSplit` are stored and exported but not acted on; other line styles (wavy, 3D, art borders) are approximated by single, double, dashed or dotted; pattern shadings (`w:val` other than `clear`/`solid`) keep only their fill colour. A table with no border lines is stored with empty `borders` and shows no lines.

Table structure is kept: a cell's paragraphs, lists, headings, images and nested tables are imported (`DocTableCell.blocks`) and written back, so `docx.tableStructure` is no longer reported. A vertical merge ends in a row that does not continue it, a continuation cell with no merge above it stays an ordinary cell, and a row that starts after skipped grid columns (`w:gridBefore`) gets an empty cell without lines in front. A cell's paragraph formatting is kept when the cell has several blocks; a cell of one paragraph keeps its text and run formatting only (and its block when the paragraph has a tracked change). Known limits: a table nested more than 20 deep is dropped with a warning; a rotated cell (`w:textDirection`) is kept and exported but shown with horizontal text; the editor draws every paragraph of a cell without space between paragraphs; the table style's paragraph properties are not applied (see table formatting above).

Headers, footers and section properties are kept in the model and written back, so `docx.headerFooterVariants` is no longer reported and `docx.sectionProperties` only counts chapter numbers in page numbers: each section's default, first-page (`w:titlePg`) and even-page header and footer (by `w:type`, with the document's even/odd setting) are imported, as are the page number restart and format (`w:pgNumType`), vertical alignment (`w:vAlign`) and text columns (`w:cols`: count, spacing, separator, unequal widths). A first-page or even-page variant that is switched on without a part of its own stays empty and is not written; a section with no header or footer reference has none (Word would repeat the previous section's, and does so again when the file is written back). The editor draws the first-page header and footer only: even-page headers and footers, page number restarts and formats, vertical alignment and columns are stored and exported but not shown (`docx.columns` keeps counting multi-column sections: they are laid out as one column). Header and footer fields other than page numbers are still dropped with a warning, and tables in headers and footers are not imported.

Bookmarks and internal links are kept: a `w:bookmarkStart` becomes a named position in the text (including `_Toc` and `_Ref` bookmarks; Word's `_GoBack` cursor marker is dropped), a `w:hyperlink` with an anchor becomes a link to `#name` (`url#name` when it has a target as well), and clicking such a link in the editor scrolls to the bookmark. A bookmark that covers text is kept as a point at its start, so its range is not preserved: a `REF` field the application that opens the exported file recalculates would then show nothing, though its cached text is kept and `PAGEREF` is unaffected. A bookmark that starts between paragraphs moves to the start of the next paragraph, and bookmarks in headers and footers are dropped. A link around a field (a table-of-contents page number) keeps the field but not the link. There is no `docx.bookmarks` id any more.

### PPTX loss ids

`detectPptxLoss(bytes, { detail })` (`@b32nio/spindle-io/pptx`) parses the package XML with the same safety measures as the XLSX detector (entities are
never expanded and a DOCTYPE is cut out, nesting is capped at 100 levels, a part is read only up to 64 MB and the whole read up to 256 MB, counted on
the bytes that actually come out of the zip so a forged header cannot make a bomb inflate, a package of more than 50,000 entries is not scanned, nothing
throws). **There is no PPTX importer yet** (plan K9 to K15): today no presentation can be opened at all, so every effect says
`lost: not imported (no PPTX importer exists yet, so no .pptx can be opened today)` (a few say more: macros are never run or kept, an encrypted package
cannot be opened). The ids are stable; as importer PRs land, only counts, effect texts and states change. The default `detail: 'summary'` returns the
16 summary ids; `'full'` adds the finer ids and is what `npm run slides:pptx-fidelity` records. `where` lists slides as `Slide n` (their place in the
presentation; a slide part the presentation does not list is named by its file), or the fonts, geometries, languages, link schemes or action kinds
named in the id's row below, capped at five entries.

Counts are of objects, not of package parts, and are taken from the **slides** (nothing on a layout or master is counted; those carry only the
masters, theme, colour map and font ids): a chart is one chart however many parts it has, an `mc:Choice`/`mc:Fallback` pair is read once (the
fallback is skipped), a slide with a transition counts once however the transition is written.

Summary ids (dialog):

| Id | What is counted |
|---|---|
| `pptx.charts` | chart frames (`c:chart`, and the Office 2016 `cx` chart form) on slides; if no slide shows one, the chart parts |
| `pptx.smartArt` | SmartArt diagram frames on slides; if none, the diagram data parts |
| `pptx.media` | audio and video pictures (`a:audioFile`, `a:videoFile`, `a:quickTimeFile`) |
| `pptx.embeddedObjects` | OLE objects (`p:oleObj`) and ActiveX controls (`p:control`) |
| `pptx.transitions` | slides with a transition |
| `pptx.animations` | animation effects (`p:cTn` with a `presetClass`); a slide whose timing tree has no preset effect but animates something (`p:set`, `p:anim`, `p:animEffect`, `p:animMotion`, `p:animClr`, `p:animScale`, `p:animRot`) counts one; a tree with only media playback nodes is not an animation |
| `pptx.comments` | comment threads, legacy (`p:cm`) and modern (`p188:cm`; replies are not counted: modern `p188:reply` is not a `cm`, and a legacy reply, a `cm` with a `p15:parentCm`, is skipped) |
| `pptx.macros` | the macro project (`ppt/vbaProject.bin`), at most one |
| `pptx.hiddenSlides` | slides with `show="0"` |
| `pptx.notes` | slides whose notes page has text (the slide number, date, header, footer and slide-image placeholders do not count) |
| `pptx.fonts.unavailable` | distinct font families (case-insensitive) used by runs, by the layouts and masters, or by the theme of a master, outside a short list (below); `where` = the names. Theme references (`+mj-lt`, `+mn-ea`) never count |
| `pptx.shapes.unsupportedGeometry` | shapes with custom geometry, or with a preset outside the library of slides-core (`ShapePreset`: the original 18 plus rectangle variants, block arrows, stars, callouts, flowchart shapes, brackets and braces, 76 in all, named by their OOXML `prst`) and the two line geometries (`line`, `straightConnector1`); `where` = the preset names (`custGeom` for custom geometry). A preset outside the library is kept as an `unknown:<prst>` shape with its adjustments, drawn as a rounded rectangle and re-exported verbatim. Connectors with an elbow or curved geometry are not counted here (see `pptx.shapes.connectors.elbowCurved`); a picture's crop-to-shape preset is not counted, but a picture with custom geometry is |
| `pptx.images.unsupportedFormat` | pictures whose image file is EMF, WMF, EMZ or WMZ (browsers cannot show them) |
| `pptx.encrypted`, `pptx.legacyBinary` | an OLE container with an `EncryptedPackage`/`EncryptionInfo` stream (password-protected), or any other OLE container (a legacy `.ppt`); these two end the scan |
| `pptx.sizeLimit` | parts not read because of the size limits (named), or one item for a package with too many entries |

`pptx.fonts.unavailable` means "the deck may not look the same here", not that anything is lost. A font is **available** when it is one of: Arial,
Helvetica, Times New Roman, Times, Courier New, Courier, Georgia, Verdana, Tahoma, Trebuchet MS (web-safe), Calibri, Calibri Light, Cambria and their
metric-compatible substitutes Carlito and Caladea, Liberation Sans, Serif and Mono, Noto Sans and Serif, DejaVu Sans, or the generic families
`sans-serif`, `serif`, `monospace` and `system-ui`. The list is deliberately short and conservative; fonts outside it (including the East Asian fonts of
CJK decks) are reported. Bullet fonts (`a:buFont`) and symbol fonts (`a:sym`) are not scanned.

Full-only ids:

| Group | Ids (what is counted) |
|---|---|
| Masters and theme | `pptx.masters.multiple` (the number of slide masters, when there is more than one), `.handoutMaster`, `.notesMaster`, `.layoutInheritance` (slide placeholders with no frame of their own: position and size come from the layout and are flattened on import); `pptx.theme.colorMap` (a master colour map that is not the default for background and text, and slide or layout `overrideClrMapping`), `.fontScheme` (a theme used by a slide master whose major or minor Latin font is not Calibri or Calibri Light; `where` = the fonts), `.effectScheme` (a master's theme with an effect style that has a shadow, glow, reflection, soft edge or blur) |
| Text | `pptx.text.autofit` (`normAutofit` that applies a `fontScale` or `lnSpcReduction`, and `spAutoFit`; plain `normAutofit` and `noAutofit` do not count), `.bullets.custom` (paragraphs with a bullet colour or size, a bullet character other than `•`, or numbering with a start value or in a scheme other than `arabicPeriod`; the bullet font is not counted because PowerPoint writes Arial on every plain bullet), `.bullets.picture`, `.tabStops` (paragraphs with tab stops), `.columns` (text bodies with two or more columns), `.vertical` (text bodies and table cells with vertical text), `.fields` (slide number, date), `.wordArt` (text with a warp), `.rtl` (paragraphs), `.spacing` (runs with character spacing or small/all caps), `.baseline` (runs raised or lowered), `.lang` (runs tagged with a language other than en-US; `where` = the language tags) |
| Shapes | `pptx.shapes.custGeom`, `.effects` (shape, group or run effect lists with a shadow, glow, reflection, soft edge or blur, and effect DAGs), `.gradientFill` and `.patternFill` (on shapes, groups, backgrounds and table cells, not on outlines), `.connectors.elbowCurved` (connectors, and shapes, with a `bentConnector` or `curvedConnector` geometry), `.groups.nested` (groups inside a group; the inner ones count), `.scene3d` (shapes or text bodies with a 3-D scene or shape) |
| Images and tables | `pptx.images.crop` (a non-zero crop edge), `.recolor` (duotone, grayscale, black and white, brightness/contrast, colour change, tint, HSL), `.linked` (a picture linked rather than embedded), `.svg` (an SVG with a PNG fallback); `pptx.tables.merge` (the anchor cell of each merged region), `.styles` (tables with a table style or header, total, first- or last-column formatting), `.banding` (banded rows or columns) |
| Presentation | `pptx.sections`, `.customShows`, `.slideSize` (a slide size that is not 16:9; `where` = the size in points), `.zoom` (slide and section zoom objects), `.ink` (ink content parts; if none, the ink parts), `.equations` (`a14:m`), `.signatures` (signature parts in `_xmlsignatures/`), `.externalLinks` (external relationships other than pictures: web pages, mail, files, linked media; counted per relationship; `where` = the URL schemes, never the URLs), `.actions` (action hyperlinks such as jump to slide or run macro; `where` = the kinds; a macro name is never recorded and macros are never run) |
| Package | `pptx.package.unaccountedParts` (parts that are neither plumbing, nor a part an importer will read, nor covered by an id above: custom XML, tags, embedded fonts, files in unexpected folders; only for a package that has `ppt/presentation.xml`) |

Slides can now hold gradient, picture and pattern fills (plan K1: shapes, text boxes, table cells and backgrounds; a pattern is kept in the model but painted as a solid blend of its two colours, so it is approximated, not exact). Nothing imports a `.pptx` yet, so `pptx.shapes.gradientFill` and `.patternFill` still report `lost: not imported`; their state changes when the reader (K9) maps these fills.

Pictures can now be cropped (plan K2: `ImageElement.crop`, the percent of the source trimmed from each edge, the same quantity as `a:srcRect`; the fit applies to the cropped region). Nothing imports a `.pptx` yet, so `pptx.images.crop` still reports `lost: not imported`; its state changes when the reader (K9) maps `a:srcRect`. Negative `srcRect` values (padding outside the picture) have no model equivalent and are clamped to zero.

Text bodies can now hold the body-level attributes of `a:bodyPr` (plan K3a: per-side insets, `anchorCtr`, `normAutofit` with its `fontScale` and `lnSpcReduction` applied as stored, `spAutoFit` as resize, `vert` and `eaVert`, `numCol`/`spcCol`). Nothing imports a `.pptx` yet, so `pptx.text.autofit`, `.vertical` and `.columns` still report `lost: not imported`; their state changes when the reader (K9) maps them. Known limits: PowerPoint's own line breaking differs from the browser's (fonts and metrics), so a stored scale can leave slightly more or less room than in PowerPoint; shrink-to-fit with several columns measures the balanced height; `resize` leaves vertical text and shapes alone (a shape keeps the imported `spAutoFit` but does not grow). Paragraph bullets, spacing forms, tab stops, baseline, caps, character spacing and fields (`pptx.text.bullets.*`, `.tabStops`, `.spacing`, `.baseline`, `.fields`) are not part of this step.

Paragraphs can now hold the attributes of `a:pPr` (plan K3b: `buNone`/`buChar`/`buAutoNum` with `buClr` and `buSzPct`, `startAt`, `lnSpc` as a multiple or exact size, `spcBef`/`spcAft` in points or percent, `marL`, `indent`, `tabLst`; `flattenLevelStyles` bakes a level's `lvlNpPr` defaults into the paragraphs). Nothing imports a `.pptx` yet, so `pptx.text.bullets.custom`, `.tabStops` and `.spacing` still report `lost: not imported`; their state changes when the reader (K9) maps them. Known limits: tab stops are kept and synced but the editor lays out default tabs; picture bullets (`buBlip`) and bullet fonts (`buFont`) are not modelled; the number scheme set is arabic, alpha and roman (other `buAutoNum` types fall back to arabic); a paragraph's bullet size and colour follow the first-run defaults only when set explicitly; the editor numbers by list run, not by indent level. 

Runs can now hold the attributes of `a:rPr` that K3c adds (`baseline`, `cap`, `spc`, `lang`) and `a:fld` runs of type `slidenum` and `datetime*` (plan K3c: marks `baseline`, `caps`, `charSpacing`, `lang` and the inline `field` node, whose cached text keeps its formatting). Nothing imports a `.pptx` yet, so `pptx.text.baseline`, `.spacing`, `.lang` and `.fields` still report `lost: not imported`; their state changes when the reader (K9) maps them. Known limits: a raised or lowered run is drawn at 65 percent of its size shifted by the stored percent (PowerPoint's exact size change differs by font); `cap="small"` uses the browser's small caps; the editor shows a field's stored text (a slide number reads `‹#›`) while the idle slide shows the live value; a date field shows the current date in the viewer's locale, not the format code of the file (`datetime1` to `datetime13` map to three layouts); the language tag is stored and exposed as `lang` on the run but does not drive spell checking; a tag with a one-character subtag (an extension such as `de-u-co-phonebk` or private use such as `en-x-foo`) is not stored.

Content controls are kept: a `w:sdt` of type rich text, plain text, drop-down list, combo box, date, check box, picture or building block is imported as a wrapper (`content-control` block, or `contentControl` inline content) around its content, with its title (`w:alias`), tag, lock, placeholder, list items, date format and value, and check box state, nested to any depth, and written back as a `w:sdt`. `docx.contentControls` now counts only what is still lost, with the reason in `where`: repeating sections and their items, data-bound controls (custom XML parts are not read), equation, citation, bibliography and group controls, controls around table rows or cells, and a control around a section break; all of these import as their content only. Known limits: the editor shows a title chip and a tooltip and toggles check boxes, but has no date, list or picture pickers, and the lock is kept and exported but not enforced; the placeholder text lives in the file's glossary part, which is not kept, so a control showing its placeholder keeps the text it displays; a control inside a field's result, in a header or footer, or around a section break is read as its content; custom check box glyphs (`w14:checkedState`) are written back as Word's default ones; `w:id`, `w:temporary` and other properties are not kept (ids are renumbered). Legacy form fields keep the text they display as plain text, and are still reported as `docx.legacyForms`.

Footnotes and endnotes are kept: each reference in the body becomes a note reference holding the text of its note (with its formatting, links, line breaks and fields), shown in the editor as a superscript number with each footnote at the bottom of its page (endnotes collected at the end of the section), and written back to `footnotes.xml` / `endnotes.xml` (with Word's separator notes and the settings that point at them), so `docx.footnotes` and `docx.endnotes` count only notes that nothing in the body refers to, notes whose only reference is in a text box, and notes in a part that cannot be read. Known limits: a note of several paragraphs becomes one paragraph with line breaks, and a table in a note is not kept (a warning); a picture in a note is kept but a list is just its text; a reference in a header, footer, field result or note is not kept (a warning); notes are numbered by the editor (footnotes 1, 2, 3, endnotes i, ii, iii, Word's defaults; the number format, restart and position settings of the file are not read) and the number Word shows after opening the written file is Word's own; notes can be inserted (Footnote menu) and edited in place in the editor; footnotes are laid out at the bottom of their page by the editor's own layout, which can break pages differently from Word (a page gives at most half its height to notes, and a note that does not fit continues on the next page); the note styles (`FootnoteText` and the reference styles) are not written, the formatting is on the runs.

A table of contents is kept as a `toc` block: a `TOC` field that starts in the first entry paragraph and ends after the last (in the body, or in a content control such as Word's Table of Contents building block) is read with its entries as they were saved (links to the `_Toc` bookmarks, tabs, `PAGEREF` fields with their cached numbers, entry styles), shown with a dotted leader and the number at the right, never recomputed on open, and written back as the field (begin, instruction and separator in the first entry, a right tab with a dotted leader at the margin on each entry, the end in a paragraph of its own) so Word can refresh it. An entry made in the editor is written with its number as a `PAGEREF` field. `docx.fields` now counts only a table of contents that stays static text (its field ends elsewhere than the body or control it starts in, has a table or section break inside, or is a simple field `w:fldSimple`) and indexes. Known limits: the entries are plain paragraphs (a heading or list style on an entry is not kept as such), the title paragraph Word puts before the table ("Contents") is an ordinary paragraph, a table of contents in a table cell, a header or a footer is static text, and the tab stops of the entries are not kept (the editor draws a leader at the margin; the export writes one at the margin too).

Comments are imported (threads, replies, resolved state and anchors), so `docx.comments` is reported only when the comments part itself is damaged and its comments cannot be read.

Parse-time problems found while importing (an unresolvable image, an unsupported field…) still arrive as `warnings`.

## Files that cannot be opened at all

The importers do not report these as loss: they refuse the file, with an `ImportError` (`@b32nio/spindle-io`) whose `code` hosts can switch on and
whose `message` is fit to show a user. `importXlsx` and `importDocx` throw nothing else. A forged or corrupt zip entry (including the `mimetype` marker) is `corrupt`, never a plain `Error`.
A CSV over the size limit (`importFromCSV`, `importWorkbookFile`) is refused with an error of the same shape (`name` `ImportError`, `code` `size-limit`) that is
not an instance of the `@b32nio/spindle-io` class, because the CSV path is synchronous and does not load that package; `isImportError(e)`
(`@b32nio/spindle-sheets-core`) recognises both, so a host handles every refusal in one place. The loss ids above (`xlsx.package.encrypted`, `.legacyBinary`, `.sizeLimit`) are
what the `full` detail level reports for the same files, so a host that inspects first can say the same thing before trying.

| `code` | When |
|---|---|
| `encrypted` | an OLE compound file with an `EncryptedPackage` or `EncryptionInfo` stream (password to open), or a zip entry marked encrypted |
| `legacy-binary` | any other OLE compound file (`.xls`, `.ppt`, `.doc`), or a binary workbook (`.xlsb`) |
| `not-a-package` | empty, not a zip, an OpenDocument or Apple iWork package, or a zip without `[Content_Types].xml` |
| `size-limit` | over a limit below |
| `corrupt` | truncated or damaged zip, a part that cannot be inflated, a part name that is unsafe (`..`, absolute, backslash, drive letter, NUL), or a reader failure on a well-formed package |

Limits (options on the importer, `limits`; defaults `DEFAULT_IMPORT_LIMITS`): package 100 MB, 20,000 entries, 1 GB uncompressed in total, 500 MB per part
(256 MB for an XML part), compression ratio 200:1. The ratio is only judged for a part (or package) that inflates to more than 16 MB: smaller parts cannot cost more
than that however well they compress, and legitimate files do contain tiny, very regular parts. The limits are checked from the zip's declared sizes first and then
again while each part is inflated, so a header that lies about a size cannot get a bomb inflated.

XML is parsed with entities never expanded, the DOCTYPE removed and nesting capped at 256.

- **DOCX:** the five predefined entities are still decoded; a declared one (`&lol9;`, an external entity) stays as written text. `importDocx(bytes, { limits })` and `detectDocxLoss(bytes, { limits })` take the same options. The detector reads nothing from a package over the limits (an empty report; the import then says why). A `.docx` without `[Content_Types].xml` is `not-a-package`.
- **Known gap:** docs-core `importFromDocx` does not forward `limits` yet, so hosts using that entry point get the defaults only (follow-up in docs-core, tracked in `documentation/docs/TODO.md`).
- **CSV:** a leading byte order mark is dropped, text over 100 MB is refused (`code` `size-limit`).

## Confirmation dialog

The sheets and docs editors (when `fileActions` is on) check the file before replacing the open document. If the report is empty the file opens
straight away; otherwise an **Import anyway / Cancel** dialog lists what would be lost. Cancel (or Escape) leaves the document untouched and
`onFileOpened` is not called.

```tsx
<WorkbookCanvas fileActions importConfirmation="dialog" />        // default
<WorkbookCanvas fileActions importConfirmation="never" />         // import silently
<DocumentEditor fileActions
  importConfirmation={async (report, file) => myConfirm(file.name, describeImportLoss(report))} /> // your own UI
```

`ImportConfirmDialog` (props `{ fileName, report, onConfirm, onCancel }`) is exported by `@b32nio/spindle-shared/react` and re-exported from the
sheets, docs and slides React packages for hosts that build their own Open flow.
