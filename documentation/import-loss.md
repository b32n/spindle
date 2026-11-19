# Import loss report

Spindle models a subset of Excel and Word. Instead of dropping the rest silently, every importer can tell you **what it cannot preserve**
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
from `@b32nio/spindle-io/{xlsx,docx}` (XLSX/DOCX) and `@b32nio/spindle-shared/export/csv`.

`ImportLossReport` is `{ format, items: ImportLossItem[] }`; an item is `{ id, label, count, where?, effect }` with a stable `id` such as
`xlsx.conditionalFormatting`. An empty `items` list means nothing detectable is lost. The scan is a cheap pre-pass over the zip (no workbook
model is built) and never throws on a malformed file; the import itself reports that error.

## What is reported

| Format | Reported |
|---|---|
| XLSX | conditional formatting, data validation, filters, protected sheets / workbook protection, grouped rows/columns, sparklines, hidden sheets (become visible), named ranges, charts, images, comments/notes, tables, pivot tables, macros, external links, slicers, rich text inside cells |
| DOCX | tracked insertions (text kept as if accepted, author/date lost) and deletions (applied), content controls (text kept, control lost), footnotes, endnotes, floating images, text boxes/shapes, equations, embedded objects, charts/diagrams, multi-column sections, tables of contents and indexes (text kept, field lost), legacy form fields |
| CSV | semicolon- or tab-separated files (they import into one column) |

`inspectDocx(bytes)` / `detectDocxLoss(bytes)` return the table above (the report the confirmation dialog uses). Pass `{ detail: 'full' }` for the
complete measurement, which adds ids for things the importer flattens or ignores rather than drops: `docx.sections` (sections after the first are kept
but not shown or edited), `docx.tableFormatting`, `docx.tableStructure`, `docx.listFormatting` (other number formats), `docx.headerFooterVariants` (first-page
and even/odd headers) and `docx.sectionProperties` (page-number restarts and formats, vertical alignment), and a `where` breakdown by field kind on
`docx.fields`. The full report also lists a long tail: `docx.trackedMoves`, `docx.formattingRevisions` (`rPrChange`, `pPrChange`, `tblPrChange`,
`sectPrChange`), `docx.dropCaps`, `docx.watermarks`, `docx.pageBorders`, `docx.lineNumbers`, `docx.hyphenation`, `docx.tabStops` (custom stops and
leaders), `docx.rtl`, `docx.ruby` (base text kept, reading marks lost), `docx.embeddedFonts`, `docx.altChunk` and `docx.legacyForms` (`FORMTEXT`/`FORMCHECKBOX`/`FORMDROPDOWN`,
which `docx.fields` also counts). In both modes `docx.charts` counts chart and SmartArt objects referenced from the body (not the XML parts behind them) and
`docx.floatingImages` counts anchored pictures only; anchored shapes and text boxes are not floating images. Tabs, line breaks, soft and non-breaking hyphens, and the text inside smart tags, custom XML, content controls, tracked insertions and ruby base text are imported, so they are no longer reported as lost. The DOCX fidelity harness (`fixtures/docx/`) records the full report.

Paragraph and run properties the importer keeps (line spacing, left/right/first-line/hanging indents, keep-with-next, keep-lines-together, page-break-before, all caps, hidden text, East Asian and complex-script fonts, list start numbers and restarts) are not reported as lost. The editor does not lay out keep-with-next, keep-lines-together, page-break-before or at-least/exact line heights yet, and shows hidden text dimmed instead of removing it; all of them are kept in the document and written back on export.

Named styles are kept: the names, ids and own properties of the paragraph and character styles the document uses (with the styles they are based on, linked to or followed by) are stored in `DocumentData.namedStyles`, blocks and runs keep a `namedStyleId`, and the export writes them back into `styles.xml` with `pStyle`/`rStyle` references, so `docx.namedStyles` is no longer reported. The editor does not show a style gallery yet; the formatting of a styled paragraph or run is also applied directly to it, as before. Table styles and the style definitions of styles nothing uses are not kept (table formatting is `docx.tableFormatting`).

Fields are kept with their cached result: `REF`, `PAGEREF`, `SEQ`, `STYLEREF`, `DATE`, `TIME`, `TITLE`, `AUTHOR`, `IF`, `MERGEFIELD` and the other kinds are imported as a generic field (instruction, kind and the result Word last computed, with its formatting), shown as that result and written back as a field so Word can refresh it. They are never recalculated in the editor, so a stale `DATE` or `PAGEREF` stays stale until Word updates it. `HYPERLINK` fields become links, and a field inside another field's result is part of that result. `docx.fields` now counts only what is not kept as a field: a table of contents or index keeps its text as static paragraphs (the field is lost until the table-of-contents work), and the legacy form fields (also under `docx.legacyForms`) lose their value. Two cases are still reported as warnings: fields in headers and footers other than `PAGE`/`NUMPAGES` (a single cached value would be wrong on most pages) and a field nested in another field's instruction (`IF { PAGE } > 1 ...`), which is kept as its cached value. A field whose result spans several paragraphs (other than a table of contents or index) keeps its text only.

Bookmarks and internal links are kept: a `w:bookmarkStart` becomes a named position in the text (including `_Toc` and `_Ref` bookmarks; Word's `_GoBack` cursor marker is dropped), a `w:hyperlink` with an anchor becomes a link to `#name` (`url#name` when it has a target as well), and clicking such a link in the editor scrolls to the bookmark. A bookmark that covers text is kept as a point at its start, so its range is not preserved: a `REF` field the application that opens the exported file recalculates would then show nothing, though its cached text is kept and `PAGEREF` is unaffected. A bookmark that starts between paragraphs moves to the start of the next paragraph, and bookmarks in headers and footers are dropped. A link around a field (a table-of-contents page number) keeps the field but not the link. There is no `docx.bookmarks` id any more.

Comments are imported (threads, replies, resolved state and anchors), so `docx.comments` is reported only when the comments part itself is damaged and its comments cannot be read.

Parse-time problems found while importing (an unresolvable image, an unsupported field…) still arrive as `warnings`.

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
