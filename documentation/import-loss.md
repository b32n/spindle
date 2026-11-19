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
| DOCX | tracked insertions (text dropped) and deletions (applied), content controls, footnotes, endnotes, comments, floating images, text boxes/shapes, equations, embedded objects, charts/diagrams, multi-column sections, fields other than PAGE/NUMPAGES |
| CSV | semicolon- or tab-separated files (they import into one column) |

`inspectDocx(bytes)` / `detectDocxLoss(bytes)` return the table above (the report the confirmation dialog uses). Pass `{ detail: 'full' }` for the
complete measurement, which adds ids for things the importer flattens or ignores rather than drops: `docx.sections` (sections after the first are kept
but not shown or edited), `docx.namedStyles`, `docx.tableFormatting`, `docx.tableStructure`, `docx.paragraphFormatting` (line spacing, right/hanging
indents, keep-with-next/lines, page-break-before, hidden text), `docx.listFormatting` (other number formats, custom starts), `docx.tabsAndBreaks`
(tabs, line breaks and hyphen characters dropped from text), `docx.bookmarks` (bookmarks and internal links), `docx.headerFooterVariants` (first-page
and even/odd headers) and `docx.sectionProperties` (page-number restarts and formats, vertical alignment), and a `where` breakdown by field kind on
`docx.fields`. The full report also lists a long tail: `docx.trackedMoves`, `docx.formattingRevisions` (`rPrChange`, `pPrChange`, `tblPrChange`,
`sectPrChange`), `docx.dropCaps`, `docx.watermarks`, `docx.pageBorders`, `docx.lineNumbers`, `docx.hyphenation`, `docx.tabStops` (custom stops and
leaders), `docx.rtl`, `docx.ruby`, `docx.embeddedFonts`, `docx.smartTags`, `docx.altChunk` and `docx.legacyForms` (`FORMTEXT`/`FORMCHECKBOX`/`FORMDROPDOWN`,
which `docx.fields` also counts). In both modes `docx.charts` counts chart and SmartArt objects referenced from the body (not the XML parts behind them) and
`docx.floatingImages` counts anchored pictures only; anchored shapes and text boxes are not floating images. The DOCX fidelity harness (`fixtures/docx/`) records the full report.

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
