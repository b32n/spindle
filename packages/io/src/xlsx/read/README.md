# Own XLSX package reader (S2a foundation, S2b sheet reader)

Internal API (not exported from the package entry points). `importXlsx` (`../reader.ts`) is built from it: the `exceljs`
reader (`../reader-exceljs.ts`) is unused and stays only as the oracle of `../reader-differential.test.ts` until S14. Built on `jszip` + `fast-xml-parser` through the X4 modules (`common/safe-zip`, `common/safe-xml`,
`common/import-error`), so every import limit is enforced before any part is read.

```ts
const pkg = await XlsxPackage.open(bytes, { limits });  // throws only ImportError (not-a-package, encrypted, size-limit, corrupt, legacy-binary)
const parts = await readWorkbookParts(pkg);   // { workbook, theme, styles, sharedStrings, colors, warnings }
// toImportExtras(parts): the same, as SpreadsheetImportResult's optional workbook/styles/sharedStrings/theme
const sheet = readSheet(await pkg.text(ref.part), ref, { styles, sharedStrings, colors, date1904, warn });  // one worksheet part
```

| Module | Reads | Notes |
|---|---|---|
| `package.ts` | zip, `[Content_Types].xml`, parts | case-insensitive part lookup; BOM and UTF-16 parts; a part that is absent or not well-formed is `undefined` plus a warning, never a throw |
| `content-types.ts` | defaults and overrides | override first, then extension default |
| `relationships.ts` | `.rels` | absolute (`/xl/...`, openpyxl), relative, `./`, `..`, backslash and percent-encoded targets; external targets kept as written; type matched by last segment, case-insensitively (LibreOffice writes `officedocument`); strict-namespace types; a fragment (`#a`) is dropped from an internal target, and a target with a URI scheme (`file:`) is kept as written and counts as external |
| `workbook.ts` | `workbook.xml` | workbook part found via relationships, else content types, else `xl/workbook.xml`; `sheets` (state, kind, part, relationships, comment parts), `definedNames`, `calcPr`, `workbookPr/date1904`, `bookViews`; shared parts located by relationship, content type or conventional name |
| `shared-strings.ts` | `sharedStrings.xml` | rich runs kept with their fonts, phonetic text left out, `_xHHHH_` escapes decoded; `forEachSharedString` pushes items instead of building a list |
| `styles.ts` | `styles.xml` | number formats, fonts, fills (pattern and gradient), borders with colours, `cellStyleXfs`, `cellXfs`, named styles, `dxfs`, custom `indexedColors`; tables keep file order so ids index straight in |
| `theme.ts`, `colors.ts` | theme part, colours | theme (Excel's index order), theme slots as `srgbClr`, `sysClr`, `prstClr`, `scrgbClr` or `hslClr` resolve (anything else warns and takes the Office default), indexed (workbook palette, per index falling back to the default; 64/65 system) and tint (HLS) resolved to `#rrggbb`; the file's own attributes are kept next to `resolved` |
| `font.ts` | `<font>` and run `<rPr>` | only what is written is set |
| `sheet.ts` | one worksheet part | one forward pass (`sax.ts`): `sheetData` cell by cell, the small elements around it (`sheetPr`, `dimension`, `sheetViews`, `sheetFormatPr`, `mergeCells`, `hyperlinks`) cut out by offset and read as little trees, `extLst` kept raw; `<col>` elements are read as events (at most 16,384) and applied by a last-writer sweep. Values by cell type, shared formulas expanded (`shared-formula.ts` moves the relative references), array formulas kept with their range, 1904 dates moved to 1900 serials, cells a merge covers dropped, links applied to the cells they name (both by one sweep over the cells in row order, `cell-ranges.ts`). A damaged part gives what precedes the damage (`details.wellFormed = false`) |
| `sax.ts` | XML text | forward-only events (start, end, text), local names, null-prototype attribute objects, entities never expanded, depth cap; never throws |
| `cell-style.ts` | a `cellXfs` entry | the editor's flat cell style and number format with colours resolved, computed once per format index |
| `cell-ref.ts`, `shared-formula.ts` | A1 references | parse and format; move a formula's relative references (quoted text, sheet names, calls, structured references and error literals are left alone) |

Output types are plain JSON-like data in `@b32nio/spindle-shared/export/xlsx` (`package-types.ts`), reached additively through the
optional `workbook` and `styles` fields of `SpreadsheetImportResult` (populated by `importXlsx`) and to the sheet fields (`details`, `styleIndex`, `sharedString`, `formulaInfo`).

Hostile names: anything keyed by file text (attribute names, shared formula ids, sheet names, relationship ids) is looked up in a
`Map`/`Set`, a null-prototype object or an array by number; tests use `__proto__`, `constructor` and `toString`.

**Streaming-ready (S20):** every part reader is a pure function from a parsed part to structured-cloneable data, with no DOM and
no shared state; the package class holds nothing but the zip and small relationship lists. Workbook-level parts use the bounded
element tree. `sheetData` is read by `readSheet` as a forward-only event stream (no tree) from `pkg.text()`; the part's text is
held as one string, so memory is at least the size of the part (a chunked feed for a worker is what S20 adds).

Limits on what a file can make the reader do (`ImportLimits`, tunable through the `limits` option of the io entry `importXlsx` and
`importFromXlsx`; the deprecated `importXlsx(bytes)` wrapper of `@b32nio/spindle-shared/export/xlsx` does not forward options):
a formula over 8192 characters (Excel's own limit) is dropped; the cells of a shared formula past a per-sheet budget
(`maxSheetFormulaChars`, 32 Mi) or the package's (`maxFormulaChars`, 64 Mi, shared by all sheets) keep their value and lose the
formula, with one warning per sheet. A formula costs its length plus 16 (its string's own weight), and an expanded formula is
one flat string, so the count is what the memory is. At most `maxTotalCells` (5,000,000) cells are kept over all sheets and
`maxSheets` (1,000) worksheets read, each with one warning when passed; a part is read once (a second `<sheet>` naming it is
refused with a warning); at most 16,384 `<col>` elements are read per sheet; a number text over 64 characters is not a
number; sheet names are made unique with a counter per name. Cell and row positions are clamped to the sheet limits; merges
and hyperlinks are applied by one sweep (`cell-ranges.ts`), near-linear in cells plus ranges. Every regular expression of
this directory is anchored, bounded or linear on hostile input (audited in S2b; `reader-hostile.test.ts` and the tests of `values.ts` and `cell-style.ts` time the ones that were not).

Tests: one file per module (malformed and producer-variety cases, an oracle table for tints), plus `index.test.ts`, which runs the
whole corpus (`fixtures/xlsx`: openpyxl, xlsxwriter, LibreOffice, hand-authored and edge files). `../reader-differential.test.ts` runs the own
reader and the `exceljs` reader over the corpus, over the files the e2e specs open and build, and over generated workbooks, and names every difference.
