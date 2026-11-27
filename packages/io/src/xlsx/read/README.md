# Own XLSX package reader: foundation layer (S2a)

Internal API (not exported from the package entry points, not wired into `importXlsx` yet: the `exceljs` path stays in
force until S2b). Built on `jszip` + `fast-xml-parser` through the X4 modules (`common/safe-zip`, `common/safe-xml`,
`common/import-error`), so every import limit is enforced before any part is read.

```ts
const pkg = await XlsxPackage.open(bytes, { limits });  // throws only ImportError (not-a-package, encrypted, size-limit, corrupt, legacy-binary)
const parts = await readWorkbookParts(pkg);   // { workbook, theme, styles, sharedStrings, colors, warnings }
// toImportExtras(parts): the same, as SpreadsheetImportResult's optional workbook/styles/sharedStrings/theme (S2b wires it in)
// S2b: pkg.xml(sheet.part) / pkg.text(sheet.part) with workbook.sheets[i].rels, resolveColor(c, colors) for sheet-level colours
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

Output types are plain JSON-like data in `@b32nio/spindle-shared/export/xlsx` (`package-types.ts`), reached additively through the
optional `workbook` and `styles` fields of `SpreadsheetImportResult` once S2b populates them.

**Streaming-ready (S20):** every part reader is a pure function from a parsed part to structured-cloneable data, with no DOM and
no shared state; the package class holds nothing but the zip and small relationship lists. Workbook-level parts use the bounded
element tree. `sheetData` is deliberately not read here: S2b reads it as a SAX-style event stream from `pkg.text()`/`pkg.bytes()`.

Tests: one file per module (malformed and producer-variety cases, an oracle table for tints), plus `index.test.ts`, which runs the
whole corpus (`fixtures/xlsx`: openpyxl, xlsxwriter, LibreOffice, hand-authored and edge files).
