# XLSX fidelity corpus

A set of `.xlsx` / `.xlsm` / `.xls` files that exercise what real workbooks contain, plus a manifest
that says exactly which constructs each file holds, plus the goldens that record what the importer does with
each file today (what it keeps, loses, throws on or gets wrong). The goldens describe **current behaviour, not
correctness**: a file that makes the importer throw has a golden saying so, and a later change that fixes it shows
up as a golden diff to review.

## What is here

- **Hand-authored constructs** (about thirty files): cell types, dates (1900 and 1904), styles, theme
  colours, sixty number formats, formulas (modern, syntax probe, references), layout, sheet
  properties, names, conditional formatting (classic and x14), validation, comments (notes and threads),
  links, images, charts, tables, filters, a pivot table with a slicer, sparklines, protection, print
  setup, rich text and an inert `.xlsm`.
- **Probe workbooks** from the fidelity plan: 159 common formulas, 33 number-format codes x 3 values and
  23 formula-syntax forms (cached values in the syntax probe are what Excel computes).
- **Producer variety**: openpyxl (three single-feature files that make the importer throw), xlsxwriter and
  LibreOffice headless. The tool and version are recorded in `source` in the manifest. The libraries are
  authoring tools only: they are not dependencies and are not installed in CI.
- **Edge files**: truncated zip, no workbook part, broken relationships, XML entity declarations, malformed
  XML, an encrypted-workbook OLE stub, a legacy `.xls` OLE stub, duplicate sheet names, indexes out of
  range and an empty workbook. The OLE stubs hold only a header and a directory, no content.
- **Never committed** (`generatedAtRuntime` in the manifest): `edge-zip-bomb` (300 KB that inflates to 300 MB),
  `edge-huge-dimension` (a sheet declaring 17 billion cells) and `perf-100k-x-10` (one million cells).
  Write them with `python3 fixtures/xlsx/generators/build_runtime.py --out DIR [name ...]`.

Every file is original content made by the scripts in [`generators/`](./generators). The hand-authored and edge
files use a stdlib `zipfile` builder (fixed timestamps, no randomness) and rebuild byte-identically; a Jest test
checks that. The producer-variety files are written by their library and then normalised: `docProps/core.xml`
is replaced (no author names, fixed dates), the zip is rewritten with fixed timestamps and order, and the
producer's own `Application` name in `docProps/app.xml` is kept. No file is over 300 KB. Names of people in
comments are invented.

## Manifest

[`manifest.json`](./manifest.json) has the same schema as the DOCX one (`id`, `filename`, `title`,
`source{tool, version?, os}`, `license`, `constructs[]`, `notes`, `generator`, `expectImportError?`) plus
`generatedAtRuntime?` for files that are never committed. `expectImportError` marks files that are broken
on purpose; the runner compares the importer's outcome with it and reports a disagreement on the scoreboard
(L4) instead of failing the run, because the goldens already pin what happens. The manifest is written by
`generators/build_all.py` from `generators/catalog.py`.

## Checks

The Jest suite in `packages/sheets-core/src/fidelity/` (part of `npm test`) verifies, with the construct
probes in `scripts/xlsx-fidelity/constructs.ts`:

- every construct a fixture claims is present in its package XML, and every construct in the registry is
  claimed by some fixture;
- every committed file is listed and every listed file exists (run-time files must be absent), none over 300 KB;
- licence and producer are recorded, core properties are normalised;
- structure: readable zip with content types, well-formed parts, relationships that resolve and have unique ids,
  every part typed, unique sheet names (damaged-on-purpose files are exempt);
- the run-time generators write files that hold their claims;
- the coverage matrix below is in sync with the manifest.

```bash
npm test -w @b32nio/spindle-sheets-core -- fidelity            # run the checks
XLSX_FIDELITY_UPDATE=1 npm test -w @b32nio/spindle-sheets-core -- fidelity   # also rewrite the matrix below
python3 fixtures/xlsx/generators/build_all.py                  # regenerate every file and manifest.json
python3 fixtures/xlsx/generators/build_all.py --out DIR --stdlib   # only the stdlib-built files, into DIR
```

The producer-variety files need openpyxl, xlsxwriter and LibreOffice installed locally to regenerate; their
output is committed as binaries. The LibreOffice file is normalised (chart axis ids, zip order, timestamps) so
rebuilding gives the same bytes.

## Goldens

`goldens/<id>.<kind>.json` holds five snapshots per committed file, written by the runner and compared on every
run (CI fails on any difference). They contain no timestamps, no absolute paths and no random ids, so a run is
byte-identical on Node 22 and 24. The importer never has to succeed: the harness catches everything, and an import
that throws is recorded as `importError` with the message.

| Kind | What it records |
|---|---|
| `model` | The sheets-core workbook after `importFromXlsx`, normalised: cells by address (`A1`), pooled styles and formats renumbered `s1`.. and `f1`.. in position order, row and column ids turned into positions, formula text without the stable-id AST, binary payloads (when the model has any) as size plus SHA-256. Also `importError` and the import `warnings`. |
| `loss` | `inspectXlsx(bytes, { detail: 'full' })` sorted by id. Today the detector ignores the option and reports its summary ids; when S1 adds the full level the goldens pick it up without a change to the runner. |
| `roundtrip` | Import, export, import again, and the structured difference between the two imported models (styles inlined so pool numbering never counts): total, added, removed, changed, counts per field and the first 40 differences. `status` is `ok`, `skipped-import-error`, `export-error` or `reimport-error`. |
| `ledger` | Every part of the input package with a status, see below. |
| `values` | What the engine computes for each formula cell after import against the value the file caches for it (read from the package XML, not from the importer): counts compared, matched and mismatched, mismatches by kind (`number -> error`), and the first 20 with sheet, cell, expected and actual. Formulas without a cached value and formulas calling NOW, TODAY or RAND* are counted but not compared. Numbers match to 9 significant digits; two errors match when the code is the same (`#NAME?`). |

The cached values in `formulas-core.xlsx` are what LibreOffice 25.2.3.2 computes with recalculation forced on load
(the five functions it lacks, TEXTBEFORE, TEXTSPLIT, HSTACK, VSTACK and TAKE, carry the value Excel gives, and MODE
of distinct numbers carries Excel's `#N/A`), so `values` over that file is the plan's 159-formula measurement.
`number-format-probe.xlsx` has no display strings to compare, so the number-format measurement is not part of `values`.

### Package ledger rules

Every part (zip entry) of the input gets exactly one status, decided by the first matching rule in
`scripts/xlsx-fidelity/ledger.ts` (the rule id is recorded next to the part):

- `mapped`: the part's content flows into the model (workbook, worksheets, shared strings, styles) or it is package
  plumbing the export rebuilds (`[Content_Types].xml`, every `.rels`). A mapped part can still hold features the
  model lacks; those are reported through loss ids, not here.
- `opaque`: the part is kept byte for byte and written back on export. Always zero until the opaque part store (plan
  X5) exists; the runner already has the hook.
- `dropped`: not imported, and the loss report covers it. A rule only applies when the file's loss report contains
  one of the rule's ids (`xlsx.charts` for charts and chart sheets, `xlsx.images` for media and drawings,
  `xlsx.comments` for notes, threads and legacy drawings, `xlsx.tables`, `xlsx.pivotTables`, `xlsx.slicers`,
  `xlsx.externalLinks`, `xlsx.macros`); the covering id is listed.
- `unaccounted`: nothing accounts for the part. It is not read and no reported loss covers it: document properties,
  themes, the calc chain, printer settings, custom XML, a drawing holding only shapes. The number is recorded as is
  (launch-bar row L8 wants 0; the plan's S1, S6 and X5 take it there).

If the import throws, nothing reached the model, so `mapped` parts count as `unaccounted` for that file (`dropped`
stays: the loss report is produced independently). A package that cannot be read at all (truncated zip, OLE stub)
has an empty ledger and `package` set to `unreadable` or `ole`.

### Running and updating

```bash
npm run build                                   # the CLI runs against the built @b32nio/spindle-sheets-core
npm run sheets:xlsx-fidelity                    # compare every file with its goldens (exit 1 on drift)
npm run sheets:xlsx-fidelity:update             # rewrite the goldens and the coverage matrix, then review the diff
npm run sheets:xlsx-fidelity -- --only cell-types,links          # a subset
npm run sheets:xlsx-fidelity -- --report md --report-out summary.md --report json --report-out report.json
npm run sheets:xlsx-fidelity -- --matrix        # verify the claims and print the coverage matrix
npm run sheets:xlsx-fidelity -- --runtime       # also import the run-time files (no goldens); --timeout <seconds> per file
npm run fidelity                                # DOCX, then XLSX
XLSX_FIDELITY_UPDATE=1 npm test -w @b32nio/spindle-sheets-core -- fidelity   # same update, bound to the source
```

The report ends with a launch-bar scoreboard (plan section 7.3): L4 imports that did not end as the manifest says,
L5 formulas matching the cached values, L7 round-trip differences and L8 unaccounted parts, each with the measured
number and its target. A row that is not met is the recorded baseline, not a failure. CI writes the markdown to the
job summary and uploads the JSON as `xlsx-fidelity-report-node<version>`. `--report md --report-out FILE` appends to
the file (as `$GITHUB_STEP_SUMMARY` requires) and never truncates it, so delete the file between local runs; JSON
reports are overwritten.

The run-time files (zip bomb, huge dimension, 100,000 x 10) are never part of the default run: `--runtime` writes
them to a temporary directory with `generators/build_runtime.py`, imports each once under the per-file time limit
and prints time, memory and outcome. Importing the bomb takes about 1 GB of memory today (plan section 2.2).

## Coverage matrix

<!-- coverage-matrix:start (generated by XLSX_FIDELITY_UPDATE=1 npm test -w @b32nio/spindle-sheets-core -- fidelity; do not edit) -->

| # | Fixture | Title | Produced by |
|--:|---|---|---|
| 1 | `cell-types` | Cell types | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 2 | `dates-1904` | 1904 date system | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 3 | `styles-basic` | Basic styles | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 4 | `styles-theme` | Theme colours, named styles and fills | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 5 | `number-formats` | Sixty number format codes | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 6 | `number-format-probe` | Number format probe (33 codes x 3 values) | xlsxwriter (Python), written by the library |
| 7 | `formulas-core` | Formulas core: the 159 probe formulas | xlsxwriter (Python), written by the library |
| 8 | `formulas-modern` | Modern formulas | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 9 | `formula-syntax` | Formula syntax probe (23 forms) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 10 | `references` | References | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 11 | `layout` | Layout | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 12 | `sheet-properties` | Sheet properties | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 13 | `names` | Defined names | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 14 | `cf-basic` | Conditional formatting | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 15 | `cf-x14` | Conditional formatting: x14 extensions | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 16 | `data-validation` | Data validation | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 17 | `comments` | Comments: notes and threads | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 18 | `links` | Hyperlinks | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 19 | `images` | Images | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 20 | `charts` | Charts | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 21 | `tables` | Tables | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 22 | `filters-sort` | Filters and sort | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 23 | `pivots-slicers` | Pivot table and slicer | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 24 | `sparklines` | Sparklines | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 25 | `protection` | Protection | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 26 | `print-setup` | Print setup | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 27 | `rich-text` | Rich-text cells | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 28 | `macro-enabled` | Macro-enabled workbook (.xlsm) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 29 | `openpyxl-image` | openpyxl: workbook with an image | openpyxl (Python), written by the library |
| 30 | `openpyxl-chart` | openpyxl: workbook with charts | openpyxl (Python), written by the library |
| 31 | `openpyxl-comment` | openpyxl: workbook with a cell comment | openpyxl (Python), written by the library |
| 32 | `xlsxwriter-rich` | xlsxwriter: many features | xlsxwriter (Python), written by the library |
| 33 | `libreoffice-rich` | LibreOffice: many features | openpyxl workbook converted by LibreOffice headless (soffice --convert-to xlsx) |
| 34 | `edge-truncated-zip` | Edge: truncated zip | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 35 | `edge-no-workbook` | Edge: no workbook part | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 36 | `edge-bad-rels` | Edge: relationships that point nowhere | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 37 | `edge-xml-entities` | Edge: XML entity declarations | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 38 | `edge-invalid-xml` | Edge: malformed worksheet XML | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 39 | `edge-encrypted` | Edge: encrypted workbook (OLE stub) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 40 | `edge-legacy-xls` | Edge: legacy .xls (OLE stub) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 41 | `edge-duplicate-sheet-names` | Edge: duplicate sheet names | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 42 | `edge-index-out-of-range` | Edge: indexes out of range | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 43 | `edge-empty-workbook` | Edge: empty workbook | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 44 | `edge-zip-bomb` | Edge: zip bomb (generated at run time) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 45 | `edge-huge-dimension` | Edge: huge dimension (generated at run time) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |
| 46 | `perf-100k-x-10` | Performance: 100,000 rows x 10 columns (generated at run time) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py) |

| Construct | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12 | 13 | 14 | 15 | 16 | 17 | 18 | 19 | 20 | 21 | 22 | 23 | 24 | 25 | 26 | 27 | 28 | 29 | 30 | 31 | 32 | 33 | 34 | 35 | 36 | 37 | 38 | 39 | 40 | 41 | 42 | 43 | 44 | 45 | 46 |
|---|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| Numeric cells (`cells.numbers`) | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |
| Shared-string cells (`cells.sharedStrings`) | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Inline-string cells (`cells.inlineStrings`) | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  | yes |
| Boolean cells (`cells.booleans`) | yes |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Error cells (`cells.errors`) | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Date / time formatted cells (`cells.dates`) | yes | yes |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Non-ASCII text (`cells.unicode`) | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Right-to-left text (Hebrew, Arabic) (`cells.rtl`) | yes |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Rich-text strings (runs) (`cells.richText`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| 1904 date system (`workbook.date1904`) |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Calculation settings (calcPr) (`workbook.calcSettings`) | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| External workbook links (`workbook.externalLinks`) |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Cell formats beyond the default (`styles.basic`) |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Alignment (wrap, rotation, indent) (`styles.alignment`) |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Custom number formats (`styles.customNumFmts`) |  |  | yes |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Theme colours (`styles.themeColors`) |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Indexed colours (other than the default pen 64) (`styles.indexedColors`) |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Colour tints (`styles.tint`) |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Named cell styles (`styles.namedStyles`) |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Pattern fills (other than solid) (`styles.patternFill`) |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Gradient fills (`styles.gradientFill`) |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Distinct border styles (13 exist) (`styles.borderStyles`) |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Diagonal borders (`styles.diagonalBorder`) |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Formulas (`formulas.basic`) |  |  |  |  |  |  | yes |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Formulas with a cached value (`formulas.cachedValues`) | yes |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Shared formulas (`formulas.shared`) |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Array formulas (t="array") (`formulas.array`) |  |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Dynamic-array formulas (cm metadata) (`formulas.dynamicArray`) |  |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Future functions (_xlfn. prefix) (`formulas.xlfn`) |  |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| LET / LAMBDA parameters (_xlpm.) (`formulas.lambda`) |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Whole-column references (A:A) (`formulas.wholeColumn`) |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Whole-row references (1:1) (`formulas.wholeRow`) |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Cross-sheet references (`formulas.crossSheet`) |  |  |  |  |  |  | yes |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| 3D references (Sheet1:Sheet3!A1) (`formulas.threeD`) |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Structured table references (`formulas.structuredRef`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| External workbook references ([1]Sheet!A1) (`formulas.externalRef`) |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Self-referencing formulas (the cell itself, or a range, column or row that contains it) (`formulas.circular`) |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Defined names (`names.defined`) |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Sheet-scoped names (`names.sheetScoped`) |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Print areas (`names.printArea`) |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Print titles (`names.printTitles`) |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Names holding constants (`names.constants`) |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Custom column widths (`layout.colWidths`) |  |  | yes |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Custom row heights (`layout.rowHeights`) |  |  | yes |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Default column / row sizes (`layout.defaultSizes`) |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Hidden rows and columns (`layout.hidden`) |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Outline groups (`layout.outline`) |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Column default styles (`layout.columnStyles`) |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Frozen panes (`layout.frozenPanes`) |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Merged cells (`layout.merges`) |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Zoom other than 100% (`layout.zoom`) |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Gridlines off (`layout.noGridlines`) |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Right-to-left sheet (`layout.rtl`) |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| More than one sheet (`sheets.multiple`) |  |  |  |  |  |  |  |  |  | yes | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  | yes |  |  |  |  |  |
| Hidden sheets (`sheets.hidden`) |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Very hidden sheets (`sheets.veryHidden`) |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Tab colours (`sheets.tabColor`) |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Active tab other than the first (`sheets.activeTab`) |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Conditional formatting: cell value (`cf.cellIs`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Conditional formatting: formula (`cf.expression`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Conditional formatting: top / bottom (`cf.top10`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Conditional formatting: above / below average (`cf.aboveAverage`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Conditional formatting: duplicate / unique (`cf.duplicates`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Conditional formatting: text rules (`cf.text`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Conditional formatting: dates occurring (`cf.dates`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Conditional formatting: colour scales (`cf.colorScale`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Conditional formatting: data bars (`cf.dataBar`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Conditional formatting: icon sets (`cf.iconSet`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Conditional formatting: stop if true (`cf.stopIfTrue`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Conditional formatting: x14 extensions (`cf.x14`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Validation: list (`dv.list`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Validation: whole number (`dv.whole`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Validation: decimal (`dv.decimal`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Validation: date (`dv.date`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Validation: time (`dv.time`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Validation: text length (`dv.textLength`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Validation: custom formula (`dv.custom`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Validation: input / error messages (`dv.messages`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Comments (notes) (`comments.notes`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Threaded comments (`comments.threaded`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| External hyperlinks (http, https) (`links.external`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| mailto hyperlinks (`links.mailto`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Internal hyperlinks (location) (`links.internal`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Hyperlinks with a script scheme (`links.unsafeScheme`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Images: PNG (`images.png`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  | yes |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Images: JPEG (`images.jpeg`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Images: GIF (`images.gif`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Images: SVG (`images.svg`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Images: EMF (`images.emf`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Images: one-cell anchor (`images.oneCell`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Images: two-cell anchor (`images.twoCell`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Images: absolute anchor (`images.absolute`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Cropped images (`images.crop`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Charts: bar / column (`charts.bar`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  | yes |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Charts: line (`charts.line`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Charts: pie (`charts.pie`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Charts: scatter (`charts.scatter`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Charts: combination (several plot types) (`charts.combo`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Tables (list objects) (`tables.listObjects`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Tables with a totals row (`tables.totals`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Auto-filters (`filters.autoFilter`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Filter criteria (`filters.criteria`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Sort state (`filters.sort`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Pivot tables (`pivots.tables`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Slicers (`pivots.slicers`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Sparklines (`sparklines`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Sheet protection (`protection.sheet`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Sheet protection with a password hash (`protection.password`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Sheet protection without a password (`protection.noPassword`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Workbook structure protection (`protection.workbook`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Page setup (orientation, paper, scale) (`print.pageSetup`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Fit to page (`print.fitToPage`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Print headers and footers (`print.headerFooter`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Manual page breaks (`print.breaks`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Macro project (vbaProject.bin) (`macros.vba`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Truncated zip (starts as a zip, cannot be read) (`edge.truncatedZip`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |
| Encrypted package (OLE with EncryptionInfo) (`edge.encrypted`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |
| Legacy .xls binary (OLE with a Workbook stream) (`edge.legacyBinary`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |
| Package without xl/workbook.xml (`edge.noWorkbook`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Relationships that point nowhere (`edge.brokenRels`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |
| Part that is not well-formed XML (`edge.invalidXml`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |
| XML entity declarations (DOCTYPE) (`edge.xmlEntities`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |
| Duplicate sheet names (`edge.duplicateSheetNames`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |
| Shared-string or style index out of range (`edge.indexOutOfRange`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |
| Workbook whose sheets hold no cells (`edge.emptyWorkbook`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |
| Sheet dimension of 10 million cells or more (`edge.hugeDimension`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |
| Zip entry with a compression ratio of 200:1 or more (1 MB+) (`edge.zipBomb`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |
| Sheet of 100,000 rows or more (`perf.largeSheet`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |

<!-- coverage-matrix:end -->
