# DOCX fidelity corpus

A set of `.docx` files plus **golden snapshots** of what Spindle's DOCX importer currently does
with them, including everything it loses. This is a *measurement* harness: the goldens record
today's behaviour (losses and all), so a later fidelity fix shows up as a deliberate, reviewed golden
update rather than a silent pass or an unnoticed regression.

## What is checked for each file

For every entry in [`manifest.json`](./manifest.json) the harness runs the editor-level API from
`@b32nio/spindle-docs-core` (`inspectDocx`, `importFromDocx`, `exportToDocx`) and compares three
snapshots under [`goldens/`](./goldens):

| Snapshot | Contents |
|---|---|
| `<id>.model.json` | The imported docs-core model (generated ids normalized to `block-1`, `row-1`, `cell-1`, ... and style-pool keys to `text-style-N` / `paragraph-style-N`; image data URLs replaced by a digest), the import warnings, and `importError` when `importFromDocx` throws. |
| `<id>.loss.json` | The import-loss report (`inspectDocx(bytes, { detail: 'full' })`): what the importer does not bring across, by loss id and count, including the formatting/structure ids (table formatting, tabs and breaks, sections, bookmarks…) and a per-kind `where` breakdown for fields. The default `detail: 'summary'` report, which drives the import-confirmation dialog, is narrower. |
| `<id>.roundtrip.json` | Import -> `exportToDocx` -> import again, then a structured diff of the two normalized models: totals, per-field counts, block/text counts before and after, and the first differences. |

Round-trip diffs and losses are **expected** to be non-empty today. CI fails only when a snapshot
*changes* (or a golden is missing/orphaned), never because a loss exists. Files that make
`importFromDocx` throw (an empty or malformed package) are recorded as an `importError` and still
pass; mark them `"expectImportError": true` in the manifest.

## Running it

```bash
npm run build                          # the CLI uses the built packages
npm run docs:docx-fidelity             # compare everything with the goldens (exit 1 on drift)
npm run docs:docx-fidelity -- --only revisions-authors,styles-lists
npm run docs:docx-fidelity -- --matrix # verify manifest claims and print the coverage matrix
npm run docs:docx-fidelity:update      # rewrite goldens + the matrix below, then review `git diff`
npm test -w @b32nio/spindle-docs-core -- fidelity   # the same checks as Jest tests (runs in `npm test`)
```

`DOCX_FIDELITY_UPDATE=1 npm test -w @b32nio/spindle-docs-core -- fidelity` also rewrites goldens.

## The CI report

The packages job runs the CLI with `--report`, which turns the run into something a reviewer can read
without opening the goldens:

```bash
npm run docs:docx-fidelity -- --report md                         # print the markdown report
npm run docs:docx-fidelity -- --report md --report-out out.md     # append it to a file
npm run docs:docx-fidelity -- --report md --report-out "$GITHUB_STEP_SUMMARY" --report json --report-out docx-fidelity-report.json
```

In CI the markdown goes to the **job summary** (Actions run page, "Packages" job, one per Node version)
and the JSON is uploaded as the `docx-fidelity-report-node<version>` artifact (14 days). The report
lists, in this order:

1. The verdict (all snapshots match, or which ones differ and how to update them) and one line of totals.
2. A per-file table: import ok/error, warning count, loss occurrences and types, round-trip result
   (clean, number of differences, or the failing stage), and whether the `model`, `loss` and
   `roundtrip` snapshots matched.
3. Loss totals by loss id across the corpus.
4. Collapsible details per file (files with drift or problems first, and open): import error, warnings,
   the itemised loss table, and the round-trip summary (differences by kind, block and character counts
   before and after, the most frequent differing fields and the first differences).

The summary stays under GitHub's 1 MiB limit by dropping per-file details (never the overview) and says
which files were omitted; the JSON artifact always has everything. A loss or a round-trip difference is
**not** a failure: only snapshot drift, a missing/orphaned golden, a false manifest claim or an
`expectImportError` that disagrees with reality fails the job.

The report code is format-neutral (`scripts/docx-fidelity/report.ts`: `FidelityReport`, `formatMarkdown`,
`formatJson`). The XLSX and PPTX corpora can map their own harness results onto `FidelityReport` and reuse it.

Fields (`fields.*`) are imported as generic fields with their cached result, so the loss goldens count only tables of
contents, indexes and legacy form fields under `docx.fields`; the `model.json` goldens show each field's instruction and
result, and the remaining `Unsupported field code` warnings are fields in headers and footers.

Bookmarks (`fields.bookmarks`) are imported as named positions and `w:hyperlink w:anchor` as `#name` links, so the loss goldens
no longer report `docx.bookmarks`; the `model.json` goldens show `bookmark` items (one per bookmark in the source, `_GoBack`
aside) and `href: "#name"` links, and a hyperlink wrapped around a `PAGEREF` field keeps only the field.

When an importer change shrinks a loss or alters the model, the goldens will fail. Re-run with
`docs:docx-fidelity:update`, read the diff, and commit the new goldens in the same PR as the importer
change. An improvement should be visible in review as a smaller loss snapshot.

## Adding a fixture

1. Put the file in this directory (`<id>.docx`, kebab-case id) and keep it small (aim for under about
   2 MB; the long document may be a few MB).
2. Add an entry to `manifest.json` (schema below). Every construct you list must really be present
   in the file: the harness probes the OOXML (`scripts/docx-fidelity/constructs.ts`) and refuses
   to run when a claim is false. Add a probe there if you introduce a new construct.
3. `npm run docs:docx-fidelity:update`, review the new files under `goldens/`, commit them with the
   fixture. The update also refreshes the coverage matrix below.

### Manifest entry

```jsonc
{
  "id": "revisions-authors",            // kebab-case, unique; names the goldens
  "filename": "revisions-authors.docx",
  "title": "...",
  "source": { "tool": "who/what produced it, with version", "os": "..." },
  "license": "who owns it and under what terms",
  "constructs": ["revisions.insertions", "comments"],   // ids from scripts/docx-fidelity/constructs.ts
  "notes": "what is in the file",
  "generator": "generators/build_revisions.py",         // optional, for reproducible files
  "expectImportError": false                            // optional: file is deliberately broken
}
```

## License-clean

Every file here is authored for this repository: either generated by the scripts in
[`generators/`](./generators) or produced by running a local tool (for example LibreOffice
headless) on content we wrote. No third-party copyrighted documents, no downloaded fonts, images or
text. Images are tiny original geometric PNGs. Any embedded font must be a freely licensed font that
was already on the generating machine; its family, file and license are recorded in the manifest
entry. The manifest states, per file, the tool that really produced it. Files made programmatically
are labelled as such, never as Microsoft Word or Google Docs.

Authentic exports from Word (Windows/Mac) and Google Docs are welcome as drop-in replacements for
spot-checks, provided we wrote the content.

## Where the files come from

| Producer | Files | Notes |
|---|---|---|
| python-docx 1.2.0 + lxml | `styles-lists` | python-docx's default template (MIT), post-processed for multilevel, restarted lists. |
| Hand-authored OOXML (Python 3 `zipfile`, [`generators/ooxml.py`](./generators/ooxml.py)) | most others | For constructs libraries cannot emit: revisions, comments with replies, OMML, content controls, wrapping, SmartArt, charts, embedded fonts, damaged packages. Word-compatible markup, **not produced by Word**. |
| LibreOffice 25.2.3.2 (headless) | `libreoffice-complex`, `libreoffice-review` | A programmatic DOCX re-saved by `soffice --convert-to docx`, so the markup is LibreOffice's own. |

Microsoft Word and Google Docs were not available when the corpus was built, so no file claims to
come from them. `word-compat-kitchen` is Word-*shaped* (rsids, paraIds, proofErr, smartTags, compat
settings) but programmatic, and says so in the manifest.

## Regenerating the files

The `.docx` files are committed binaries; the generators exist so they are reproducible and
reviewable.

```bash
python3 fixtures/docx/generators/build_all.py                  # all programmatic fixtures (byte-identical on rerun)
python3 fixtures/docx/generators/build_all.py kitchen-sink     # one fixture
python3 fixtures/docx/generators/build_libreoffice.py          # needs soffice; output is version-specific
npm run docs:docx-fidelity:update                              # then refresh the goldens and review the diff
```

`build_all.py` needs Python 3 with `lxml` and `python-docx` (only `styles-lists` uses python-docx).
Images are tiny geometric PNGs drawn by the script, and the one embedded font is read from the
system font directory at generation time (see its manifest entry for family, path and licence).
`edge-malformed` is intentionally not a valid package: LibreOffice cannot open it either.

## Coverage matrix

<!-- coverage-matrix:start (generated by `npm run docs:docx-fidelity -- --matrix`; do not edit) -->

| # | Fixture | Title | Produced by |
|--:|---|---|---|
| 1 | `charts-objects` | Charts and embedded objects | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 2 | `content-controls-forms` | Content controls and forms | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 3 | `edge-malformed` | Malformed package | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 4 | `edge-no-body` | Package without a main document | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 5 | `edge-odd-valid` | Odd but valid | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 6 | `equations-omml` | Equations | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 7 | `fonts-themes` | Embedded font and custom theme | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 8 | `i18n-bidi-cjk` | Languages and symbols | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 9 | `images-wrap-crop` | Images | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 10 | `kitchen-sink` | Kitchen sink | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 11 | `libreoffice-complex` | Harbour works programme (LibreOffice) | LibreOffice 25.2.3.2 520(Build:2), headless re-save (soffice --convert-to docx:"MS Word 2007 XML") of a programmatically generated DOCX |
| 12 | `libreoffice-review` | Review of the harbour bye-laws (LibreOffice) | LibreOffice 25.2.3.2 520(Build:2), headless re-save (soffice --convert-to docx:"MS Word 2007 XML") of a programmatically generated DOCX |
| 13 | `long-document` | The Marguerite Harbour Handbook (long) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 14 | `notes-toc-fields` | Fields, notes and navigation | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 15 | `revisions-authors` | Revision review: quarterly harbour report | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 16 | `sections-headers` | Sections and headers | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 17 | `shapes-textboxes-smartart` | Shapes, text boxes, SmartArt | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 18 | `styles-lists` | Styles and lists | programmatic: python-docx 1.2.0 + lxml (python-docx default template), post-processed for numbering |
| 19 | `tables-advanced` | Advanced tables | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 20 | `typography-extras` | Typography extras | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |
| 21 | `word-compat-kitchen` | Quarterly harbour operations report (Word-compatible, programmatic) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/docx/generators/ooxml.py) |

| Construct | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12 | 13 | 14 | 15 | 16 | 17 | 18 | 19 | 20 | 21 |
|---|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| Tracked insertions (`revisions.insertions`) |  |  |  |  | yes |  |  |  |  | yes |  | yes |  |  | yes |  |  |  |  |  |  |
| Tracked deletions (`revisions.deletions`) |  |  |  |  | yes |  |  |  |  | yes |  | yes |  |  | yes |  |  |  |  |  |  |
| Tracked moves (`revisions.moves`) |  |  |  |  |  |  |  |  |  | yes |  |  |  |  | yes |  |  |  |  |  |  |
| Tracked formatting changes (`revisions.formatting`) |  |  |  |  |  |  |  |  |  | yes |  |  |  |  | yes |  |  |  |  |  |  |
| Revisions by 3+ authors (`revisions.multiAuthor`) |  |  |  |  |  |  |  |  |  | yes |  | yes |  |  | yes |  |  |  |  |  |  |
| Comments (`comments`) |  |  |  |  |  |  |  |  |  | yes |  | yes |  |  | yes |  |  |  |  |  | yes |
| Threaded comment replies (`comments.threaded`) |  |  |  |  |  |  |  |  |  | yes |  | yes |  |  | yes |  |  |  |  |  | yes |
| Footnotes (`notes.footnotes`) |  |  | yes |  |  |  |  |  |  | yes | yes |  | yes | yes |  |  |  |  |  |  | yes |
| Endnotes (`notes.endnotes`) |  |  |  |  |  |  |  |  |  | yes | yes |  |  | yes |  |  |  |  |  |  |  |
| Table of contents field (`fields.toc`) |  |  |  |  |  |  |  |  |  | yes | yes |  | yes | yes |  |  |  |  |  |  | yes |
| PAGE field (`fields.page`) |  |  |  |  | yes |  |  |  |  | yes | yes |  | yes | yes |  | yes |  |  |  | yes | yes |
| PAGEREF field (`fields.pageref`) |  |  |  |  |  |  |  |  |  | yes | yes |  | yes | yes |  |  |  |  |  |  | yes |
| REF field (`fields.ref`) |  |  |  |  |  |  |  |  |  | yes | yes |  | yes | yes |  |  |  |  |  |  | yes |
| SEQ field (`fields.seq`) |  |  |  |  |  |  |  |  | yes | yes | yes |  | yes | yes |  |  |  |  |  |  | yes |
| DATE field (`fields.date`) |  |  |  |  |  |  |  |  |  | yes | yes |  |  | yes |  |  |  |  |  |  | yes |
| Other fields (NUMPAGES, AUTHOR, STYLEREF, ...) (`fields.other`) |  |  |  |  |  |  |  |  |  | yes |  |  |  | yes |  |  |  |  |  |  | yes |
| Bookmarks (`fields.bookmarks`) |  |  |  |  |  |  |  |  |  | yes | yes |  | yes | yes |  |  |  |  |  |  | yes |
| External hyperlinks (`hyperlinks.external`) |  |  | yes |  | yes |  |  |  |  | yes | yes |  | yes | yes |  |  |  |  |  |  | yes |
| Internal hyperlinks (anchors) (`hyperlinks.internal`) |  |  |  |  |  |  |  |  |  | yes | yes |  | yes | yes |  |  |  |  |  |  | yes |
| Headings (`headings`) | yes | yes |  |  |  | yes | yes | yes | yes | yes | yes | yes | yes | yes | yes | yes | yes | yes | yes | yes | yes |
| Numbered / bulleted lists (`lists.numbered`) |  | yes | yes |  |  |  |  | yes |  | yes | yes | yes | yes |  | yes |  | yes | yes |  |  |  |
| Multilevel lists (`lists.multilevel`) |  |  |  |  |  |  |  |  |  | yes | yes |  | yes |  |  |  |  | yes |  |  |  |
| List restarts / start overrides (used by a paragraph) (`lists.restart`) |  |  |  |  |  |  |  |  |  | yes |  |  | yes |  |  |  |  | yes |  |  |  |
| Style inheritance (basedOn) (`styles.inheritance`) |  |  |  |  |  |  |  |  |  | yes | yes |  |  | yes | yes |  |  | yes |  |  |  |
| Linked paragraph/character styles (`styles.linked`) |  |  |  |  |  |  |  |  |  | yes | yes |  |  | yes | yes |  |  | yes |  |  |  |
| Nested tables (`tables.nested`) |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  | yes |  | yes |
| Horizontally merged cells (`tables.mergeH`) |  |  |  |  | yes |  |  |  |  | yes | yes |  |  |  |  |  |  |  | yes |  | yes |
| Vertically merged cells (`tables.mergeV`) |  |  |  |  | yes |  |  |  |  | yes | yes |  |  |  |  |  |  |  | yes |  | yes |
| Table / cell borders (`tables.borders`) |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  | yes |  | yes |
| Cell shading (`tables.shading`) |  | yes |  |  |  |  | yes |  |  | yes | yes | yes | yes |  |  |  |  |  | yes |  | yes |
| Repeating header rows (`tables.headerRow`) |  |  |  |  |  |  |  |  |  | yes | yes | yes | yes |  | yes |  |  |  | yes |  | yes |
| Autofit tables (`tables.autofit`) |  |  |  |  | yes |  |  |  |  | yes |  |  |  |  |  |  |  |  | yes |  | yes |
| Landscape sections (`sections.landscape`) |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  | yes |  |  |  |  |  |
| Portrait sections (`sections.portrait`) |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  | yes |  |  |  |  |  |
| Newspaper columns (`sections.columns`) |  |  |  |  |  |  |  |  |  | yes | yes |  | yes |  |  | yes |  |  |  |  |  |
| Different first-page header/footer (`sections.firstPage`) |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  | yes |  |  |  |  |  |
| Odd/even headers and footers (`sections.oddEven`) |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  | yes |  |  |  |  |  |
| Page-numbering restart (`sections.pageNumberRestart`) |  |  |  |  |  |  |  |  |  | yes | yes |  | yes |  |  | yes |  |  |  |  |  |
| Inline images (`images.inline`) |  | yes | yes |  |  |  |  |  | yes | yes | yes |  | yes |  |  |  |  |  |  |  | yes |
| Floating images (`images.floating`) |  |  |  |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  | yes |
| Wrapped (square/tight/top-bottom) images (`images.wrapped`) |  |  |  |  | yes |  |  |  | yes | yes | yes |  |  |  |  |  | yes |  |  |  | yes |
| Cropped images (`images.cropped`) |  |  |  |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  | yes |
| Drawing shapes (`shapes`) |  |  |  |  | yes |  |  |  |  | yes |  |  |  |  |  |  | yes |  |  |  | yes |
| Text boxes (`textboxes`) |  |  |  |  | yes |  |  |  |  | yes |  |  |  |  |  |  | yes |  |  |  | yes |
| SmartArt (with fallback drawing) (`smartart`) |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  | yes |  |  |  | yes |
| OMML equations (`equations`) |  |  |  |  |  | yes |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Content controls (SDT) (`sdt`) |  | yes |  |  | yes |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  | yes |
| Legacy form fields (`forms.legacy`) |  | yes |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  | yes |
| Embedded fonts (`fonts.embedded`) |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Theme colours / fonts used in content (`themes`) |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Right-to-left / bidi text (`i18n.rtl`) |  |  |  |  |  |  |  | yes |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| CJK text (`i18n.cjk`) |  |  |  |  | yes |  |  | yes |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Symbols (`text.symbols`) |  |  |  |  |  |  |  | yes |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Emoji (`text.emoji`) |  |  |  |  |  |  |  | yes |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Hyphenation (`typography.hyphenation`) |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  | yes |  |
| Tabs with leaders (`typography.tabLeaders`) |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  | yes |  |
| Drop caps (`typography.dropCap`) |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  | yes |  |
| Watermarks (`watermark`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |
| Charts (`charts`) | yes |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  | yes |
| Embedded objects (`objects.embedded`) | yes |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  | yes |
| Hidden (vanish) text (`text.hidden`) |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Custom XML parts (`customXml`) |  | yes |  |  | yes |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  | yes |
| Long document (15,000+ words) (`doc.long`) |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |
| Content in headers/footers (`headers.content`) |  |  |  |  | yes |  |  |  |  | yes | yes |  | yes |  |  | yes |  |  |  | yes | yes |
| Relationship targets that do not exist (`edge.brokenRels`) |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Part that is not well-formed XML (`edge.invalidXml`) |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Missing optional parts (styles/settings/numbering) (`edge.missingOptionalParts`) |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |

<!-- coverage-matrix:end -->
