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

- **DOCX comments are imported.** `importDocx` reads `word/comments.xml`, `commentsExtended.xml` (replies and the
  resolved state) and `people.xml` (author ids), anchors each thread to its commented text across runs, links,
  paragraphs and table cells, and returns the threads in the additive `DocumentImportResult.comments`
  (`DocCommentThread`, `DocComment` in `@b32nio/spindle-shared/export/docx`); the anchored runs carry the additive
  `commentThreadId` on `DocRunBase`. `importFromDocx` fills the document's comment store, so threads appear in the
  existing comments sidebar with author, initials (additive `Comment.authorInitials`), date, replies and resolved
  state. Overlapping comments are tagged with the innermost thread (all threads keep their anchor quote), a comment
  with no surviving anchor stays as a thread without one, and damaged comment parts import what they hold plus a
  warning. Comments are not written back by `exportDocx` yet.

### Changed

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
