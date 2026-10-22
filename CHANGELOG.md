# Changelog

All notable changes to Spindle are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project
follows [Semantic Versioning](https://semver.org/) (pre-1.0: minor versions may
contain breaking changes). All `@b32nio/spindle-*` packages are versioned in
lockstep.

## [Unreleased]

### Added

- **Import loss report.** `ImportLossReport` (in `@b32nio/spindle-shared`) lists what an .xlsx / .docx / .csv import cannot
  preserve (conditional formats, validation, named ranges, charts, comments, tracked changes, footnotes, floating images…).
  `inspectWorkbookFile` / `inspectDocumentFile` (react packages) and `inspectXlsx` / `inspectCsv` / `inspectDocx` (core) check a file
  without importing it; `importWorkbookFile` / `importDocumentFile` / `importFromXlsx` / `importFromDocx` now also return `loss`, and
  `onFileOpened` receives it. See `documentation/import-loss.md`.

### Changed

- **Privacy: image hooks on the docs and slides editors.** The editors no longer fetch typed image URLs themselves.
  `resolveImage(url) => Promise<Blob | dataURL>` enables insert-by-URL (otherwise the By-URL UI is hidden and only
  local files can be inserted); `uploadImage(file) => Promise<url>` stores local files with the host;
  `maxImageBytes` bounds embedded images. See `documentation/images.md`.
- **Privacy: `fontSource` option on the sheets, docs and slides editors.** The hard-coded Google Fonts
  loader in the sheets toolbar is gone. Web fonts now come from the Spindle font mirror
  (`https://fonts.b32n.io/css2`) by default and fall back to system fonts if it is unreachable.
  `fontSource={{ kind: 'local' }}` makes zero font requests; `{ kind: 'css', url }` and
  `{ kind: 'custom', load }` let hosts bring their own source. See `documentation/fonts.md`.

## [0.4.0] - 2026-10-20

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

### Changed

- **Minimum Node.js is now 22** (`engines.node` is `>=22.0.0` in every package);
  Node 20 reached end of life. CI runs the package job on Node 22 and 24; the
  website, deploy and publish jobs run on Node 22.
- Internal dependency ranges bumped to `^0.4.0`.

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

[0.4.0]: https://github.com/b32n/spindle/compare/v0.3.0...v0.4.0
[0.3.1]: https://github.com/b32n/spindle/commit/d04923d
[0.3.0]: https://github.com/b32n/spindle/releases/tag/v0.3.0
[0.2.0]: https://github.com/b32n/spindle/commits/cfaefac
[0.1.0]: https://github.com/b32n/spindle/commit/012b8a9
