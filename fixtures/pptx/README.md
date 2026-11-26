# PPTX fidelity corpus

A set of `.pptx` / `.pptm` / `.ppt` files that exercise what real presentations contain, plus a manifest that says
exactly which constructs each file holds. There is no PPTX importer yet, so there are no goldens: the runner and the
goldens arrive with the loss detector (plan K16) and the importer (K9 onwards). Until then the corpus is verified
against its own claims by the construct probes in `scripts/pptx-fidelity/constructs.ts`.

## What is here

- **Hand-authored constructs** (about thirty files): text (basic and rich, right-to-left and CJK), masters and layouts,
  themes, every preset geometry, custom geometry, fills and lines, effects, five slide sizes, images (PNG, JPEG,
  GIF, SVG with a PNG fallback, EMF, WMF, crops, transforms, a linked picture), audio and video stubs, tables
  (merges, banding, borders, a custom table style), groups and connectors, charts with embedded workbooks, SmartArt,
  transitions and animations, notes, comments (legacy and modern), hidden slides, sections and custom shows, WordArt,
  ink, an OLE object, an equation, action hyperlinks, a kitchen-sink deck and an inert `.pptm`.
- **Producer variety**: a deck written by python-pptx and the same deck converted by LibreOffice headless. The tool
  and version are recorded in `source` in the manifest. The libraries are authoring tools only: they are not
  dependencies and are not installed in CI.
- **Edge files**: truncated zip, no `presentation.xml`, a listed slide whose part is missing, relationships to
  missing parts, malformed slide XML, 60 nested groups, relationship cycles, an encrypted-presentation OLE stub and
  a legacy `.ppt` OLE stub. The OLE stubs hold only a header and a directory, no content.
- **Never committed** (`generatedAtRuntime` in the manifest): `edge-1000-slides` (a deck of 1000 slides, about
  900 KB). Write it with `python3 fixtures/pptx/generators/build_runtime.py --out DIR [name ...]`.

Every file is original content made by the scripts in [`generators/`](./generators). The hand-authored and edge
files use a stdlib `zipfile` builder (`pptxbuilder.py`: fixed timestamps, no randomness, one theme, master, layouts,
slides and every relationship and content-type entry) and rebuild byte-identically; a Jest test checks that. The
producer-variety files are written by their library and then normalised: `docProps/core.xml` is replaced (no author
names, fixed dates), `docProps/app.xml` keeps only the application name, dates inside embedded workbooks are fixed,
LibreOffice's random field GUIDs and chart axis ids are renumbered and the zip is rewritten with fixed timestamps
and order, so rebuilding them locally gives the same bytes. The python-pptx deck carries python-pptx's default
template (MIT). The macro and OLE payloads are readable placeholder text, never a real macro or object, the audio
and video are a 144-byte WAV of silence and a 32-byte MP4 stub, and the JPEG is a 16 x 16 picture stored as
base64 in `generators/assets.py` so that the generators stay stdlib-only. No file is over 300 KB and core
properties are normalised (no author, fixed dates). All preset geometries come from the `PRESETS` list in
`generators/build_shapes.py` (the 187 values of `ST_ShapeType`), 24 per slide, each shape labelled with its name.

## Manifest

[`manifest.json`](./manifest.json) has the same schema as the XLSX and DOCX ones (`id`, `filename`, `title`,
`source{tool, version?, os}`, `license`, `constructs[]`, `notes`, `generator`, `expectImportError?`) plus
`generatedAtRuntime?` for files that are never committed. It is written by `generators/build_all.py` from
`generators/catalog.py`. The generic modules (manifest validation, claim verification, coverage matrix, package
reader) are `scripts/fidelity-core`; `scripts/pptx-fidelity` binds them to the PPTX registry and extensions.

## Checks

The Jest suite in `packages/slides-core/src/fidelity/` (part of `npm test`) verifies:

- every construct a fixture claims is present in its package XML, and every construct in the registry is
  claimed by some fixture;
- every committed file is listed and every listed file exists (run-time files must be absent), none over 300 KB;
- licence and producer are recorded, core properties are normalised;
- structure: readable zip with content types, well-formed parts, relationships that resolve and have unique ids,
  every part typed, unique slide, master and layout ids (damaged-on-purpose files are exempt);
- the generated 1000-slide deck holds its claims;
- the stdlib-built files rebuild byte for byte, the opaque payloads are inert, the library-written decks record their
  producer and have normalised embedded workbooks, custom shows and sections agree with the slide list, and the
  kitchen sink holds at least 45 constructs;
- the coverage matrix below is in sync with the manifest.

```bash
npm test -w @b32nio/spindle-slides-core -- fidelity            # run the checks
PPTX_FIDELITY_UPDATE=1 npm test -w @b32nio/spindle-slides-core -- fidelity   # also rewrite the matrix below
npm run test:fidelity:coverage -w @b32nio/spindle-slides-core  # the same suite with coverage of scripts/pptx-fidelity
python3 fixtures/pptx/generators/build_all.py                  # regenerate every file and manifest.json (needs python-pptx and soffice)
python3 fixtures/pptx/generators/build_all.py --out DIR --stdlib   # only the stdlib-built files, into DIR
```

Fixtures are checked in LibreOffice 25.2 as a second opinion (`soffice --headless --convert-to pdf` succeeds for every
file that is not an edge file); that check is run by hand and is not part of CI.

## Coverage matrix

<!-- coverage-matrix:start (generated by PPTX_FIDELITY_UPDATE=1 npm test -w @b32nio/spindle-slides-core -- fidelity; do not edit) -->

| # | Fixture | Title | Produced by |
|--:|---|---|---|
| 1 | `text-basic` | Basic text | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 2 | `text-rich` | Rich text | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 3 | `layouts-masters` | Masters and layouts | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 4 | `themes` | Themes | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 5 | `shapes-presets` | Every preset geometry | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 6 | `shapes-custgeom` | Custom geometry | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 7 | `fills-lines` | Fills and lines | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 8 | `effects` | Effects | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 9 | `sizes-16x9` | Slide size 16:9 | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 10 | `sizes-4x3` | Slide size 4:3 | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 11 | `sizes-a4` | Slide size A4 | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 12 | `sizes-custom` | Custom slide size | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 13 | `sizes-portrait` | Portrait slides | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 14 | `images` | Images | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 15 | `media` | Audio and video | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 16 | `tables` | Tables | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 17 | `groups-connectors` | Groups and connectors | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 18 | `charts` | Charts | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 19 | `smartart` | SmartArt | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 20 | `wordart-ink-ole-equations` | WordArt, ink, an OLE object, an equation and actions | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 21 | `transitions-animations` | Transitions and animations | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 22 | `notes` | Notes | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 23 | `comments-legacy` | Comments (legacy format) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 24 | `comments-modern` | Comments (modern format) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 25 | `hidden-sections` | Hidden slides, sections and custom shows | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 26 | `rtl-cjk` | Right-to-left and CJK text | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 27 | `python-pptx-deck` | Made with python-pptx | python-pptx (Python), written by the library |
| 28 | `libreoffice-deck` | Converted by LibreOffice | python-pptx deck converted by LibreOffice headless (soffice --convert-to pptx) |
| 29 | `kitchen-sink` | Kitchen sink | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 30 | `macro-enabled` | Macro-enabled presentation | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 31 | `edge-truncated-zip` | Edge: truncated zip | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 32 | `edge-no-presentation` | Edge: no presentation part | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 33 | `edge-slide-missing` | Edge: listed slide without its part | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 34 | `edge-bad-rels` | Edge: relationships that point nowhere | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 35 | `edge-invalid-xml` | Edge: malformed slide XML | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 36 | `edge-deep-groups` | Edge: deeply nested groups | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 37 | `edge-relationship-cycles` | Edge: relationship cycles | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 38 | `edge-encrypted` | Edge: encrypted presentation (OLE stub) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 39 | `edge-legacy-ppt` | Edge: legacy .ppt (OLE stub) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |
| 40 | `edge-1000-slides` | Edge: 1000 slides (generated at run time) | programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py) |

| Construct | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12 | 13 | 14 | 15 | 16 | 17 | 18 | 19 | 20 | 21 | 22 | 23 | 24 | 25 | 26 | 27 | 28 | 29 | 30 | 31 | 32 | 33 | 34 | 35 | 36 | 37 | 38 | 39 | 40 |
|---|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|:-:|
| Several slides (`slides.multiple`) | yes | yes | yes | yes | yes |  | yes | yes | yes | yes | yes | yes | yes | yes |  | yes | yes | yes |  | yes | yes | yes | yes | yes | yes | yes | yes | yes | yes |  |  | yes | yes | yes | yes |  | yes |  |  | yes |
| Text runs (`text.basic`) | yes | yes | yes | yes | yes | yes | yes |  | yes | yes | yes | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  | yes |
| Run formatting (bold, italic, underline, strike) (`text.runs`) | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Paragraph alignment other than left (`text.alignment`) | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Paragraph levels (lvl 1 to 8) (`text.levels`) | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |
| Bullet characters (`text.bullets`) | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Auto-numbered paragraphs (`text.numbering`) | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Bullet colour or size (`text.bulletsCustom`) |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Picture bullets (`text.bulletsPicture`) |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Hyperlinks on text runs (`text.hyperlinks`) | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Paragraph spacing (`text.spacing`) | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Highlighted text (`text.highlight`) | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Tab stops (`text.tabStops`) |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Text columns (`text.columns`) |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Vertical text direction (`text.vertical`) |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Shrink-on-overflow autofit (fontScale) (`text.autofit`) |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Resize shape to fit text (`text.autofitShape`) |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Custom text insets (`text.insets`) | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Fields (slide number, date) (`text.fields`) | yes |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Superscript and subscript (`text.baseline`) |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| All caps and small caps (`text.caps`) |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Character spacing (`text.charSpacing`) |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| More than one slide master (`masters.multiple`) |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Handout master (`masters.handoutMaster`) |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Master text styles (`masters.txStyles`) |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Master with its own background (`masters.background`) |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Non-placeholder shapes on a master (`masters.shapes`) |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Layout placeholders (`layouts.placeholders`) |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Distinct placeholder types (`layouts.placeholderKinds`) |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Layout with its own background (`layouts.background`) |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |
| Layout that hides master shapes (`layouts.hideMasterShapes`) |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Slide placeholders that inherit their position (`layouts.inheritance`) | yes |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Date, footer and slide-number placeholders on slides (`layouts.footerPlaceholders`) | yes |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Masters with different themes (`theme.multiple`) |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Colour map that is not the default (`theme.colorMap`) |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Theme with its own font scheme (`theme.fontScheme`) |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Theme East Asian fonts (`theme.fontsEastAsian`) |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Theme effect styles with effects (`theme.effectScheme`) |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Theme colours (`fills.themeColors`) | yes |  |  | yes | yes | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Colour transforms (lumMod, tint, shade, satMod) (`fills.colorTransforms`) |  |  |  | yes | yes |  | yes |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Solid shape fills (`fills.solid`) |  |  |  | yes | yes | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Shapes with no fill (`fills.noFill`) |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Gradient fills (`fills.gradient`) |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |
| Path (radial, rectangular) gradients (`fills.gradientPath`) |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Pattern fills (`fills.pattern`) |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Picture fills (`fills.picture`) |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Transparency (alpha) (`fills.alpha`) |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Slide backgrounds (`fills.background`) |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Dashed lines (`lines.dash`) |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Arrowheads (`lines.arrowheads`) |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Line caps and joins (`lines.joinCap`) |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Compound lines (`lines.compound`) |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Shadows (`effects.shadow`) |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Glow (`effects.glow`) |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Reflection (`effects.reflection`) |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Soft edges (`effects.softEdge`) |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Blur (`effects.blur`) |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| 3-D scene (camera, lights) (`effects.scene3d`) |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| 3-D shape (extrusion, material) (`effects.sp3d`) |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Bevels (`effects.bevel`) |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Text effects (`effects.text`) |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Preset geometry shapes (`shapes.presets`) |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |
| Distinct preset geometries (`shapes.presetKinds`) |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Preset geometries with adjustments (`shapes.adjustments`) |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Custom geometry (`shapes.custGeom`) |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Flipped shapes (`shapes.flip`) |  |  |  |  | yes |  |  |  |  |  |  |  |  | yes |  |  | yes |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Rotated shapes (`shapes.rotation`) |  |  |  |  | yes |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Text boxes (`shapes.textBox`) | yes | yes |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| PNG pictures (`images.png`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |
| JPEG pictures (`images.jpeg`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| GIF pictures (`images.gif`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| SVG pictures (`images.svg`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| EMF pictures (`images.emf`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| WMF pictures (`images.wmf`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Cropped pictures (`images.crop`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Rotated or flipped pictures (`images.transform`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Linked (not embedded) pictures (`images.linked`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Pictures with alternative text (`images.altText`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Picture recolouring (duotone, greyscale, brightness, colour change) (`images.recolor`) |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Tables (`tables.basic`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |
| Merged table cells (`tables.merge`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |
| Banded rows or columns (`tables.banding`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Header, first-column, total or last-column formatting (`tables.headerRow`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Tables that reference a table style (`tables.styleRef`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Table styles defined in the package (`tables.customStyles`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Cell borders (`tables.borders`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Cell fills (`tables.cellFill`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Groups (`groups.basic`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Groups inside groups (`groups.nested`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Groups whose children are scaled (chExt differs from ext) (`groups.scaled`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Straight connectors (`connectors.straight`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Elbow connectors (`connectors.elbow`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Curved connectors (`connectors.curved`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Connectors attached to shapes (connection sites) (`connectors.connected`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Bar and column charts (`charts.bar`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |
| Line charts (`charts.line`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Pie charts (`charts.pie`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Scatter charts (`charts.scatter`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Chart data caches (`charts.cache`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Charts with an embedded workbook (`charts.embeddedWorkbook`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  | yes |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| SmartArt diagrams (`smartart.diagram`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| SmartArt cached drawings (`smartart.drawing`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Audio (`media.audio`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Video (`media.video`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Slide transitions (`transitions.basic`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Transitions with automatic advance or click control (`transitions.timed`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| PowerPoint 2010 transitions (with fallback) (`transitions.extended`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Animation timing trees (`animations.timing`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Entrance effects (`animations.entrance`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Exit effects (`animations.exit`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Emphasis effects (`animations.emphasis`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Motion paths (`animations.motionPath`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Builds by paragraph (`animations.build`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Notes pages (`notes.slides`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |
| Notes master (`notes.master`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  | yes | yes | yes |  |  |  |  |  |  |  |  |  |  |  |
| Formatted or bulleted notes (`notes.rich`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Comments (legacy format) (`comments.legacy`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Modern comments (`comments.modern`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Comment replies (`comments.replies`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Comment author lists (`comments.authors`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Hidden slides (`slides.hidden`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Sections (`sections.list`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Custom shows (`sections.customShows`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Right-to-left paragraphs (`text.rtl`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Chinese, Japanese and Korean text (`text.cjk`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| Runs with a language other than en-US (`text.lang`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |
| WordArt (text warp) (`objects.wordArt`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Ink (`objects.ink`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Embedded OLE objects (`objects.ole`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Equations (`objects.equations`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Action hyperlinks (`objects.actions`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |
| Macro actions (never run) (`objects.macroActions`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |
| Macro project (vbaProject.bin) (`macros.vba`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |
| Slide size 4:3 (`sizes.4x3`) |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |  |  |  |  |  |  |
| Slide size 16:9 (`sizes.16x9`) |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Slide size A4 (`sizes.a4`) |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Custom slide size (`sizes.custom`) |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Portrait slides (`sizes.portrait`) |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |
| Truncated zip (starts as a zip, cannot be read) (`edge.truncatedZip`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |  |
| Encrypted package (OLE with EncryptionInfo) (`edge.encrypted`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |
| Legacy .ppt binary (OLE with a PowerPoint Document stream) (`edge.legacyBinary`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |
| Package without ppt/presentation.xml (`edge.noPresentation`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |  |
| Listed slide whose part is missing (`edge.slideMissing`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |  |  |
| Relationships that point nowhere (`edge.brokenRels`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes | yes |  |  |  |  |  |  |
| Part that is not well-formed XML (`edge.invalidXml`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |  |
| 1000 or more slides declared (`edge.manySlides`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |
| Groups nested 20 levels deep or more (`edge.deepGroups`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |  |
| Cycles in the relationship graph (`edge.relCycle`) |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  |  | yes |  |  |  |

<!-- coverage-matrix:end -->
