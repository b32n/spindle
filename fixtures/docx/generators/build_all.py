"""Regenerate every fixture whose generator lives in this directory.

    python3 fixtures/docx/generators/build_all.py                   # all programmatic fixtures
    python3 fixtures/docx/generators/build_all.py revisions-authors # one

LibreOffice-produced fixtures are made by build_libreoffice.py (needs `soffice`); the edge-case
packages by build_edge.py. Everything is deterministic: rerunning leaves the files byte-identical
(LibreOffice output aside, which is committed as a binary and only rebuilt on request).
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import build_corpus as C  # noqa: E402
import build_edge  # noqa: E402
import build_revisions  # noqa: E402
import build_styles_lists  # noqa: E402
import features as F  # noqa: E402
import features2 as F2  # noqa: E402
import features3 as F3  # noqa: E402
from ooxml import sect  # noqa: E402

BUILDERS = {
    "styles-lists": build_styles_lists.build,
    "revisions-authors": build_revisions.build,
    "notes-toc-fields": C.notes_toc_fields,
    "tables-advanced": C.simple("tables-advanced", "Advanced tables", "Nested, merged, bordered, shaded, banded and autofit tables", F.tables),
    "sections-headers": C.sections_headers,
    "images-wrap-crop": C.simple("images-wrap-crop", "Images", "Inline, floating, wrapped and cropped pictures", F2.images),
    "shapes-textboxes-smartart": C.simple("shapes-textboxes-smartart", "Shapes, text boxes, SmartArt", "DrawingML and VML shapes, a group, text boxes and a SmartArt diagram with its fallback drawing", F2.shapes),
    "equations-omml": C.simple("equations-omml", "Equations", "Office Math (OMML): fractions, radicals, integrals, sums, matrices and systems", F2.equations),
    "content-controls-forms": C.simple("content-controls-forms", "Content controls and forms", "Structured document tags and legacy form fields", F2.sdt_forms),
    "fonts-themes": C.fonts_themes,
    "i18n-bidi-cjk": C.simple("i18n-bidi-cjk", "Languages and symbols", "Arabic, Hebrew, Chinese, Japanese, Korean, symbols and emoji", F3.i18n),
    "typography-extras": C.typography_extras,
    "charts-objects": C.charts_objects,
    "long-document": C.long_document,
    "kitchen-sink": C.kitchen_sink,
    "word-compat-kitchen": C.word_compat_kitchen,
    "edge-malformed": build_edge.edge_malformed,
    "edge-odd-valid": build_edge.edge_odd_valid,
    "edge-no-body": build_edge.edge_no_body,
}


def main(argv):
    wanted = argv or sorted(BUILDERS)
    for name in wanted:
        if name not in BUILDERS:
            raise SystemExit(f"unknown fixture {name!r}; known: {', '.join(sorted(BUILDERS))}")
        path = os.path.join(OUT, f"{name}.docx")
        BUILDERS[name](path)
        print(f"wrote {path} ({os.path.getsize(path)} bytes)")


if __name__ == "__main__":
    main(sys.argv[1:])
