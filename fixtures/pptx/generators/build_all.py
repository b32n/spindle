"""Regenerate the PPTX corpus and manifest.json (everything here is deterministic; rerunning leaves files byte-identical).

    python3 fixtures/pptx/generators/build_all.py                 # every committed fixture + manifest.json
    python3 fixtures/pptx/generators/build_all.py text-basic      # one fixture (manifest untouched)
    python3 fixtures/pptx/generators/build_all.py --out DIR --stdlib   # the hand-authored and edge files only, elsewhere

The run-time fixture (1000 slides) is not written here: see build_runtime.py.
The producer-variety files need python-pptx and LibreOffice (`soffice`) installed locally; they are authoring tools
only and are committed as binaries, never rebuilt in CI.
"""
import argparse
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import build_edge as E  # noqa: E402
import build_layouts as L  # noqa: E402
import build_shapes as S  # noqa: E402
import build_text as T  # noqa: E402
from catalog import CATALOG, LICENSE, OS, RUNTIME  # noqa: E402

BUILDERS = {
    "text-basic": T.text_basic, "text-rich": T.text_rich,
    "layouts-masters": L.layouts_masters, "themes": L.themes, "sizes-16x9": L.sizes_16x9, "sizes-4x3": L.sizes_4x3, "sizes-a4": L.sizes_a4, "sizes-custom": L.sizes_custom,
    "sizes-portrait": L.sizes_portrait,
    "shapes-presets": S.shapes_presets, "shapes-custgeom": S.shapes_custgeom, "fills-lines": S.fills_lines, "effects": S.effects,
    "edge-truncated-zip": E.truncated_zip, "edge-no-presentation": E.no_presentation, "edge-slide-missing": E.slide_missing, "edge-bad-rels": E.bad_rels,
    "edge-invalid-xml": E.invalid_xml, "edge-deep-groups": E.deep_groups, "edge-relationship-cycles": E.relationship_cycles,
    "edge-encrypted": E.encrypted, "edge-legacy-ppt": E.legacy_ppt,
}


def tool_versions():
    import pptx

    lo = subprocess.run(["soffice", "--version"], capture_output=True, text=True, check=True).stdout.split()
    return {"python-pptx": pptx.__version__, "libreoffice": lo[1]}


def manifest():
    tools = {"programmatic": ("programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/pptx/generators/pptxbuilder.py)", None)}
    if any(e["tool"] != "programmatic" for e in CATALOG):
        v = tool_versions()
        tools.update({
            "python-pptx": ("python-pptx (Python), written by the library", v["python-pptx"]),
            "libreoffice": ("python-pptx deck converted by LibreOffice headless (soffice --convert-to pptx)", f"LibreOffice {v['libreoffice']}, python-pptx {v['python-pptx']}"),
        })
    entries = []
    for e in [dict(x) for x in CATALOG + RUNTIME]:
        tool, version = tools[e.pop("tool")]
        src = {"tool": tool, **({"version": version} if version else {}), "os": OS}
        entries.append({"id": e.pop("id"), "filename": e.pop("filename"), "title": e.pop("title"), "source": src, "license": LICENSE, **e})
    return {"version": 1, "fixtures": entries}


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("ids", nargs="*")
    ap.add_argument("--out", default=OUT, help="directory to write into (default: fixtures/pptx; the manifest is only written there)")
    ap.add_argument("--stdlib", action="store_true", help="only the files built without python-pptx or LibreOffice (these must rebuild byte-identically)")
    args = ap.parse_args(argv)
    files = {e["id"]: e["filename"] for e in CATALOG}
    wanted = args.ids or [e["id"] for e in CATALOG if not args.stdlib or e["tool"] == "programmatic"]
    os.makedirs(args.out, exist_ok=True)
    for name in wanted:
        if name not in BUILDERS:
            raise SystemExit(f"unknown fixture {name!r}; known: {', '.join(sorted(BUILDERS))}")
        path = os.path.join(args.out, files[name])
        BUILDERS[name](path)
        print(f"wrote {path} ({os.path.getsize(path)} bytes)")
    if not args.ids and not args.stdlib and args.out == OUT:
        with open(os.path.join(OUT, "manifest.json"), "w") as f:
            json.dump(manifest(), f, indent=2, ensure_ascii=False)
            f.write("\n")


if __name__ == "__main__":
    main(sys.argv[1:])
