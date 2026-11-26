"""Regenerate the XLSX corpus and manifest.json (everything here is deterministic; rerunning leaves files byte-identical).

    python3 fixtures/xlsx/generators/build_all.py                 # every committed fixture + manifest.json
    python3 fixtures/xlsx/generators/build_all.py cell-types      # one fixture (manifest untouched)
    python3 fixtures/xlsx/generators/build_all.py --out DIR --stdlib   # the hand-authored and edge files only, elsewhere

Run-time fixtures (zip bomb, huge dimension, 100k x 10 perf) are not written here: see build_runtime.py.
The producer-variety files need openpyxl, xlsxwriter and LibreOffice (`soffice`) installed locally; they are
authoring tools only and are committed as binaries, never rebuilt in CI.
"""
import argparse
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import build_core as C  # noqa: E402
import build_edge as E  # noqa: E402
import build_probes as P  # noqa: E402
import build_producers as PR  # noqa: E402
from catalog import CATALOG, LICENSE, OS, RUNTIME  # noqa: E402

BUILDERS = {
    "cell-types": C.cell_types, "dates-1904": C.dates_1904, "styles-basic": C.styles_basic, "layout": C.layout, "sheet-properties": C.sheet_properties,
    "names": C.names, "cf-basic": C.cf_basic, "data-validation": C.data_validation,
    "formulas-core": P.formulas_core, "number-format-probe": P.numfmt_probe, "formula-syntax": P.formula_syntax,
    "openpyxl-image": PR.openpyxl_image, "openpyxl-chart": PR.openpyxl_chart, "openpyxl-comment": PR.openpyxl_comment, "xlsxwriter-rich": PR.xlsxwriter_rich,
    "libreoffice-rich": PR.libreoffice_rich,
    "edge-truncated-zip": E.truncated_zip, "edge-no-workbook": E.no_workbook, "edge-bad-rels": E.bad_rels, "edge-xml-entities": E.xml_entities,
    "edge-encrypted": E.encrypted, "edge-legacy-xls": E.legacy_xls, "edge-duplicate-sheet-names": E.duplicate_sheet_names,
    "edge-index-out-of-range": E.index_out_of_range, "edge-invalid-xml": E.invalid_xml, "edge-empty-workbook": E.empty_workbook,
}


def tool_versions():
    import openpyxl
    import xlsxwriter

    lo = subprocess.run(["soffice", "--version"], capture_output=True, text=True, check=True).stdout.split()
    return {"openpyxl": openpyxl.__version__, "xlsxwriter": xlsxwriter.__version__, "libreoffice": lo[1]}


def manifest():
    v = tool_versions()
    tools = {
        "programmatic": ("programmatic: hand-authored OOXML (Python 3 zipfile via fixtures/xlsx/generators/xlsxbuilder.py)", None),
        "openpyxl": ("openpyxl (Python), written by the library", v["openpyxl"]),
        "xlsxwriter": ("xlsxwriter (Python), written by the library", v["xlsxwriter"]),
        "libreoffice": ("openpyxl workbook converted by LibreOffice headless (soffice --convert-to xlsx)", f"LibreOffice {v['libreoffice']}, openpyxl {v['openpyxl']}"),
    }
    entries = []
    for e in [dict(x) for x in CATALOG + RUNTIME]:
        tool, version = tools[e.pop("tool")]
        src = {"tool": tool, **({"version": version} if version else {}), "os": OS}
        entries.append({"id": e.pop("id"), "filename": e.pop("filename"), "title": e.pop("title"), "source": src, "license": LICENSE, **e})
    return {"version": 1, "fixtures": entries}


def main(argv):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("ids", nargs="*")
    ap.add_argument("--out", default=OUT, help="directory to write into (default: fixtures/xlsx; the manifest is only written there)")
    ap.add_argument("--stdlib", action="store_true", help="only the files built without openpyxl, xlsxwriter or LibreOffice (these must rebuild byte-identically)")
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
