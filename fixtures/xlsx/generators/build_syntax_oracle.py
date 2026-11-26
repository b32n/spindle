"""Regenerate the formula-syntax oracle table that the sheets-core engine tests check against.

    python3 -I fixtures/xlsx/generators/build_syntax_oracle.py

Each row of the formula-syntax probe (probe_data.SYNTAX, plan 2.2 item 4) is recalculated by LibreOffice headless
(a throwaway profile with OOXMLRecalcMode = 0, "always recalculate on load") and the table records, per formula:
the formula as the file stores it, the same formula in editor syntax, the value the file caches (what Excel
computes), and the value LibreOffice computed. `expected` is LibreOffice's value; where this LibreOffice build has
no such function (`_xlfn.SINGLE`, `_xlfn.ANCHORARRAY` -> `@` and the spill operator) it falls back to the
file's cached value and says so in `source`.

LibreOffice and openpyxl are authoring tools only: the output is committed, nothing here runs in CI.
"""
import json
import os
import re
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
sys.path.insert(0, HERE)

import openpyxl  # noqa: E402

from probe_data import SYNTAX  # noqa: E402

SOURCE = os.path.join(ROOT, "fixtures", "xlsx", "formula-syntax.xlsx")
OUT = os.path.join(ROOT, "packages", "sheets-core", "src", "formula-syntax-oracle.json")

PROFILE = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""


def editor_syntax(formula):
    """The file's `_xlfn.` forms of implicit intersection and the spill operator, as typed in the editor."""
    formula = re.sub(r"_xlfn\.SINGLE\(([^()]*)\)", r"@\1", formula)
    return re.sub(r"_xlfn\.ANCHORARRAY\(([^()]*)\)", r"\1#", formula)


def recalculate():
    with tempfile.TemporaryDirectory() as tmp:
        profile = os.path.join(tmp, "profile", "user")
        os.makedirs(profile)
        with open(os.path.join(profile, "registrymodifications.xcu"), "w") as f:
            f.write(PROFILE)
        out = os.path.join(tmp, "out")
        subprocess.run(
            ["soffice", f"-env:UserInstallation=file://{os.path.join(tmp, 'profile')}", "--headless", "--convert-to", "xlsx", "--outdir", out, SOURCE],
            check=True, capture_output=True, timeout=300,
        )
        sheet = openpyxl.load_workbook(os.path.join(out, "formula-syntax.xlsx"), data_only=True)["Sheet1"]
        return [sheet.cell(row=i, column=7).value for i in range(2, 2 + len(SYNTAX))]


def main():
    computed = recalculate()
    lo = subprocess.run(["soffice", "--version"], capture_output=True, text=True, check=True).stdout.split()[1]
    rows = []
    for (formula, cached, note), value in zip(SYNTAX, computed):
        is_error = isinstance(value, str) and value.startswith("#")
        rows.append({
            "fileFormula": formula,
            "editorFormula": editor_syntax(formula),
            "note": note,
            "cached": cached,
            "libreoffice": value,
            "expected": cached if is_error else value,
            "source": "file cached value (LibreOffice has no such function)" if is_error else "LibreOffice",
        })
        if not is_error and value != cached:
            raise SystemExit(f"{formula}: LibreOffice says {value!r}, the file caches {cached!r}")
    table = {"generator": "fixtures/xlsx/generators/build_syntax_oracle.py", "libreoffice": lo, "rows": rows}
    with open(OUT, "w") as f:
        json.dump(table, f, indent=2, ensure_ascii=False)
        f.write("\n")
    print(f"wrote {len(rows)} rows to {os.path.relpath(OUT, ROOT)}")


if __name__ == "__main__":
    main()
