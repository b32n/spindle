"""Regenerate the formula conformance oracle that the sheets-core engine tests check against (plan S11, L5).

    python3 -I fixtures/xlsx/generators/build_conformance_oracle.py

Two groups of rows are recalculated by LibreOffice headless (a throwaway profile with OOXMLRecalcMode = 0, "always
recalculate on load"), the same method as build_syntax_oracle.py:

  * the 159-formula probe (fixtures/xlsx/formulas-core.xlsx, formulas from probe_data.FORMULAS, formula in column I of
    sheet "Data" next to its data), and
  * the cases of conformance_cases.py, each authored with xlsxwriter (`use_future_functions`, so Excel's `_xlfn.`
    prefixes are written) on a sheet of its own with the formula in AA1.

A row's `expected` is LibreOffice's value (`oracle: "libreoffice"`). Where LibreOffice differs from Excel the row
carries the documented Excel value instead, tagged `excel: true, unverified: true`, with LibreOffice's value in
`libreoffice` and the reason in `note`. An unexplained LibreOffice error, or a probe row whose recalculated value
disagrees with probe_data.FORMULA_EXPECTED, stops the script.

LibreOffice, openpyxl and xlsxwriter are authoring tools only: the output is committed, nothing here runs in CI.
"""
import datetime
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
import xlsxwriter  # noqa: E402
from openpyxl.utils.datetime import to_excel  # noqa: E402
from xlsxwriter.utility import xl_cell_to_rowcol  # noqa: E402

from conformance_cases import CASES  # noqa: E402
from probe_data import FORMULAS, FORMULA_EXPECTED  # noqa: E402

PROBE = os.path.join(ROOT, "fixtures", "xlsx", "formulas-core.xlsx")
OUT = os.path.join(ROOT, "packages", "sheets-core", "src", "formula-conformance-oracle.json")

PROFILE = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""

# How an error literal in a case's input cells is authored (the cell then holds that error in LibreOffice).
ERROR_CELLS = {"#DIV/0!": "=1/0", "#N/A": "=NA()", "#VALUE!": '=1+"a"', "#NUM!": "=SQRT(-1)"}

# The probe values LibreOffice cannot give: its 25.2 build has no TEXTBEFORE, TEXTSPLIT, HSTACK, VSTACK or TAKE (it
# returns #NAME?), and MODE of all-distinct numbers is #VALUE! there but #N/A in Excel.
PROBE_EXCEL = {
    "TEXTBEFORE": "LibreOffice 25.2 has no TEXTBEFORE; value from Excel's documentation",
    "TEXTSPLIT": "LibreOffice 25.2 has no TEXTSPLIT; first cell of the spilled result as Excel gives it",
    "HSTACK": "LibreOffice 25.2 has no HSTACK; first cell of the spilled result as Excel gives it",
    "VSTACK": "LibreOffice 25.2 has no VSTACK; first cell of the spilled result as Excel gives it",
    "TAKE": "LibreOffice 25.2 has no TAKE; first cell of the spilled result as Excel gives it",
    "MODE": "MODE of all-distinct numbers is #N/A in Excel; LibreOffice gives #VALUE!",
}


def recalculate(path, tmp):
    """Recalculate a workbook in LibreOffice; return the recalculated workbook (values only)."""
    profile = os.path.join(tmp, "profile", "user")
    os.makedirs(profile, exist_ok=True)
    with open(os.path.join(profile, "registrymodifications.xcu"), "w") as f:
        f.write(PROFILE)
    out = os.path.join(tmp, "out")
    subprocess.run(
        ["soffice", f"-env:UserInstallation=file://{os.path.join(tmp, 'profile')}", "--headless", "--convert-to", "xlsx", "--outdir", out, path],
        check=True, capture_output=True, timeout=540,
    )
    return openpyxl.load_workbook(os.path.join(out, os.path.basename(path)), data_only=True)


def clean(value):
    """A cell value as JSON: no blank (a formula never yields one), floats as written."""
    if isinstance(value, datetime.datetime):  # openpyxl reads date-formatted results as dates: back to the serial number
        serial = to_excel(value)
        return int(serial) if serial == int(serial) else serial
    return "" if value is None else value


def close(a, b):
    if isinstance(a, bool) or isinstance(b, bool) or not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
        return a == b
    return abs(a - b) <= 1e-9 * max(abs(a), abs(b)) or abs(a - b) < 1e-12


def probe_rows(tmp):
    """The 159 formulas of the probe workbook: LibreOffice's value in column I, the file's data in A1:E10."""
    data = openpyxl.load_workbook(PROBE, data_only=False)["Data"]
    cells = {}
    for row in data.iter_rows(min_row=1, max_row=10, max_col=5):
        for c in row:
            if c.value is not None:
                cells[c.coordinate] = c.value
    sheet = recalculate(PROBE, tmp)["Data"]
    rows = []
    for i, formula in enumerate(FORMULAS):
        lo = clean(sheet.cell(row=i + 1, column=9).value)
        wanted = FORMULA_EXPECTED[i]
        name = re.match(r"[A-Z][A-Z0-9.]*", formula)
        fn = name.group(0) if name and formula[len(name.group(0)):len(name.group(0)) + 1] == "(" else "(operator)"
        row = {"id": f"probe-{i + 1:03d}", "function": fn, "formula": formula, "sheet": "Data", "cell": f"I{i + 1}", "cells": {}, "group": "probe"}
        if close(lo, wanted):
            row.update(expectation=kind(lo), expected=lo, oracle="libreoffice")
        elif fn in PROBE_EXCEL:
            row.update(expectation=kind(wanted), expected=wanted, oracle="libreoffice", excel=True, unverified=True, libreoffice=lo, note=PROBE_EXCEL[fn])
        else:
            raise SystemExit(f"probe row {i + 1} {formula}: LibreOffice says {lo!r}, probe_data says {wanted!r}")
        rows.append(row)
    return cells, rows


def kind(value):
    if isinstance(value, list):
        return "array"
    return "error" if isinstance(value, str) and re.fullmatch(r"#[A-Z0-9/?!]+", value) else "value"


def author_cases(path):
    wb = xlsxwriter.Workbook(path, {"use_future_functions": True})
    for n, c in enumerate(CASES):
        ws = wb.add_worksheet(f"c{n:03d}")
        for ref, v in c["cells"].items():
            row, col = xl_cell_to_rowcol(ref)
            if v is None:
                continue
            if isinstance(v, bool):
                ws.write_boolean(row, col, v)
            elif isinstance(v, (int, float)):
                ws.write_number(row, col, v)
            elif v in ERROR_CELLS:
                ws.write_formula(row, col, ERROR_CELLS[v], None, 0)
            elif v.startswith("="):
                ws.write_formula(row, col, v, None, 0)
            else:
                ws.write_string(row, col, v)
        if "spill" in c:
            rows, cols = c["spill"]
            ws.write_array_formula(0, 26, rows - 1, 26 + cols - 1, "{=" + c["formula"] + "}", None, 0)
        elif c.get("array"):
            ws.write_array_formula(0, 26, 0, 26, "{=" + c["formula"] + "}", None, 0)
        else:
            ws.write_formula(0, 26, "=" + c["formula"], None, 0)
    wb.set_properties({"title": "Conformance oracle", "created": datetime.datetime(2026, 1, 15, 9, 0, 0)})
    wb.close()


def case_rows(tmp):
    path = os.path.join(tmp, "conformance-cases.xlsx")
    author_cases(path)
    book = recalculate(path, tmp)
    rows = []
    for n, c in enumerate(CASES):
        ws = book[f"c{n:03d}"]
        if "spill" in c:
            r, k = c["spill"]
            lo = [[clean(ws.cell(row=i + 1, column=27 + j).value) for j in range(k)] for i in range(r)]
        else:
            lo = clean(ws.cell(row=1, column=27).value)
        row = {"id": c["id"], "function": c["function"], "formula": c["formula"], "sheet": None, "cell": "AA1", "cells": c["cells"], "group": "batch-a"}
        if "spill" in c:
            row["spill"] = list(c["spill"])
        if "excel" in c:
            if not c.get("note"):
                raise SystemExit(f"{c['id']}: an excel override needs a note")
            if close(lo, c["excel"]):
                raise SystemExit(f"{c['id']}: LibreOffice already gives the Excel value {lo!r}; drop the override")
            row.update(expectation=kind(c["excel"]), expected=c["excel"], oracle="libreoffice", excel=True, unverified=True, libreoffice=lo, note=c["note"])
        else:
            flat = [v for r in lo for v in r] if isinstance(lo, list) else [lo]
            if any(isinstance(v, str) and v.startswith("#NAME") for v in flat):
                raise SystemExit(f"{c['id']} {c['formula']}: LibreOffice has no such function (#NAME?)")
            row.update(expectation="array" if isinstance(lo, list) else kind(lo), expected=lo, oracle="libreoffice")
            if c.get("note"):
                row["note"] = c["note"]
        rows.append(row)
    return rows


def main():
    with tempfile.TemporaryDirectory() as tmp_probe, tempfile.TemporaryDirectory() as tmp_cases:
        probe_cells, probe = probe_rows(tmp_probe)
        cases = case_rows(tmp_cases)
    lo = subprocess.run(["soffice", "--version"], capture_output=True, text=True, check=True).stdout.split()[1]
    table = {
        "generator": "fixtures/xlsx/generators/build_conformance_oracle.py",
        "libreoffice": lo,
        "probeData": {"sheet": "Data", "cells": probe_cells},
        "rows": probe + cases,
    }
    with open(OUT, "w") as f:
        json.dump(table, f, indent=1, ensure_ascii=False)
        f.write("\n")
    print(f"wrote {len(probe)} probe rows and {len(cases)} batch-A rows to {os.path.relpath(OUT, ROOT)}")


if __name__ == "__main__":
    main()
