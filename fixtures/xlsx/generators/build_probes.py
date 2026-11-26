"""Probe workbooks from plan 2.2: 159 formulas, 33 number formats x 3 values, 23 formula-syntax forms."""
import datetime

import xlsxbuilder as X
from probe_data import DYNAMIC, FORMULAS, NUMBER_FORMATS, SYNTAX


def _props(wb, title):
    wb.set_properties({"title": title, "created": datetime.datetime(2026, 1, 15, 9, 0, 0)})


def formulas_core(path):
    import xlsxwriter

    wb = xlsxwriter.Workbook(path, {"use_future_functions": True})
    _props(wb, "Formulas core")
    ws = wb.add_worksheet("Data")
    for i in range(10):
        ws.write(i, 0, i + 1)
        ws.write(i, 1, "x" if i % 2 == 0 else "y")
        ws.write(i, 2, (i + 1) * 10)
    for i, (v, d) in enumerate([(-100, 46023), (30, 46100), (40, 46200), (50, 46300)]):
        ws.write(i, 3, v)
        ws.write(i, 4, d)
    for i, f in enumerate(FORMULAS):
        ws.write_string(i, 7, f)
        if f.split("(")[0] in DYNAMIC:
            ws.write_dynamic_array_formula(i, 8, i, 8, "=" + f)
        else:
            ws.write_formula(i, 8, "=" + f)
    wb.close()
    X.normalise(path)


def numfmt_probe(path):
    import xlsxwriter

    wb = xlsxwriter.Workbook(path)
    _props(wb, "Number formats")
    ws = wb.add_worksheet("Formats")
    for i, code in enumerate(NUMBER_FORMATS):
        f = wb.add_format({"num_format": code})
        ws.write_string(i, 0, code)
        for j, v in enumerate((1234.5678, -0.256, 46024.75)):
            ws.write_number(i, 1 + j, v, f)
    wb.close()
    X.normalise(path)


def formula_syntax(path):
    b = X.Book(title="Formula syntax")
    X.dynamic_array_metadata(b)
    sh = b.sheet("Sheet1")
    for i in range(1, 6):
        sh.set(i, 1, i)
    sh.set(1, 2, 10).set(1, 3, 20)
    sh.set(1, 10, ("f", "_xlfn.SEQUENCE(3)", 1, ' t="array" ref="J1:J3"'), cm="1").set(2, 10, 2).set(3, 10, 3)
    for i, (f, cached, note) in enumerate(SYNTAX, 2):
        sh.set(i, 7, ("f", f, cached, ""))
        sh.set(i, 9, note)
    b.sheet("Sheet2").set(1, 1, 100)
    b.sheet("Q1 Data").set(1, 1, 7)
    b.write(path)
