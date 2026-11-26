"""Producer-variety files: openpyxl, xlsxwriter and LibreOffice headless.

The libraries are authoring tools only (never dependencies). The three openpyxl files each hold one
feature (image, chart, comment) that makes the importer at the time of writing throw (plan 2.2 item 1).
"""
import datetime
import io
import os
import re
import shutil
import subprocess
import tempfile

import xlsxbuilder as X

ROWS = [("North", 10, 20, 30), ("South", 5, 15, 25), ("East", 7, 8, 9), ("West", 1, 2, 3)]


def _openpyxl_base():
    from openpyxl import Workbook

    wb = Workbook()
    wb.properties.created = wb.properties.modified = datetime.datetime(2026, 1, 15, 9, 0, 0)
    ws = wb.active
    ws.title = "Data"
    ws.append(["Region", "Q1", "Q2", "Q3", "Total"])
    for i, r in enumerate(ROWS, 2):
        ws.append([*r, f"=SUM(B{i}:D{i})"])
    return wb, ws


def openpyxl_image(path):
    from openpyxl.drawing.image import Image

    wb, ws = _openpyxl_base()
    ws.add_image(Image(io.BytesIO(X.png(24, 24))), "H2")
    wb.save(path)
    X.normalise(path)


def openpyxl_chart(path):
    from openpyxl.chart import BarChart, LineChart, Reference

    wb, ws = _openpyxl_base()
    bar = BarChart()
    bar.title = "Sales"
    bar.add_data(Reference(ws, min_col=2, min_row=1, max_col=4, max_row=5), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=2, max_row=5))
    ws.add_chart(bar, "H2")
    line = LineChart()
    line.add_data(Reference(ws, min_col=5, min_row=1, max_row=5), titles_from_data=True)
    ws.add_chart(line, "H20")
    wb.save(path)
    X.normalise(path)


def openpyxl_comment(path):
    from openpyxl.comments import Comment

    wb, ws = _openpyxl_base()
    ws["B2"].comment = Comment("Check this figure.", "Reviewer")
    wb.save(path)
    X.normalise(path)


def xlsxwriter_rich(path):
    import xlsxwriter

    wb = xlsxwriter.Workbook(path, {"use_future_functions": True})
    wb.set_properties({"title": "Fixture", "created": datetime.datetime(2026, 1, 15, 9, 0, 0)})
    ws = wb.add_worksheet("Data")
    ws.write_row(0, 0, ["Region", "Q1", "Q2"])
    for i, r in enumerate([("N", 1, 2), ("S", 3, 4)], 1):
        ws.write_row(i, 0, r)
    ws.write_comment("A2", "A note on the first region.")
    ws.insert_image("E2", "px.png", {"image_data": io.BytesIO(X.png(24, 24)), "x_offset": 2})
    chart = wb.add_chart({"type": "column"})
    chart.add_series({"values": "=Data!$B$2:$B$3"})
    ws.insert_chart("H2", chart)
    ws.conditional_format("B2:B3", {"type": "cell", "criteria": ">", "value": 1, "format": wb.add_format({"bg_color": "#FFC7CE"})})
    ws.data_validation("D1:D3", {"validate": "list", "source": ["a", "b"]})
    ws.write_formula("F1", '=_xlfn.XLOOKUP("S",A2:A3,B2:B3)', None, 3)
    ws.write_formula("F2", "=_xlfn.LET(_xlpm.x,2,_xlpm.x*3)", None, 6)
    ws.write_dynamic_array_formula("F4:F5", "=_xlfn.SORT(A2:A3)")
    ws.write_formula("F8", "=IFERROR(1/0,0)")
    wb.define_name("Tot", "=Data!$B$2:$B$3")
    ws.set_column("A:A", 20)
    ws.freeze_panes(1, 1)
    ws.merge_range("A6:B6", "merged")
    ws.write_rich_string("A8", "a ", wb.add_format({"bold": True}), "b", " c")
    wb.close()
    X.normalise(path)


def libreoffice_rich(path):
    """An openpyxl workbook with many features, converted by `soffice --headless` (so LibreOffice wrote the package)."""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.comments import Comment
    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
    from openpyxl.styles import PatternFill
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo

    wb, ws = _openpyxl_base()
    ws.conditional_formatting.add("B2:B5", CellIsRule(operator="greaterThan", formula=["8"], fill=PatternFill("solid", bgColor="FFC7CE")))
    ws.conditional_formatting.add("C2:C5", ColorScaleRule(start_type="min", start_color="F8696B", end_type="max", end_color="63BE7B"))
    dv = DataValidation(type="list", formula1='"Yes,No"')
    ws.add_data_validation(dv)
    dv.add("K1:K5")
    ws["B2"].comment = Comment("Check this figure.", "Reviewer")
    ws["A8"] = "Lookup"
    ws["B8"] = '=VLOOKUP("South",A2:E5,5,FALSE)'
    ws.merge_cells("H1:J2")
    ws.freeze_panes = "B2"
    ws.sheet_properties.tabColor = "FF8800"
    ws.page_setup.orientation = "landscape"
    ws.add_table(Table(displayName="Table1", ref="A1:E5", tableStyleInfo=TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)))
    chart = BarChart()
    chart.add_data(Reference(ws, min_col=2, min_row=1, max_col=4, max_row=5), titles_from_data=True)
    ws.add_chart(chart, "H5")
    hidden = wb.create_sheet("Hidden")
    hidden["A1"] = 5
    hidden.sheet_state = "hidden"
    wb.defined_names["Quarter1"] = DefinedName("Quarter1", attr_text="Data!$B$2:$B$5")
    tmp = tempfile.mkdtemp()
    try:
        src = os.path.join(tmp, "libreoffice-rich.xlsx")
        out = os.path.join(tmp, "out")
        wb.save(src)
        profile = "-env:UserInstallation=file://" + os.path.join(tmp, "profile")
        subprocess.run(["soffice", profile, "--headless", "--convert-to", "xlsx", "--outdir", out, src], check=True, capture_output=True, timeout=300)
        parts = X.read_zip(os.path.join(out, "libreoffice-rich.xlsx"))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    for name in [n for n in parts if n.startswith("xl/charts/chart")]:  # LibreOffice picks random axis ids
        xml = parts[name].decode()
        for i, old in enumerate(dict.fromkeys(re.findall(r'<c:axId val="(\d+)"', xml))):
            xml = xml.replace(f'val="{old}"', f'val="{500000 + i}"')
        parts[name] = xml.encode()
    X.write_zip(path, parts)
    X.normalise(path)
