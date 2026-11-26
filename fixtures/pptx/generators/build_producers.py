"""Producer-variety decks: python-pptx and LibreOffice headless.

The libraries are authoring tools only (never dependencies). They write what real producers write (their own template, ids,
relationship targets), which catches parser assumptions that the hand-authored files cannot.
"""
import datetime
import io
import os
import re
import shutil
import subprocess
import tempfile

import pptxbuilder as X


def _python_pptx_deck():
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Emu, Inches, Pt

    prs = Presentation()
    cp = prs.core_properties
    cp.title, cp.author, cp.last_modified_by = "Fixture", "", ""
    cp.created = cp.modified = datetime.datetime(2026, 1, 15, 9, 0, 0)
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = "Made with python-pptx"
    s.placeholders[1].text = "A producer-variety deck"
    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = "Bullets and levels"
    tf = s.placeholders[1].text_frame
    tf.text = "First level"
    for text, level in (("Second level", 1), ("Third level", 2), ("Back to the first", 0)):
        p = tf.add_paragraph()
        p.text, p.level = text, level
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Table, picture and shapes"
    table = s.shapes.add_table(3, 3, Inches(0.5), Inches(1.6), Inches(4), Inches(1.5)).table
    for r in range(3):
        for c in range(3):
            table.cell(r, c).text = f"R{r}C{c}"
    table.cell(0, 0).merge(table.cell(0, 1))
    s.shapes.add_picture(io.BytesIO(X.png(32, 32)), Inches(5), Inches(1.6), Inches(1.5), Inches(1.5))
    for n, kind in enumerate((MSO_SHAPE.ROUNDED_RECTANGLE, MSO_SHAPE.STAR_5_POINT, MSO_SHAPE.RIGHT_ARROW)):
        shape = s.shapes.add_shape(kind, Inches(0.5 + n * 1.6), Inches(4), Inches(1.4), Inches(1.2))
        shape.text = kind.name.split("_")[0].title()
    box = s.shapes.add_textbox(Inches(5), Inches(4), Inches(3.5), Inches(0.8))
    box.text_frame.text = "A text box"
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(20)
    s.notes_slide.notes_text_frame.text = "Speaker notes written by python-pptx."
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "A chart with an embedded workbook"
    data = CategoryChartData()
    data.categories = ["North", "South", "East", "West"]
    data.add_series("2024", (10, 5, 7, 1))
    data.add_series("2025", (20, 15, 8, 2))
    s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.6), Inches(8), Inches(5), data)
    return prs


def python_pptx_deck(path):
    _python_pptx_deck().save(path)
    X.normalise(path, application="python-pptx")


def _stabilise(parts):
    """LibreOffice writes random field GUIDs and chart axis ids on every run: renumber them in order of appearance."""
    for name in [n for n in parts if n.endswith(".xml") and n.startswith("ppt/")]:
        xml = parts[name].decode()
        count = iter(range(1, 10**6))
        xml = re.sub(r'(<a:fld id=")\{[0-9A-Fa-f-]{36}\}', lambda m: m.group(1) + "{00000000-0000-4000-8000-%012X}" % next(count), xml)
        if name.startswith("ppt/charts/chart"):
            for i, old in enumerate(dict.fromkeys(re.findall(r'<c:axId val="(\d+)"', xml))):
                xml = xml.replace(f'val="{old}"', f'val="{500000 + i}"')
        parts[name] = xml.encode()


def libreoffice_deck(path):
    """The python-pptx deck converted by `soffice --headless --convert-to pptx` (so LibreOffice wrote the package)."""
    tmp = tempfile.mkdtemp()
    try:
        src = os.path.join(tmp, "libreoffice-deck.pptx")
        out = os.path.join(tmp, "out")
        _python_pptx_deck().save(src)
        profile = "-env:UserInstallation=file://" + os.path.join(tmp, "profile")
        subprocess.run(["soffice", profile, "--headless", "--convert-to", "pptx", "--outdir", out, src], check=True, capture_output=True, timeout=300)
        parts = X.read_zip(os.path.join(out, "libreoffice-deck.pptx"))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    _stabilise(parts)
    X.write_zip(path, parts)
    X.normalise(path)
