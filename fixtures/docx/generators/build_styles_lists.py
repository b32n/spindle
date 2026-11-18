"""styles-lists.docx: style inheritance (basedOn / linked), headings, numbered, multilevel and restarted lists.

Producer: python-docx 1.2.0 + lxml (python-docx default template), post-processed for numbering.
"""
import re
import sys
import zipfile
import io
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor, Inches

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from ooxml import prose, write_zip  # noqa: E402


def add_numbering(doc):
    numbering = doc.part.numbering_part.element
    first_num = numbering.find(qn("w:num"))

    def abstract(aid, levels):
        lv = "".join(
            f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{txt}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{720 + 360 * i}" w:hanging="360"/></w:pPr></w:lvl>'
            for i, (fmt, txt) in enumerate(levels)
        )
        return parse_xml(f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{aid}"><w:multiLevelType w:val="multilevel"/>{lv}</w:abstractNum>')

    outline = abstract(90, [("decimal", "%1."), ("decimal", "%1.%2"), ("decimal", "%1.%2.%3"), ("lowerLetter", "(%4)")])
    letters = abstract(91, [("upperLetter", "%1)"), ("lowerRoman", "%2."), ("decimal", "%3.")])
    for a in (outline, letters):
        first_num.addprevious(a)

    def num(nid, aid, overrides=""):
        numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{nid}"><w:abstractNumId w:val="{aid}"/>{overrides}</w:num>'))

    num(90, 90)
    num(91, 91)
    num(92, 90, '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>')  # restart of the outline
    num(93, 90, '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>')  # continues at 5
    return {"outline": 90, "letters": 91, "restart": 92, "five": 93}


def set_num(par, num_id, ilvl=0):
    ppr = par._p.get_or_add_pPr()
    ppr.append(parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr>'))


def build(out):
    doc = Document()
    cp = doc.core_properties
    cp.title, cp.author, cp.subject = "Styles and lists", "Spindle fixture generator", "DOCX fidelity corpus"
    import datetime
    cp.created = cp.modified = datetime.datetime(2026, 1, 15, 9, 0, 0)
    cp.last_modified_by = "Spindle fixture generator"
    cp.revision = 1
    st = doc.styles

    # --- style inheritance: Body Text -> Callout -> Callout Warning -> Callout Warning Strong (paragraph),
    #     Accent Word -> Accent Loud (character), plus a linked paragraph+character pair.
    callout = st.add_style("Callout", WD_STYLE_TYPE.PARAGRAPH)
    callout.base_style = st["Body Text"]
    callout.paragraph_format.left_indent = Inches(0.5)
    callout.paragraph_format.space_after = Pt(10)
    callout.font.italic = True
    callout.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    warn = st.add_style("Callout Warning", WD_STYLE_TYPE.PARAGRAPH)
    warn.base_style = callout
    warn.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
    warn.font.bold = True
    strong = st.add_style("Callout Warning Large", WD_STYLE_TYPE.PARAGRAPH)
    strong.base_style = warn
    strong.font.size = Pt(15)
    strong.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    accent = st.add_style("Accent Word", WD_STYLE_TYPE.CHARACTER)
    accent.font.small_caps = True
    accent.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
    loud = st.add_style("Accent Loud", WD_STYLE_TYPE.CHARACTER)
    loud.base_style = accent
    loud.font.bold = True
    loud.font.size = Pt(13)
    mono = st.add_style("Inline Code", WD_STYLE_TYPE.CHARACTER)
    mono.font.name = "DejaVu Sans Mono"
    mono.font.size = Pt(9.5)
    ref = st.add_style("Reference Item", WD_STYLE_TYPE.PARAGRAPH)
    ref.base_style = st["Normal"]
    ref.paragraph_format.left_indent = Inches(0.4)
    ref.paragraph_format.first_line_indent = Inches(-0.4)
    ref.paragraph_format.line_spacing = 1.5

    ids = add_numbering(doc)

    doc.add_heading("Harbour Records Handbook", 0)
    doc.add_paragraph("A compact reference that exercises styles, heading levels and several kinds of list.", style="Subtitle")

    doc.add_heading("1. Style inheritance", 1)
    doc.add_paragraph(prose(3, 1), style="Body Text")
    doc.add_paragraph("Plain callout: derived from Body Text with an indent and italic blue text.", style="Callout")
    doc.add_paragraph("Warning callout: derived from the callout; bold and red.", style="Callout Warning")
    doc.add_paragraph("Large warning: three levels of basedOn, centred and larger.", style="Callout Warning Large")
    p = doc.add_paragraph(style="Body Text")
    p.add_run("A sentence mixing ")
    p.add_run("accent words", style="Accent Word")
    p.add_run(", ")
    p.add_run("loud accents", style="Accent Loud")
    p.add_run(", and ")
    p.add_run("inline code()", style="Inline Code")
    p.add_run(". Direct formatting follows: ")
    b = p.add_run("bold italic underline")
    b.bold = b.italic = b.underline = True
    p.add_run(", ")
    s = p.add_run("struck through")
    s.font.strike = True
    p.add_run(", ")
    sup = p.add_run("2")
    sup.font.superscript = True
    p.add_run(" and ")
    sub = p.add_run("2")
    sub.font.subscript = True
    p.add_run(".")
    doc.add_paragraph(prose(2, 4), style="Intense Quote")
    doc.add_paragraph("Hanging-indent reference entries (first-line indent via a custom style):", style="Body Text")
    for i in range(3):
        doc.add_paragraph(f"[{i + 1}] {prose(1, 7 + i)} {prose(1, 9 + i)}", style="Reference Item")

    doc.add_heading("2. Headings at every level", 1)
    doc.add_heading("2.1 Second level", 2)
    doc.add_paragraph(prose(2, 2), style="Body Text")
    doc.add_heading("2.1.1 Third level", 3)
    doc.add_paragraph(prose(2, 3), style="Body Text")
    doc.add_heading("2.1.1.1 Fourth level", 4)
    doc.add_paragraph(prose(1, 5), style="Body Text")
    doc.add_heading("2.2 Another second level", 2)
    doc.add_heading("Fifth level and sixth level follow", 5)
    doc.add_heading("Sixth level", 6)

    doc.add_heading("3. Lists", 1)
    doc.add_paragraph("Bulleted list (built-in List Bullet styles, two nesting levels):", style="Body Text")
    for text, style in [("Lanterns", "List Bullet"), ("Oil lamps", "List Bullet 2"), ("Candles", "List Bullet 2"), ("Mirrors", "List Bullet"), ("Ropes", "List Bullet")]:
        doc.add_paragraph(text, style=style)

    doc.add_paragraph("Simple numbered list (List Number):", style="Body Text")
    for t in ("Open the ledger.", "Record the tide.", "Sign the entry."):
        doc.add_paragraph(t, style="List Number")

    doc.add_paragraph("Multilevel outline list (1. / 1.1 / 1.1.1 / (a)):", style="Body Text")
    for lvl, t in [(0, "Inspection"), (1, "Hull"), (2, "Planking"), (3, "Fastenings"), (2, "Keel"), (1, "Rigging"), (0, "Repair"), (1, "Order parts"), (1, "Schedule yard"), (0, "Sign-off")]:
        par = doc.add_paragraph(t, style="Normal")
        set_num(par, ids["outline"], lvl)

    doc.add_paragraph("A new list that restarts at 1 (separate w:num with startOverride):", style="Body Text")
    for t in ("First again", "Second again", "Third again"):
        par = doc.add_paragraph(t, style="Normal")
        set_num(par, ids["restart"], 0)

    doc.add_paragraph("A list that starts at 5:", style="Body Text")
    for t in ("Fifth item", "Sixth item"):
        par = doc.add_paragraph(t, style="Normal")
        set_num(par, ids["five"], 0)

    doc.add_paragraph("Letter / roman list:", style="Body Text")
    for lvl, t in [(0, "Materials"), (1, "Timber"), (1, "Iron"), (2, "Nails"), (0, "Labour"), (1, "Carpenters")]:
        par = doc.add_paragraph(t, style="Normal")
        set_num(par, ids["letters"], lvl)

    doc.add_heading("4. Paragraph formatting", 1)
    for al, name in [(WD_ALIGN_PARAGRAPH.LEFT, "left"), (WD_ALIGN_PARAGRAPH.CENTER, "centre"), (WD_ALIGN_PARAGRAPH.RIGHT, "right"), (WD_ALIGN_PARAGRAPH.JUSTIFY, "justified")]:
        par = doc.add_paragraph(f"This paragraph is {name}-aligned. {prose(2, 6)}", style="Normal")
        par.alignment = al
    par = doc.add_paragraph(f"Double-spaced with space before and after. {prose(2, 8)}", style="Normal")
    par.paragraph_format.line_spacing = 2.0
    par.paragraph_format.space_before = Pt(12)
    par.paragraph_format.space_after = Pt(18)
    par = doc.add_paragraph(f"Left indent 0.75in and first-line indent 0.3in. {prose(2, 10)}", style="Normal")
    par.paragraph_format.left_indent = Inches(0.75)
    par.paragraph_format.first_line_indent = Inches(0.3)
    doc.add_page_break()
    doc.add_heading("5. After a page break", 1)
    doc.add_paragraph(prose(3, 11), style="Body Text")
    doc.add_paragraph("A horizontal-rule-like paragraph border follows.", style="Body Text")
    hr = doc.add_paragraph("")
    hr._p.get_or_add_pPr().append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="8" w:space="1" w:color="808080"/></w:pBdr>'))

    buf = io.BytesIO()
    doc.save(buf)
    # normalise: fixed zip timestamps for reproducible bytes
    with zipfile.ZipFile(io.BytesIO(buf.getvalue())) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    parts.pop("docProps/thumbnail.jpeg", None)  # drop the template's third-party-looking thumbnail
    parts["_rels/.rels"] = re.sub(rb'<Relationship [^>]*thumbnail[^>]*/>', b"", parts["_rels/.rels"])
    write_zip(out, parts)


if __name__ == "__main__":
    build(sys.argv[1])
