"""Feature blocks. Each function adds content (and any parts it needs) to a Doc and returns body XML."""
from docbuilder import PEOPLE, Doc, paragraph_text
from ooxml import (border, br, esc, fld, bookmark, hyperlink_int, num, p, r, rpr, tab, tbl, tc, tr)


def H(text, lvl=1, bm=None, bid=None):
    inner = r(text)
    if bm:
        inner = bookmark(bid, bm, inner)
    return p(inner, style=f"Heading{lvl}")


def body_p(seed, n=4, style="BodyText"):
    return p(paragraph_text(seed, n), style=style)


# ============================================================ TOC / fields
def toc_fields(doc: Doc, headings=None):
    """Heading list with _Toc bookmarks, a TOC field with cached hyperlinked entries (PAGEREF per entry),
    plus PAGE / NUMPAGES / REF / SEQ / DATE / AUTHOR / TITLE / STYLEREF / HYPERLINK / IF fields, bookmarks and links."""
    headings = headings or [(1, "Overview"), (2, "Scope of the survey"), (2, "Method"), (1, "Findings"), (2, "Tonnage"), (3, "Seasonal variation"), (1, "Recommendations")]
    out = [p("Contents", style="TOCHeading")]
    entries = []
    for i, (lvl, text) in enumerate(headings):
        entries.append((lvl, text, f"_Toc{9000 + i}", 1 + i // 2))
    toc = []
    for k, (lvl, text, bm, pg) in enumerate(entries):
        run = (f'<w:hyperlink w:anchor="{bm}" w:history="1">{r(text, style="Hyperlink")}{tab()}'
               + fld(f"PAGEREF {bm} \\h", str(pg)) + "</w:hyperlink>")
        if k == 0:
            begin = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
                     '<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
            run = begin + run
        if k == len(entries) - 1:
            run += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        toc.append(p(run, style=f"TOC{lvl}"))
    out += toc
    out.append(p(""))
    # body: headings with bookmarks + fields
    for i, (lvl, text, bm, pg) in enumerate(entries):
        out.append(H(text, lvl, bm, 500 + i))
        out.append(body_p(40 + i, 3))
    out.append(H("Fields and cross-references", 1, "_TocFields", 599))
    out.append(p(bookmark(700, "keyfigure", r("The key figure of this report is 42 vessels per day.")), style="BodyText"))
    out.append(p(r("As stated on page ") + fld("PAGEREF keyfigure \\h", "3") + r(": ") + fld("REF keyfigure \\h", "The key figure of this report is 42 vessels per day.") + r("."), style="BodyText"))
    out.append(p(r("Figure ") + fld("SEQ Figure \\* ARABIC", "1") + r(" shows the quay; Figure ") + fld("SEQ Figure \\* ARABIC", "2") + r(" shows the lighthouse; Table ") + fld("SEQ Table \\* ARABIC", "1") + r(" lists berths."), style="Caption"))
    out.append(p(r("Generated on ") + fld('DATE \\@ "yyyy-MM-dd"', "2026-01-15") + r(" at ") + fld('TIME \\@ "HH:mm"', "09:00") + r(" by ") + fld("AUTHOR", "Spindle fixture generator") + r(" for ") + fld("TITLE", "Fields and notes") + r("."), style="BodyText"))
    out.append(p(r("Page ") + fld("PAGE", "1") + r(" of ") + fld("NUMPAGES", "9") + r(" (section pages: ") + fld("SECTIONPAGES", "9") + r("; words: ") + fld("NUMWORDS", "800") + r(")."), style="BodyText"))
    out.append(p(r("Current chapter: ") + fld('STYLEREF "Heading 1"', "Fields and cross-references") + r(". Conditional: ") + fld('IF 1 = 1 "yes" "no"', "yes") + r("."), style="BodyText"))
    rid = doc.rel("hyperlink", "https://example.com/harbour", external=True)
    out.append(p(r("External field link: ") + fld(f'HYPERLINK "https://example.com/harbour"', "example.com/harbour", style="Hyperlink") + r(" and run link: ") + doc.link("https://example.org/tides", "example.org/tides") + r("."), style="BodyText"))
    out.append(p(r("Internal links: ") + hyperlink_int("keyfigure", "jump to the key figure") + r(", ") + hyperlink_int("_Toc9003", "jump to Findings") + r("; a mail link: ") + doc.link("mailto:records@example.com", "records@example.com") + r("."), style="BodyText"))
    out.append(p(fld('MERGEFIELD Recipient \\* MERGEFORMAT', "«Recipient»") + r(" (unresolved mail-merge field) and a "), style="BodyText"))
    nested = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>' + fld("PAGE", "1")
              + '<w:r><w:instrText xml:space="preserve"> &gt; 1 "multi-page" "single page" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + r("single page") + '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    out.append(p(r("Simple field: ") + '<w:fldSimple w:instr=" PAGE \\* MERGEFORMAT ">' + r("1") + "</w:fldSimple>" + r(" and a nested field: ") + nested, style="BodyText"))
    return "".join(out)


# ================================================================== notes
def notes(doc: Doc):
    out = [H("Footnotes and endnotes", 1)]
    out.append(p(r("Harbour tonnage rose again last year.") + doc.footnote("Source: the harbour master's annual ledger, table 4.") + r(" A second claim needs support too.")
                 + doc.footnote("Compare the 2019 survey, which counted vessels rather than tonnes.") + r(" The third claim is contested") + doc.endnote("Disputed by the grain exchange; see their memorandum of March.") + r("."), style="BodyText"))
    # note with multiple paragraphs, bold/italic, and a hyperlink
    n = len(doc.fn) + 1
    rich = (doc.note_para("footnote", "A longer note.") .replace("</w:p>", "") + r(" With ", i=True) + r("bold", b=True) + r(" text and a link: ") + doc.link("https://example.org/ledger", "example.org/ledger") + "</w:p>"
            + p(r("Second paragraph inside the same footnote."), style="FootnoteText"))
    doc.fn.append((n, rich))
    out.append(p(r("Rich footnote here.") + doc.note_ref("footnote", n) + r(" And an endnote right after it.") + doc.endnote("Endnotes are collected at the end of the document.") + doc.endnote("A second endnote with a different tone."), style="BodyText"))
    out.append(tbl([tr([tc(p(r("Cell with note") + doc.footnote("Footnote anchored inside a table cell."), style="Normal"), w=4000), tc("Plain cell", w=4000)])], [4000, 4000]))
    out.append(p(""))
    for i in range(6):
        out.append(p(r(paragraph_text(300 + i, 2) + " ") + doc.footnote(f"Supporting remark number {i + 1}: {paragraph_text(900 + i, 1)}") + r(" Closing words of paragraph.") + (doc.endnote(f"Endnote {i + 1}: {paragraph_text(950 + i, 1)}") if i % 2 == 0 else ""), style="BodyText"))
    return "".join(out)


# ================================================================ tables
def tables(doc: Doc):
    out = [H("Tables", 1)]
    # --- nested
    inner2 = tbl([tr([tc("deep A", w=1200), tc("deep B", w=1200)])], [1200, 1200])
    inner = tbl([tr([tc("n1", w=1500), tc("n2", w=1500, extra="")]), tr([tc(p("nest-3"), w=1500), tc(p("level 2 cell") + inner2, w=1500)])], [1500, 1500])
    outer = tbl([
        tr([tc("Outer 1", w=2400, shade="DDEBF7"), tc("Outer 2", w=2400, shade="DDEBF7"), tc("Outer 3", w=2400, shade="DDEBF7")], header=True),
        tr([tc("a", w=2400), tc(p("Contains a nested table:") + inner, w=2400), tc("c", w=2400)]),
        tr([tc("d", w=2400), tc("e", w=2400), tc("f", w=2400)]),
    ], [2400, 2400, 2400])
    out += [p("Nested tables (outer, inner, innermost):", style="BodyText"), outer, p("")]

    # --- merged cells, both directions
    G = [1500] * 5
    merged = tbl([
        tr([tc(p(r("Quarter schedule", b=True)), w=7500, span=5, shade="1F3864", extra=""), ], header=True),
        tr([tc("Berth", w=1500, vmerge="restart", shade="D9E2F3"), tc("Week 1-2", w=3000, span=2, shade="D9E2F3"), tc("Week 3-4", w=3000, span=2, shade="D9E2F3")], header=True),
        tr([tc("", w=1500, vmerge="cont"), tc("In", w=1500), tc("Out", w=1500), tc("In", w=1500), tc("Out", w=1500)], header=True),
        tr([tc("North", w=1500, vmerge="restart"), tc("12", w=1500), tc("9", w=1500), tc("14", w=1500), tc("11", w=1500)]),
        tr([tc("", w=1500, vmerge="cont"), tc("note", w=3000, span=2, vmerge=None), tc("n/a", w=3000, span=2)]),
        tr([tc("East", w=1500, vmerge="restart"), tc("7", w=1500, vmerge="restart"), tc("3", w=1500), tc("8", w=1500), tc("2", w=1500)]),
        tr([tc("", w=1500, vmerge="cont"), tc("", w=1500, vmerge="cont"), tc("4", w=1500), tc("5", w=1500), tc("6", w=1500)]),
        tr([tc(p(r("Total", b=True)), w=1500), tc(p(r("all berths combined"), ), w=6000, span=4, shade="FFF2CC")]),
    ], G)
    out += [p("Merged cells (horizontal and vertical, with a repeating three-row header):", style="BodyText"), merged, p("")]

    # --- borders & shading, per-cell borders
    bd_double = border("top", "double", 6, "C00000") + border("left", "dashed", 8, "2E75B6") + border("bottom", "thick", 12, "548235") + border("right", "dotted", 8, "7030A0")
    styled = tbl([
        tr([tc(p(r("Style", b=True)), w=2400, shade="FFC000", borders=bd_double), tc(p(r("Fill", b=True)), w=2400, shade="70AD47"), tc(p(r("Notes", b=True)), w=2400, shade="5B9BD5")]),
        tr([tc("Double / dashed / thick / dotted", w=2400, borders=bd_double), tc("Solid green", w=2400, shade="E2EFDA"), tc("Light blue", w=2400, shade="DDEBF7", valign="center")]),
        tr([tc("No borders", w=2400, borders=border("top", "nil") + border("left", "nil") + border("bottom", "nil") + border("right", "nil")), tc("Grey", w=2400, shade="D9D9D9"), tc("Bottom aligned", w=2400, valign="bottom")], height=800),
    ], [2400, 2400, 2400], style="TableGrid",
        borders=border("top", "single", 12, "1F3864") + border("left", "single", 12, "1F3864") + border("bottom", "single", 12, "1F3864") + border("right", "single", 12, "1F3864") + border("insideH", "single", 4, "8EAADB") + border("insideV", "single", 4, "8EAADB"))
    out += [p("Borders and shading:", style="BodyText"), styled, p("")]

    # --- banded style, long table with repeating header over page breaks
    rows = [tr([tc(p(r(h)), w=w) for h, w in (("Date", 1800), ("Vessel", 2800), ("Cargo", 2400), ("Tonnes", 1200))], header=True, cant_split=True)]
    vessels = ["Marguerite", "Thornback", "Ostrich Bay", "Calloway", "Bright Venture", "Lantern Moth"]
    cargo = ["grain", "timber", "salt", "coal", "fish", "ballast stone"]
    for i in range(34):
        rows.append(tr([tc(f"2026-02-{1 + i % 28:02d}", w=1800), tc(vessels[i % 6] + f" {i + 1}", w=2800), tc(cargo[(i * 5) % 6], w=2400), tc(p(str(120 + (i * 37) % 900), ppr='<w:jc w:val="right"/>'), w=1200)], cant_split=True))
    out += [p("A banded table style with a repeating header row across pages:", style="BodyText"), tbl(rows, [1800, 2800, 2400, 1200], style="BandedBlue", look="04A0"), p("")]

    # --- autofit / percentage widths / direction / cell margins
    auto = tbl([
        tr([tc("autofit col 1"), tc("autofit column two is wider"), tc("c3")]),
        tr([tc("x"), tc("y"), tc("z")]),
    ], [2000, 3000, 2000], layout="autofit", jc="center", cellmar='<w:top w:w="80" w:type="dxa"/><w:left w:w="200" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="200" w:type="dxa"/>')
    pct = (f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="autofit"/><w:tblLook w:val="04A0"/></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="6000"/></w:tblGrid>'
           '<w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="pct"/></w:tcPr><w:p><w:r><w:t>30%</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3500" w:type="pct"/></w:tcPr><w:p><w:r><w:t>70% (percentage widths, autofit)</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
    vert = tbl([tr([tc("Rotated", w=900, extra='<w:textDirection w:val="btLr"/>'), tc(p(prose_short(1)), w=7000)], height=1400)], [900, 7000])
    out += [p("Autofit, percentage widths, cell margins and rotated text:", style="BodyText"), auto, p(""), pct, p(""), vert, p("")]
    return "".join(out)


def prose_short(i):
    return paragraph_text(70 + i, 1)


# ============================================================== sections
def sections(doc: Doc, fills=None):
    """Returns (body_xml, final_sectPr). Four sections: portrait with first/even/odd headers, landscape in two
    columns (page numbers restart, lower roman), portrait in three columns, then a closing single-column section."""
    from ooxml import sect, landscape_sect
    hd = lambda t, ppr="": p(r(t) + tab() + r("Harbour Report") + tab() + fld("PAGE", "1"), style="Header", ppr=ppr)
    h_def = doc.header_footer("header", hd("Default header"), "header1")
    h_first = doc.header_footer("header", p(r("FIRST PAGE HEADER: title page only", b=True), style="Header"), "header2")
    h_even = doc.header_footer("header", p(r("Even-page header (left pages)"), style="Header", ppr='<w:jc w:val="left"/>'), "header3")
    f_def = doc.header_footer("footer", p(r("Page ") + fld("PAGE", "1") + r(" of ") + fld("NUMPAGES", "6"), style="Footer", ppr='<w:jc w:val="center"/>'), "footer1")
    f_first = doc.header_footer("footer", p(r("First-page footer"), style="Footer"), "footer2")
    f_even = doc.header_footer("footer", p(fld("PAGE", "1") + tab() + r("Even footer"), style="Footer"), "footer3")
    h_land = doc.header_footer("header", p(r("Landscape section header") + tab() + tab() + fld("PAGE", "1", i=True), style="Header"), "header4")
    ref = lambda kind, t, rid: f'<w:{kind}Reference w:type="{t}" r:id="{rid}"/>'
    refs1 = ref("header", "default", h_def) + ref("header", "first", h_first) + ref("header", "even", h_even) + ref("footer", "default", f_def) + ref("footer", "first", f_first) + ref("footer", "even", f_even)
    doc.settings_mid += "<w:evenAndOddHeaders/>"
    fill = lambda i: fills[i]() if fills and i < len(fills) and fills[i] else ""
    out = []
    out.append(p(r("Sections and headers", b=True, sz=40), style="Title"))
    out.append(p(r("Section 1 is portrait with a different first page and different odd/even headers."), style="Subtitle"))
    for i in range(7):
        out.append(body_p(120 + i, 6))
    out.append(fill(0))
    out.append(p("", ppr=sect(refs=refs1, titlepg=True)))  # end of section 1
    # section 2: landscape, 2 columns with separator, page numbers restart lowerRoman
    out.append(H("Landscape in two columns", 1))
    for i in range(8):
        out.append(body_p(140 + i, 5))
    out.append(fill(1))
    out.append(p(r("Column break follows.") + br("column") + r("This text starts the second column."), style="BodyText"))
    out.append(p("", ppr=landscape_sect(refs=ref("header", "default", h_land) + ref("footer", "default", f_def), cols=(2, 540, True), pgnum=("lowerRoman", 1))))
    # section 3: portrait, 3 columns, restarts decimal at 10
    out.append(H("Portrait in three columns", 1))
    for i in range(9):
        out.append(body_p(160 + i, 3))
    out.append(fill(2))
    out.append(p("", ppr=sect(refs=refs1, cols=(3, 360, False), pgnum=(None, 10), kind="nextPage")))
    # section 4 (final): continuous, single column, A4
    out.append(H("Closing section (A4, single column)", 1))
    for i in range(3):
        out.append(body_p(180 + i, 4))
    out.append(fill(3))
    final = sect(w=11906, h=16838, refs=ref("header", "default", h_def) + ref("footer", "default", f_def), pgnum=(None, 1), kind="nextPage")
    return "".join(out), final
