"""Programmatic (hand-authored OOXML) fixtures: one function per file. See build_all.py."""
import os

import features as F
import features2 as F2
import features3 as F3
from docbuilder import PEOPLE, Doc, paragraph_text
from ooxml import (bookmark, br, fld, numbering_xml, p, r, sect, tab, tbl, tc, tr, theme_xml)

LIB = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"


def title(doc, t, sub):
    return p(r(t), style="Title") + p(r(sub), style="Subtitle")


def simple(name, ttl, sub, *blocks, final=None, setup=None):
    def build(out):
        doc = Doc(ttl)
        if setup:
            setup(doc)
        body = title(doc, ttl, sub) + "".join(b(doc) if callable(b) else b for b in blocks)
        doc.finish(body, final if final is not None else sect(), out)
    return build


def notes_toc_fields(out):
    doc = Doc("Fields and notes")
    body = title(doc, "Fields, notes and navigation", "TOC, PAGE / PAGEREF / REF / SEQ / DATE, bookmarks, hyperlinks, footnotes and endnotes")
    body += F.toc_fields(doc) + F.notes(doc)
    doc.finish(body, sect(), out)


def sections_headers(out):
    doc = Doc("Sections and headers")
    body, final = F.sections(doc)
    doc.finish(body, final, out)


def typography_extras(out):
    doc = Doc("Typography extras")
    wm = F3.watermark_header(doc)
    plain_footer = doc.header_footer("footer", p(r("Page ") + fld("PAGE", "1"), style="Footer", ppr='<w:jc w:val="center"/>'), "footer2")
    refs = f'<w:headerReference w:type="default" r:id="{wm}"/><w:footerReference w:type="default" r:id="{plain_footer}"/>'
    body = title(doc, "Typography extras", "Hyphenation, tab leaders, drop caps, a watermark, line numbers and page borders") + F3.typography(doc)
    for i in range(6):
        body += F.body_p(500 + i, 6, "BodyText")
    borders = "".join(f'<w:{s} w:val="single" w:sz="8" w:space="24" w:color="2F5496"/>' for s in ("top", "left", "bottom", "right"))
    final = sect(refs=refs, borders=borders, lnnum='<w:lnNumType w:countBy="1" w:restart="newPage"/>')
    doc.finish(body, final, out)


def fonts_themes(out):
    doc = Doc("Fonts and themes")
    palette = ("C0392B", "2980B9", "27AE60", "F39C12", "8E44AD", "16A085")
    doc.theme = theme_xml(major="Liberation Serif", minor="Liberation Sans", name="Harbour Palette", accents=palette)
    F3.embed_font(doc, "Liberation Sans", LIB, "spindle-fixture-font-1")
    body = title(doc, "Embedded font and custom theme", "Liberation Sans (SIL OFL 1.1) is embedded, obfuscated, in word/fonts; the theme overrides colours and fonts")
    body += p(r("The body font below is Liberation Sans through the theme's minor font, and it is also embedded in this file so the document renders the same where the font is not installed."), style="BodyText")
    body += p(r("Explicit font names: ") + r("Liberation Sans (embedded)", font="Liberation Sans") + r(", ") + r("Liberation Serif (not embedded)", font="Liberation Serif") + r(", ") + r("DejaVu Sans Mono (not embedded)", font="DejaVu Sans Mono") + r("."), style="BodyText")
    body += F3.themed(doc, palette)
    for i in range(4):
        body += F.body_p(520 + i, 5)
    doc.finish(body, sect(), out)


def charts_objects(out):
    doc = Doc("Charts and objects")
    body = title(doc, "Charts and embedded objects", "Three charts with embedded workbooks and an OLE worksheet object") + F3.charts(doc)
    body += F.body_p(530, 4)
    doc.finish(body, sect(), out)


# ------------------------------------------------------------ nested combos
def nested_combos(doc: Doc):
    out = [F.H("Combined constructs", 1)]
    c1, p1 = doc.comment(0, 8, "Which table does this footnote belong to?")
    c2, _ = doc.comment(1, 9, "The berth table; I moved it.", parent=p1, done=True)
    # field inside a tracked insertion, inside a hyperlink, with a comment spanning a bookmark
    ins_field = f'<w:ins w:id="{doc.nid()}" w:author="{PEOPLE[1]}" w:date="2026-01-06T10:00:00Z">{fld("DATE \\@ \"d MMM yyyy\"", "15 Jan 2026")}</w:ins>'
    out.append(p(r("A tracked-in date field: ") + ins_field + r(". A tracked link: ") +
                 f'<w:hyperlink r:id="{doc.rel("hyperlink", "https://example.com/tracked", external=True)}" w:history="1">{doc.ins(2, "tracked hyperlink text", 6, style="Hyperlink")}</w:hyperlink>'
                 + r(". A comment over a bookmark: ") + doc.commented([c1, c2], bookmark(801, "nestedbm", r("bookmarked and commented") ) + doc.footnote("Footnote reference inside a comment range.")) + r("."), style="BodyText"))
    # table with tracked row, comment, footnote, move and nested list
    trk = f'<w:trPr><w:ins w:id="{doc.nid()}" w:author="{PEOPLE[3]}" w:date="2026-01-07T10:00:00Z"/></w:trPr>'
    from ooxml import tbl, tc, tr
    mv_from = f'<w:moveFromRangeStart w:id="811" w:name="mv-t" w:author="{PEOPLE[0]}" w:date="2026-01-07T10:00:00Z"/><w:moveFrom w:id="{doc.nid()}" w:author="{PEOPLE[0]}" w:date="2026-01-07T10:00:00Z"><w:r><w:delText>moved cell text</w:delText></w:r></w:moveFrom><w:moveFromRangeEnd w:id="811"/>'
    mv_to = f'<w:moveToRangeStart w:id="812" w:name="mv-t" w:author="{PEOPLE[0]}" w:date="2026-01-07T10:00:00Z"/><w:moveTo w:id="{doc.nid()}" w:author="{PEOPLE[0]}" w:date="2026-01-07T10:00:00Z">{r("moved cell text")}</w:moveTo><w:moveToRangeEnd w:id="812"/>'
    out.append(tbl([
        tr([tc(p(r("Header A", b=True)), w=3000, shade="D9E2F3"), tc(p(r("Header B", b=True)), w=3000, shade="D9E2F3"), tc(p(r("Header C", b=True)), w=3000, shade="D9E2F3")], header=True),
        tr([tc(p(mv_from), w=3000), tc(p(doc.dele(1, "deleted in cell", 6) + doc.ins(1, "inserted in cell", 6)), w=3000), tc(p(r("note") + doc.footnote("Footnote inside a table with revisions.")), w=3000)]),
        f'<w:tr>{trk}' + tc(p(doc.ins(3, "tracked row, cell 1", 7)), w=3000) + tc(p(mv_to), w=3000) + tc(p(doc.ins(3, "tracked row, cell 3", 7) + r(" ") + fld("SEQ Table \\* ARABIC", "2")), w=3000) + "</w:tr>",
    ], [3000, 3000, 3000]))
    out.append(p(""))
    # list with revisions and a deleted footnote reference
    dfn = f'<w:del w:id="{doc.nid()}" w:author="{PEOPLE[2]}" w:date="2026-01-08T10:00:00Z">{doc.note_ref("footnote", 1)}</w:del>'
    out.append(p(r("List item with a deleted footnote reference") + dfn, style="ListParagraph", ppr=F2.num_ppr().replace('numId w:val="2"', 'numId w:val="1"')))
    out.append(p(doc.ins(0, "A tracked list item at level 2", 8), style="ListParagraph", ppr='<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>'))
    out.append(p(r("A formatting change inside a heading-like paragraph: ") + f'<w:r><w:rPr><w:b/><w:rPrChange w:id="{doc.nid()}" w:author="{PEOPLE[2]}" w:date="2026-01-08T10:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>now bold</w:t></w:r>', style="BodyText"))
    return "".join(out)


def kitchen_sink(out):
    doc = Doc("Kitchen sink")
    from ooxml import landscape_sect
    def sec0():
        return F.toc_fields(doc, [(1, "Navigation"), (1, "Lists and styles"), (1, "Review"), (1, "Tables"), (1, "Media"), (1, "Equations"), (1, "Controls"), (1, "Languages"), (1, "Typography"), (1, "Charts")]) + F.notes(doc)
    def sec1():
        lists = F.H("Lists and styles", 1)
        for i, (lvl, t) in enumerate([(0, "Survey"), (1, "North quay"), (2, "Bollards"), (1, "East quay"), (0, "Review"), (1, "Draft"), (1, "Final")]):
            lists += p(r(t), style="ListParagraph", ppr=f'<w:numPr><w:ilvl w:val="{lvl}"/><w:numId w:val="1"/></w:numPr>')
        for t in ("First", "Second"):
            lists += p(r(t), style="ListParagraph", ppr='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr>')
        lists += p(r("Quoted: " + paragraph_text(540, 2)), style="Quote") + p(r(paragraph_text(541, 2)), style="IntenseQuote")
        lists += p(r("def tide(h):\n    return h % 12"), style="CodeBlock")
        return lists + nested_combos(doc)
    def sec2():
        return F.tables(doc) + F2.images(doc)
    def sec3():
        return F2.shapes(doc) + F2.equations(doc) + F2.sdt_forms(doc) + F3.i18n(doc) + F3.typography(doc) + F3.charts(doc)
    body, final = F.sections(doc, [sec0, sec1, sec2, sec3])
    doc.finish(F.H("Kitchen sink", 1) .replace("Heading1", "Title") + body, final, out)


def word_compat_kitchen(out):
    """A report in the style Word writes (rsids, paraIds, proofErr, smartTag, AlternateContent, compat settings, cover page control)."""
    doc = Doc("Quarterly harbour operations report", creator="Harbour Records Office")
    doc.settings_extra = ('<w:zoom w:percent="100"/><w:proofState w:spelling="clean" w:grammar="clean"/>')
    doc.settings_mid = ('<w:characterSpacingControl w:val="doNotCompress"/>'.replace('<w:characterSpacingControl w:val="doNotCompress"/>', '') +
                        '<w:rsids><w:rsidRoot w:val="00A1B2C3"/><w:rsid w:val="00A1B2C3"/><w:rsid w:val="00B2C3D4"/></w:rsids>'
                        '<w:themeFontLang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>'
                        '<w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="2050"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>')
    wp = lambda para, rsid="00A1B2C3": para.replace("<w:p>", f'<w:p w14:paraId="{0x30000000 + doc.nid():08X}" w14:textId="77777777" w:rsidR="{rsid}" w:rsidRDefault="{rsid}">', 1)
    cover = (f'<w:sdt><w:sdtPr><w:id w:val="-1"/><w:docPartObj><w:docPartGallery w:val="Cover Pages"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>'
             + p(r("Quarterly Harbour Operations Report", b=True, sz=56, color="1F3864"), style="Title") + p(r("Q4 2025 — Marguerite Harbour Authority"), style="Subtitle")
             + p(r("Prepared by the Records Office"), style="BodyText") + p("", ppr="") + p(br("page")) + "</w:sdtContent></w:sdt>")
    body = cover
    body += F.toc_fields(doc, [(1, "Executive summary"), (1, "Operations"), (2, "Vessel traffic"), (2, "Cargo"), (1, "Finance"), (1, "Outlook")])
    # Word noise: proofErr, smartTag, lastRenderedPageBreak, bookmark _GoBack, noProof runs
    noisy = (f'<w:p w14:paraId="40000001" w14:textId="77777777" w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3"><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:rPr><w:lang w:val="en-GB"/></w:rPr><w:t xml:space="preserve">Spelling errors are marked by </w:t></w:r>'
             '<w:proofErr w:type="spellStart"/><w:r><w:t>teh</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:t xml:space="preserve"> editor, </w:t></w:r><w:r><w:lastRenderedPageBreak/><w:t xml:space="preserve">page-break markers are cached, </w:t></w:r>'
             '<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place"><w:smartTagPr><w:attr w:name="ProductID" w:val="Marguerite"/></w:smartTagPr><w:r><w:t>Marguerite</w:t></w:r></w:smartTag>'
             '<w:r><w:t xml:space="preserve"> is a smart-tagged place, </w:t></w:r><w:bookmarkStart w:id="9" w:name="_GoBack"/><w:bookmarkEnd w:id="9"/><w:r><w:rPr><w:noProof/></w:rPr><w:t>noProof text.</w:t></w:r></w:p>')
    body += F.H("Executive summary", 1)
    body += wp(F.body_p(600, 5)) + noisy
    cid, par = doc.comment(1, 12, "Please add the Q3 comparison here.")
    doc.comment(0, 13, "Added in the table below.", parent=par)
    body += p(doc.commented([0, 1], r("Throughput rose across all four quarters.")) + r(" ") + doc.link("https://example.com/report", "Full dataset") + doc.footnote("Figures are provisional until the audit closes."), style="BodyText")
    body += F.H("Operations", 1)
    body += F.H("Vessel traffic", 2)
    body += F.tables(doc).split(F.H("Tables", 1))[-1][:20000] if False else ""
    body += F2.images(doc).split(F.H("Images", 1))[-1]
    body += F.H("Cargo", 2) + F3.charts(doc).split(F.H("Charts and embedded objects", 1))[-1]
    body += F.H("Finance", 1) + F.tables(doc).split(F.H("Tables", 1))[-1]
    body += F.H("Outlook", 1) + F.body_p(610, 5) + F2.shapes(doc).split(F.H("Shapes, text boxes and SmartArt", 1))[-1]
    body += F2.sdt_forms(doc).split(F.H("Content controls and form fields", 1))[-1]
    h = doc.header_footer("header", p(r("Marguerite Harbour Authority") + tab() + tab() + fld("STYLEREF \"Heading 1\"", "Outlook"), style="Header"), "header1")
    f = doc.header_footer("footer", p(r("Quarterly Report — page ") + fld("PAGE", "1") + r(" of ") + fld("NUMPAGES", "12"), style="Footer", ppr='<w:jc w:val="center"/>'), "footer1")
    final = sect(refs=f'<w:headerReference w:type="default" r:id="{h}"/><w:footerReference w:type="default" r:id="{f}"/>')
    doc.finish(body, final, out)


# ------------------------------------------------------------ long document
def long_document(out):
    doc = Doc("The Marguerite Harbour Handbook")
    chapters = 12
    body = (p(r("The Marguerite Harbour Handbook", sz=64, b=True), style="Title") + p(r("Operations, records and procedures — a long-form reference"), style="Subtitle")
            + p(r("Compiled by the Records Office"), style="BodyText") + p("", ppr=sect(kind="nextPage")))
    h = doc.header_footer("header", p(r("The Marguerite Harbour Handbook") + tab() + tab() + fld('STYLEREF "Heading 1"', "Preface"), style="Header"), "header1")
    f = doc.header_footer("footer", p(r("Page ") + fld("PAGE", "1") + r(" of ") + fld("NUMPAGES", "70") + tab() + tab() + r("Confidential"), style="Footer"), "footer1")
    refs = f'<w:headerReference w:type="default" r:id="{h}"/><w:footerReference w:type="default" r:id="{f}"/>'
    heads = []
    for c in range(1, chapters + 1):
        heads.append((1, f"Chapter {c}: " + ["Tides and Currents", "Berth Allocation", "Cargo Handling", "Pilotage", "Lighthouse Operations", "Dredging", "Customs and Records", "Winter Procedures", "Fishing Fleet", "Ferry Services", "Finance", "Safety and Emergency"][c - 1]))
        for s in range(1, 7):
            heads.append((2, f"{c}.{s} " + ["Background", "Current practice", "Responsibilities", "Records", "Exceptions", "Review"][s - 1]))
    # Contents
    toc = p("Contents", style="TOCHeading")
    pages = 2
    entries = []
    for i, (lvl, t) in enumerate(heads):
        pages += 1 if lvl == 2 and (i % 3 == 0) else 0
        entries.append((lvl, t, f"_TocL{i}", pages + (i // 7)))
    for k, (lvl, t, bm, pg) in enumerate(entries):
        run = f'<w:hyperlink w:anchor="{bm}" w:history="1">{r(t, style="Hyperlink")}{tab()}{fld(f"PAGEREF {bm} \\h", str(pg))}</w:hyperlink>'
        if k == 0:
            run = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run
        if k == len(entries) - 1:
            run += '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        toc += p(run, style=f"TOC{lvl}")
    body += toc + p("", ppr=sect(refs=refs, kind="nextPage", pgnum=("lowerRoman", 1)))
    # chapters
    ext_num = ""
    hi = 0
    seed = 1000
    tbl_n = 0
    for c in range(1, chapters + 1):
        numid = 20 + c
        ext_num += f'<w:num w:numId="{numid}"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>'
        lvl, t = heads[hi]; bm = f"_TocL{hi}"; hi += 1
        body += p(bookmark(2000 + hi, bm, r(t)), style="Heading1", ppr="<w:pageBreakBefore/>" if False else "")
        body += F.body_p(seed := seed + 1, 5)
        for s in range(1, 7):
            lvl, t = heads[hi]; bm = f"_TocL{hi}"; hi += 1
            body += p(bookmark(2000 + hi, bm, r(t)), style="Heading2")
            for k in range(5):
                para = r(paragraph_text(seed := seed + 1, 6))
                if k == 1:
                    para += doc.footnote(f"Reference note {c}.{s}: {paragraph_text(seed + 5000, 1)}")
                if k == 2 and s == 1:
                    para += r(" See also ") + fld(f"REF {'_TocL' + str(hi - s)} \\h", f"chapter {c}") + r(" and ") + doc.link("https://example.com/handbook/" + str(c), "the online edition") + r(".")
                body += p(para, style="BodyText")
            if s == 2:
                for it in range(5):
                    body += p(r(f"Checklist item {it + 1}: " + paragraph_text(seed + 10 * it, 1)), style="ListParagraph", ppr=F2.num_ppr())
                for it in range(4):
                    body += p(r(f"Step {it + 1}: " + paragraph_text(seed + 20 * it + 3, 1)), style="ListParagraph", ppr=f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{numid}"/></w:numPr>')
                body += p(r(f"Sub-step detail: " + paragraph_text(seed + 99, 1)), style="ListParagraph", ppr=f'<w:numPr><w:ilvl w:val="1"/><w:numId w:val="{numid}"/></w:numPr>')
            if s == 3:
                tbl_n += 1
                rows = [tr([tc(p(r(x, b=True)), w=w, shade="D9E2F3") for x, w in (("Item", 2200), ("Owner", 2400), ("Frequency", 2400), ("Notes", 2600))], header=True)]
                for k in range(7):
                    rows.append(tr([tc(f"Task {c}.{k + 1}", w=2200), tc(PEOPLE[k % 4], w=2400), tc(["daily", "weekly", "monthly", "yearly"][k % 4], w=2400), tc(paragraph_text(seed + k, 1)[:60], w=2600)]))
                body += p(r(f"Table {tbl_n}. ") + r("Standing tasks, chapter " + str(c)) , style="Caption") + tbl(rows, [2200, 2400, 2400, 2600], style="BandedBlue")
                body += p("")
            if s == 4:
                body += p(doc.inline(["rect", "circle", "diag", "stripes", "checker"][c % 5], 200, 110, f"Diagram for chapter {c}", fg=(30 + c * 12, 90, 200 - c * 8)), ppr='<w:jc w:val="center"/>')
                body += p(r("Figure ") + fld("SEQ Figure \\* ARABIC", str(c)) + r(f": diagram for chapter {c}."), style="Caption", ppr='<w:jc w:val="center"/>')
            if s == 5:
                body += p(r(paragraph_text(seed + 7, 2)), style="Quote")
                body += p(r(f"harbour --chapter {c} --verify"), style="CodeBlock")
        body += p("", ppr=sect(refs=refs, kind="nextPage", pgnum=(None, 1) if c == 1 else None))
    # appendix in two columns: glossary table + index-like lists
    body += F.H("Appendix: glossary", 1)
    for k in range(40):
        body += p(r(f"Term {k + 1}. ", b=True) + r(paragraph_text(9000 + k, 2)), style="BodyText")
    final = sect(refs=refs, cols=(2, 540, True), kind="nextPage")
    doc.extra_numbering = ext_num
    doc.finish(body, final, out)
