"""revisions-authors.docx: tracked insertions / deletions / moves / formatting changes by several authors,
plus comments with threaded replies (commentsExtended), resolved threads and people.xml.

Producer: hand-authored OOXML via ooxml.py (Python zipfile).
"""
import sys

sys.path.insert(0, __file__.rsplit("/", 1)[0])
from ooxml import *  # noqa: E402,F401,F403
from ooxml import NS_DECL, CT, REL, esc, p, r, tbl, tr, tc, Package, rpr, prose, num  # noqa: E402

AUTH = {"A": ("Alice Archer", "AA"), "B": ("Bram Keller", "BK"), "C": ("Chidi Okafor", "CO"), "D": ("Dana Whitfield", "DW")}
_id = [100]


def nid():
    _id[0] += 1
    return _id[0]


def date(day, hour=10):
    return f"2026-01-{day:02d}T{hour:02d}:00:00Z"


def ins(who, text, day=5, **kw):
    return f'<w:ins w:id="{nid()}" w:author="{AUTH[who][0]}" w:date="{date(day)}">{r(text, **kw)}</w:ins>'


def dele(who, text, day=5, **kw):
    rp = rpr(**kw)
    return (f'<w:del w:id="{nid()}" w:author="{AUTH[who][0]}" w:date="{date(day)}"><w:r>{rp}'
            f'<w:delText xml:space="preserve">{esc(text)}</w:delText></w:r></w:del>')


def fmt_change(who, text, new, old="", day=6):
    """Run whose formatting changed: new rPr + rPrChange recording the old (empty) properties."""
    inner = f'<w:rPrChange w:id="{nid()}" w:author="{AUTH[who][0]}" w:date="{date(day)}"><w:rPr>{old}</w:rPr></w:rPrChange>'
    return f'<w:r><w:rPr>{new}{inner}</w:rPr><w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def move(who, name, mid_from, mid_to, text, day=7):
    """Returns (moveFrom xml, moveTo xml) for the same text."""
    a, d = AUTH[who][0], date(day)
    src = (f'<w:moveFromRangeStart w:id="{mid_from}" w:name="{name}" w:author="{a}" w:date="{d}"/>'
           f'<w:moveFrom w:id="{nid()}" w:author="{a}" w:date="{d}"><w:r><w:delText xml:space="preserve">{esc(text)}</w:delText></w:r></w:moveFrom>'
           f'<w:moveFromRangeEnd w:id="{mid_from}"/>')
    dst = (f'<w:moveToRangeStart w:id="{mid_to}" w:name="{name}" w:author="{a}" w:date="{d}"/>'
           f'<w:moveTo w:id="{nid()}" w:author="{a}" w:date="{d}">{r(text)}</w:moveTo>'
           f'<w:moveToRangeEnd w:id="{mid_to}"/>')
    return src, dst


# -- comments ------------------------------------------------------------------
COMMENTS = []  # (cid, who, day, paraId, text, parentParaId|None, done)


def comment(cid, who, day, text, parent=None, done=False):
    para = f"{0x10000000 + cid:08X}"
    COMMENTS.append((cid, who, day, para, text, parent, done))
    return para


def crange(cid_list, inner):
    """Wrap inner in the (shared) range of a thread: root + replies."""
    s = "".join(f'<w:commentRangeStart w:id="{c}"/>' for c in cid_list)
    e = "".join(f'<w:commentRangeEnd w:id="{c}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{c}"/></w:r>' for c in cid_list)
    return s + inner + e


def comments_xml():
    out = []
    for cid, who, day, para, text, parent, _ in COMMENTS:
        name, ini = AUTH[who]
        out.append(
            f'<w:comment w:id="{cid}" w:author="{name}" w:date="{date(day, 14)}" w:initials="{ini}">'
            f'<w:p w14:paraId="{para}" w14:textId="77777777"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>'
            f'<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>{r(text)}</w:p></w:comment>'
        )
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments {NS_DECL}>{"".join(out)}</w:comments>'


def comments_ext_xml():
    out = []
    for cid, who, day, para, text, parent, done in COMMENTS:
        pp = f' w15:paraIdParent="{parent}"' if parent else ""
        out.append(f'<w15:commentEx w15:paraId="{para}"{pp} w15:done="{1 if done else 0}"/>')
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx {NS_DECL}>{"".join(out)}</w15:commentsEx>'


def people_xml():
    pe = "".join(f'<w15:person w15:author="{n}"><w15:presenceInfo w15:providerId="None" w15:userId="{n}"/></w15:person>' for n, _ in AUTH.values())
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:people {NS_DECL}>{pe}</w15:people>'


def build_body():
    B = []
    B.append(p("Revision review: quarterly harbour report", style="Title"))
    B.append(p("Four reviewers marked up this draft; tracked changes and comments are preserved.", style="Subtitle"))

    # --- 1. insertions & deletions by different authors
    B.append(p("1. Insertions and deletions", style="Heading1"))
    # paragraph with comment thread: root by Alice, reply by Bram, reply to reply by Chidi (flat, as Word does), resolved
    c1 = comment(1, "A", 8, "Is the 2025 figure final? Please confirm with the harbourmaster.")
    comment(2, "B", 9, "Confirmed by phone on Monday, the number stands.", parent=c1)
    comment(3, "C", 10, "Thanks, I will cite the call in the footer.", parent=c1, done=True)
    B.append(p(
        r("Total tonnage handled in ")
        + dele("A", "2024", 5)
        + ins("A", "2025", 5)
        + r(" rose ")
        + crange([1, 2, 3], r("by roughly "))
        + dele("B", "eight", 6)
        + ins("B", "eleven", 6)
        + r(" percent, ")
        + ins("C", "mostly because of the new grain terminal, ", 6, i=True)
        + r("while passenger traffic was flat."), style="BodyText"))
    B.append(p(
        r(prose(1, 3) + " ")
        + dele("C", "This sentence was judged redundant and removed by the third reviewer. ", 7)
        + ins("D", "A replacement sentence was supplied by the fourth reviewer, with a note about winter dredging. ", 7, b=True)
        + r(prose(1, 4)), style="BodyText"))

    # inserted paragraph (paragraph mark tracked) and a deleted whole paragraph
    ins_mark = f'<w:rPr><w:ins w:id="{nid()}" w:author="{AUTH["D"][0]}" w:date="{date(7)}"/></w:rPr>'
    B.append(p(ins("D", "Entire paragraph inserted by Dana: berth fees will be reviewed in spring.", 7), style="BodyText", ppr=ins_mark))
    del_mark = f'<w:rPr><w:del w:id="{nid()}" w:author="{AUTH["B"][0]}" w:date="{date(7)}"/></w:rPr>'
    B.append(p(dele("B", "Entire paragraph deleted by Bram: the old fee table is obsolete.", 7), style="BodyText", ppr=del_mark))

    # --- 2. moves
    B.append(p("2. Moved text", style="Heading1"))
    src, dst = move("C", "move-1", 201, 202, "The dredging schedule is published every October.", 7)
    B.append(p(r("Before the move: ") + src + r(" (that sentence used to be here)."), style="BodyText"))
    B.append(p(r(prose(1, 6) + " ") + dst + r(" It now lives in this paragraph."), style="BodyText"))
    src2, dst2 = move("A", "move-2", 203, 204, "Tide tables are posted at the north gate.", 8)
    B.append(p(src2 + r(" "), style="BodyText"))
    B.append(p(r("Second move target: ") + dst2, style="BodyText"))

    # --- 3. formatting changes
    B.append(p("3. Formatting changes", style="Heading1"))
    c4 = comment(4, "D", 9, "Bold here is for the board summary; please keep.")
    B.append(p(
        r("Run formatting: ")
        + crange([4], fmt_change("D", "now bold", "<w:b/>"))
        + r(", ")
        + fmt_change("B", "now italic and red", '<w:i/><w:color w:val="C00000"/>', old="<w:b/>")
        + r(", ")
        + fmt_change("C", "larger", '<w:sz w:val="28"/><w:szCs w:val="28"/>', day=8)
        + r(" and plain again."), style="BodyText"))
    ppr_change = (f'<w:jc w:val="center"/><w:pPrChange w:id="{nid()}" w:author="{AUTH["A"][0]}" w:date="{date(8)}"><w:pPr><w:pStyle w:val="BodyText"/></w:pPr></w:pPrChange>')
    B.append(p(r("This paragraph was centred by Alice (paragraph property change)."), style="BodyText", ppr=ppr_change))
    style_change = (f'<w:pStyle w:val="Heading2"/><w:pPrChange w:id="{nid()}" w:author="{AUTH["C"][0]}" w:date="{date(8)}"><w:pPr><w:pStyle w:val="BodyText"/></w:pPr></w:pPrChange>')
    B.append(f'<w:p><w:pPr>{style_change}</w:pPr>{r("Promoted from body text to a heading by Chidi")}</w:p>')

    # --- 4. tracked list + table changes
    B.append(p("4. Lists and tables", style="Heading1"))
    c5 = comment(5, "B", 10, "Should the dredging item be first?")
    comment(6, "A", 10, "Yes; schedule risk is highest there.", parent=c5)
    B.append(p(crange([5, 6], ins("A", "Order dredging equipment", 9)), style="ListParagraph", ppr=num(1)))
    B.append(p(r("Inspect the sea wall"), style="ListParagraph", ppr=num(1)))
    B.append(p(dele("C", "Repaint the buoys", 9), style="ListParagraph", ppr=num(1)))
    B.append(p(r("Publish the report"), style="ListParagraph", ppr=num(1)))

    def cell(text, w, tracked=None, who="D"):
        if tracked == "ins":
            return tc(p(ins(who, text, 9)), w=w)
        if tracked == "del":
            return tc(p(dele(who, text, 9)), w=w)
        return tc(p(r(text)), w=w)

    ins_row = f'<w:trPr><w:ins w:id="{nid()}" w:author="{AUTH["D"][0]}" w:date="{date(9)}"/></w:trPr>'
    del_row = f'<w:trPr><w:del w:id="{nid()}" w:author="{AUTH["B"][0]}" w:date="{date(9)}"/></w:trPr>'
    rows = [
        tr([cell("Berth", 3000), cell("Vessels", 3000), cell("Fee", 3000)], header=True),
        tr([cell("North", 3000), cell("12", 3000), cell("40", 3000)]),
        f'<w:tr>{ins_row}{cell("East (new)", 3000, "ins")}{cell("7", 3000, "ins")}{cell("55", 3000, "ins")}</w:tr>',
        f'<w:tr>{del_row}{cell("West (closed)", 3000, "del", "B")}{cell("0", 3000, "del", "B")}{cell("0", 3000, "del", "B")}</w:tr>',
        tr([cell("South", 3000), tc(p(dele("A", "9", 9) + ins("A", "10", 9)), w=3000), cell("45", 3000)]),
    ]
    B.append(tbl(rows, [3000, 3000, 3000]))
    B.append(p("", style="BodyText"))

    # --- 5. misc: tracked hyperlink-ish text, field inside insertion, symbols
    B.append(p("5. Mixed revisions in one paragraph", style="Heading1"))
    B.append(p(
        r("Final paragraph: ")
        + ins("A", "alpha ", 10) + dele("B", "beta ", 10) + ins("C", "gamma ", 10) + dele("D", "delta ", 10)
        + ins("B", "epsilon ", 11, u=True) + r("zeta.") , style="BodyText"))
    B.append(p(r("A closing paragraph without markup, so the document ends cleanly. " + prose(2, 9)), style="BodyText"))
    return "".join(B)


def build(out):
    pkg = Package("Revision review", creator="Alice Archer")
    pkg.standard(settings=settings_with_track())
    body = build_body()
    pkg.document(body)
    pkg.part("word/comments.xml", comments_xml(), CT["comments"], "comments")
    pkg.part("word/commentsExtended.xml", comments_ext_xml(), CT["commentsExtended"], "http://schemas.microsoft.com/office/2011/relationships/commentsExtended")
    pkg.part("word/people.xml", people_xml(), "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml", "http://schemas.microsoft.com/office/2011/relationships/people")
    pkg.save(out)


def settings_with_track():
    from ooxml import settings_xml
    return settings_xml(extra="<w:trackRevisions/>", footnote_pr=False)


if __name__ == "__main__":
    build(sys.argv[1])
