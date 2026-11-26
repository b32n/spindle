"""Text constructs: basic and rich paragraphs, runs, bullets, tab stops, columns, vertical text and autofit."""
import pptxbuilder as X
from dml import (blip_fill, body_ph, body_pr, bullet_char, bullet_num, line, para, rgb, run, scheme, solid, spacing, sp, tabs, textbox, title, tx_body)

EMU = 914400


def _title_slide(d, text, sub):
    s = d.slide(0)
    s.add(X.ph_shape(2, "Title 1", ("ctrTitle", None, ""), None, run(text)))
    s.add(X.ph_shape(3, "Subtitle 2", ("subTitle", 1, ""), None, run(sub)))
    return s


def text_basic(path):
    d = X.Deck(title="Basic text")
    _title_slide(d, "Basic text", "Levels, bullets, numbering, links and spacing")
    s = d.slide(1)
    s.add(title(2, "Levels and bullets"))
    lvls = "".join(para(f"Level {n + 1} paragraph", lvl=n if n else None, ppr_children=bullet_char("&#8226;" if n % 2 == 0 else "&#8211;")) for n in range(5))
    s.add(body_ph(3, lvls))
    s2 = d.slide(1)
    s2.add(title(2, "Numbering and links"))
    ext, mail = s2.link("https://example.com/spindle-fixture"), s2.link("mailto:fixture@example.com")
    nums = (para("First step", marL=457200, indent=-342900, ppr_children=bullet_num())
            + para("Second step", marL=457200, indent=-342900, ppr_children=bullet_num())
            + para("Roman numerals start at three", marL=457200, indent=-342900, ppr_children=bullet_num("romanLcPeriod", 3))
            + para(run("A web link", hlink=f'<a:hlinkClick r:id="{ext}"/>') + run(" and ") + run("a mail link", hlink=f'<a:hlinkClick r:id="{mail}"/>')))
    s2.add(body_ph(3, nums))
    s3 = d.slide(3)
    s3.add(title(2, "Runs, alignment and spacing"))
    runs = (run("Bold ", b=1) + run("italic ", i=1) + run("underlined ", u="sng") + run("struck ", strike="sngStrike") + run("large ", sz=3200)
            + run("coloured ", fill=solid(rgb("C00000"))) + run("themed ", fill=solid(scheme("accent2"))) + run("serif", font="Times New Roman")
            + run(" highlighted", highlight=rgb("FFFF00")))
    s3.add(sp(4, "Runs", 685800, 1500000, 10800000, 700000, paras=para(runs), fill="<a:noFill/>", txbox=True))
    al = "".join(para(f"Aligned {a}", algn=a) for a in ("l", "ctr", "r", "just"))
    s3.add(sp(5, "Alignment", 685800, 2300000, 5200000, 1700000, paras=al, ln=line(9525), bpr=body_pr(wrap="square", lIns=182880, tIns=91440, rIns=182880, bIns=91440, rtlCol="0", anchor="t")))
    spc = (para("Spacing before 12pt, after 6pt", ppr_children=spacing(before=1200, after=600))
           + para("Line spacing 150 percent. The paragraph is long enough to wrap onto a second line so the effect is visible.", ppr_children=spacing(line_pct=150000))
           + para("Exactly 30 points between lines. The paragraph is long enough to wrap onto a second line.", ppr_children=spacing(line_pts=3000)))
    s3.add(sp(6, "Spacing", 6200000, 2300000, 5200000, 2800000, paras=spc, ln=line(9525)))
    s3.add(textbox(7, "Fields", 685800, 5400000, 8000000, 600000,
                   para('<a:fld id="{B2000000-0000-4000-8000-000000000001}" type="slidenum"><a:rPr lang="en-US"/><a:t>3</a:t></a:fld>' + run(" of the deck, on ")
                        + '<a:fld id="{B2000000-0000-4000-8000-000000000002}" type="datetime1"><a:rPr lang="en-US"/><a:t>1/15/2026</a:t></a:fld>')))
    s3.add(X.ph_shape(8, "Footer Placeholder 7", ("ftr", 11, ' sz="quarter"'), None, run("Fixture footer")))
    s3.add(X.ph_shape(9, "Slide Number Placeholder 8", ("sldNum", 12, ' sz="quarter"'), None,
                      '<a:fld id="{B2000000-0000-4000-8000-000000000003}" type="slidenum"><a:rPr lang="en-US"/><a:t>3</a:t></a:fld>'))
    d.write(path)


def text_rich(path):
    d = X.Deck(title="Rich text")
    s = d.slide(3)
    s.add(title(2, "Tabs, columns and vertical text"))
    tab = para("Name\tQuantity\tPrice", ppr_children=tabs((1828800, "l"), (3657600, "ctr"), (5486400, "r")))
    tab += para("Pencil\t12\t0.35", ppr_children=tabs((1828800, "l"), (3657600, "ctr"), (5486400, "r")))
    s.add(sp(3, "Tab stops", 685800, 1500000, 6400000, 900000, paras=tab, ln=line(9525)))
    cols = "".join(para(f"Column text {n}: the quick brown fox jumps over the lazy dog.") for n in range(1, 7))
    s.add(sp(4, "Two columns", 685800, 2600000, 6400000, 1500000, paras=cols, ln=line(9525), bpr=body_pr(numCol=2, spcCol=274320, rtlCol="0")))
    for n, (mode, label) in enumerate((("vert", "Rotated 90"), ("vert270", "Rotated 270"), ("eaVert", "East Asian vertical"), ("wordArtVert", "Stacked")), 5):
        s.add(sp(n, label, 7400000 + (n - 5) * 1000000, 1500000, 900000, 2800000, paras=para(label), ln=line(9525), bpr=body_pr(vert=mode, wrap="square", rtlCol="0", anchor="ctr")))
    s2 = d.slide(3)
    s2.add(title(2, "Autofit, scripts and bullets"))
    long_text = "".join(para(f"Shrink on overflow, line {n}: this text is deliberately too long for its box.", sz=None) for n in range(1, 9))
    s2.add(sp(3, "Shrink text", 685800, 1500000, 3600000, 1200000, paras=long_text, ln=line(9525),
              bpr=body_pr(wrap="square", rtlCol="0", kids='<a:normAutofit fontScale="62500" lnSpcReduction="20000"/>')))
    s2.add(sp(4, "Resize shape", 4600000, 1500000, 3000000, 600000, paras=para("Shape grows to fit its text, however long it gets."), ln=line(9525),
              bpr=body_pr(wrap="square", rtlCol="0", kids="<a:spAutoFit/>")))
    sup = (run("H") + run("2", baseline="-25000") + run("O and x") + run("2", baseline="30000") + run(" ") + run("all caps", cap="all") + run(" ")
           + run("small caps", cap="small") + run(" ") + run("spaced", spc=300) + run(" ") + run("tight", spc=-100))
    s2.add(textbox(5, "Baseline", 4600000, 2300000, 6000000, 500000, para(sup)))
    rid = s2.media("bullet.png", X.png(16, 16))
    bl = (para("Custom character bullet", marL=342900, indent=-342900, ppr_children='<a:buClr>' + rgb("C00000") + '</a:buClr><a:buSzPct val="120000"/>' + bullet_char("&#9733;", "Arial"))
          + para("Symbol font bullet", marL=342900, indent=-342900, ppr_children=bullet_char("&#167;", "Wingdings"))
          + para("Picture bullet", marL=342900, indent=-342900, ppr_children=f'<a:buBlip><a:blip r:embed="{rid}"/></a:buBlip>')
          + para("Numbered, coloured, sized", marL=342900, indent=-342900, ppr_children='<a:buClr>' + rgb("0070C0") + '</a:buClr><a:buSzPts val="1800"/>' + bullet_num("alphaUcParenR")))
    s2.add(sp(6, "Bullets", 685800, 3200000, 6400000, 1800000, paras=bl, ln=line(9525)))
    s2.add(sp(7, "No wrap", 7400000, 3200000, 1600000, 600000, paras=para("A line that never wraps at the edge"), ln=line(9525), bpr=body_pr(wrap="none", rtlCol="0", kids="<a:noAutofit/>")))
    d.write(path)
