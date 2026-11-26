"""The kitchen-sink deck (many constructs together) and a macro-enabled deck with an inert macro part."""
import assets as A
import build_charts as C
import build_flow as F
import build_tables as T
import pptxbuilder as X
from dml import body_ph, bullet_char, bullet_num, cxn, frame, grad, grp, line, para, pic, rgb, run, scheme, solid, sp, tabs, textbox, title
from pptxbuilder import Master


def kitchen_sink(path):
    theme = X.theme_xml(name="Kitchen sink", major="Georgia", minor="Verdana", ea=("MS Mincho", "MS Gothic"), effects="shadow")
    d = X.Deck(title="Kitchen sink", masters=[Master("Kitchen sink master", theme=theme)], notes_master=True, handout_master=True)
    d.comment_authors = [(0, "Reviewer One", "R1")]
    d.parts_override["ppt/tableStyles.xml"] = T.TABLE_STYLES.encode()
    s = d.slide(0)
    s.add(X.ph_shape(2, "Title 1", ("ctrTitle", None, ""), None, run("Kitchen sink")))
    s.add(X.ph_shape(3, "Subtitle 2", ("subTitle", 1, ""), None, run("Text, shapes, tables, charts, pictures, groups and flow")))
    s.transition = '<p:transition spd="med"><p:fade/></p:transition>'
    s = d.slide(1)
    s.add(title(2, "Text"))
    ext = s.link("https://example.com/spindle-fixture")
    s.add(body_ph(3, para("Level one", ppr_children=bullet_char()) + para("Level two", lvl=1, ppr_children=bullet_char("&#8211;")) + para("Numbered", ppr_children=bullet_num())
                  + para(run("A link", hlink=f'<a:hlinkClick r:id="{ext}"/>') + run(" with ") + run("bold", b=1) + run(" and ") + run("italic", i=1))
                  + para("Name\tQty", ppr_children=tabs((2400000, "l")))))
    s.notes = para(run("Notes: ", b=1) + run("text slide")) + para("Second line", marL=342900, indent=-342900, ppr_children=bullet_char())
    s = d.slide(3)
    s.add(title(2, "Shapes"))
    s.add(sp(3, "Gradient", 700000, 1500000, 2200000, 1200000, prst="roundRect", fill=grad([(0, rgb("FFFFFF")), (100000, rgb("4472C4"))]), ln=line(19050, dash="dash"), paras=para("gradient", algn="ctr")))
    s.add(sp(4, "Shadow", 3300000, 1500000, 2200000, 1200000, prst="star5", fill=solid(scheme("accent2")),
             effects='<a:effectLst><a:outerShdw blurRad="63500" dist="76200" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="45000"/></a:srgbClr></a:outerShdw></a:effectLst>'))
    s.add(sp(5, "Rotated", 5900000, 1500000, 2200000, 1200000, prst="rightArrow", rot=1200000, flipH=True, fill=solid(scheme("accent3", ("lumMod", 75000))), paras=para("rotated", algn="ctr")))
    s.add(sp(6, "Custom", 8500000, 1500000, 1800000, 1200000, custom='<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="50" y="0"/></a:lnTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>', fill=solid(scheme("accent4"))))
    s = d.slide(3)
    s.add(title(2, "Table and picture"))
    rows = [[T.cell("Spans two columns", span=' gridSpan="2"'), T.cell(span=' hMerge="1"'), T.cell("C1")], [T.cell("A2"), T.cell("B2", fill=solid(rgb("FFF2CC")), borders=T.border("B")), T.cell("C2")]]
    s.add(frame(3, "Table", 700000, 1500000, 5400000, 840000, T.table([1800000] * 3, rows, tbl_pr=f'<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{T.STYLE_ID}</a:tableStyleId></a:tblPr>')))
    s.add(pic(4, "Picture", s.media("picture.png", X.png(32, 32)), 6800000, 1500000, 1800000, 1400000, crop='<a:srcRect l="10000" r="10000"/>', descr="A small picture"))
    s = d.slide(3)
    s.add(title(2, "Chart"))
    C.add_chart(d, s, "bar", 3, 1500000, 1500000, 8000000, 4300000, "Chart with a workbook")
    s = d.slide(3)
    s.add(title(2, "Groups and connectors"))
    box = lambda i, name, x, t: sp(i, name, x, 1700000, 1600000, 800000, fill=solid(scheme("accent1")), paras=para(t, algn="ctr"))
    s.add(grp(10, "Group", 700000, 1700000, 3600000, 800000, box(11, "A", 700000, "A") + box(12, "B", 2700000, "B")))
    s.add(sp(13, "Target", 6000000, 3200000, 1600000, 800000, fill=solid(scheme("accent2")), paras=para("target", algn="ctr")))
    s.add(cxn(14, "Elbow", 4300000, 2100000, 1700000, 1500000, "bentConnector3", line(25400, tail=("triangle", "med")), st=(12, 3), end=(13, 1)))
    s = d.slide(3)
    s.hidden = True
    s.add(title(2, "Hidden slide with animation"))
    s.add(sp(3, "Animated", 700000, 1700000, 2200000, 900000, fill=solid(scheme("accent1")), paras=para("fade in", algn="ctr")))
    s.transition = '<p:transition spd="slow" advTm="3000"><p:push dir="u"/></p:transition>'
    s.timing = F.timing([("entrance", 3)], builds='<p:bldP spid="3" grpId="0" animBg="1"/>')
    s = d.slide(3)
    s.add(title(2, "Right-to-left and CJK"))
    s.add(sp(3, "Hebrew", 700000, 1500000, 5000000, 600000, paras=para(run("שלום עולם", lang="he-IL"), algn="r", rtl=1), txbox=True, fill="<a:noFill/>"))
    s.add(sp(4, "Japanese", 700000, 2300000, 5000000, 600000, paras=para(run("こんにちは世界", lang="ja-JP", ea="MS Gothic")), txbox=True, fill="<a:noFill/>"))
    d.slide_comments = {0: [(0, 1000, 800, "Check the subtitle.")]}
    d.pres_extra["ext"] = (f'<p:extLst><p:ext uri="{{521415D9-36F7-43E2-AB2F-B90AF26B5E84}}"><p14:sectionLst xmlns:p14="{F.P14}"><p14:section name="Content" id="{{00000000-0000-4000-8000-000000000061}}"><p14:sldIdLst>'
                           + "".join(f'<p14:sldId id="{i}"/>' for i in range(256, 260)) + '</p14:sldIdLst></p14:section><p14:section name="Backup" id="{00000000-0000-4000-8000-000000000062}"><p14:sldIdLst>'
                           + "".join(f'<p14:sldId id="{i}"/>' for i in range(260, 264)) + "</p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst>")
    d.write(path)


def macro_enabled(path):
    d = X.Deck(title="Macro-enabled presentation", macro=True)
    s = d.slide(3)
    s.add(title(2, "Macro-enabled deck"))
    s.add(sp(3, "Macro button", 700000, 1700000, 3000000, 800000, prst="roundRect", fill=solid(scheme("accent2")), paras=para("Run a macro (never run)", algn="ctr"),
             hlink='<a:hlinkClick r:id="" action="ppaction://macro?name=Module1.Hello"/>'))
    d.extra["ppt/vbaProject.bin"] = (A.opaque_blob("a VBA project"), "application/vnd.ms-office.vbaProject")
    d.pres_rels_extra.append(("http://schemas.microsoft.com/office/2006/relationships/vbaProject", "vbaProject.bin"))
    d.write(path)
