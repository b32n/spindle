"""Tables (merges, banding, borders, a custom table style), groups and connectors."""
import pptxbuilder as X
from dml import cxn, frame, grp, line, para, rgb, run, scheme, solid, sp, textbox, title

TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table"
STYLE_ID = "{7A1B2C3D-0000-4000-8000-0000000000A1}"
TABLE_STYLES = (
    f'{X.HEAD}<a:tblStyleLst xmlns:a="{X.A}" def="{STYLE_ID}"><a:tblStyle styleId="{STYLE_ID}" styleName="Spindle test style">'
    '<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr>'
    + "".join(f'<a:{side}><a:ln w="12700" cmpd="sng"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln></a:{side}>' for side in ("left", "right", "top", "bottom", "insideH", "insideV"))
    + '</a:tcBdr><a:fill><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl>'
    '<a:band1H><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H>'
    '<a:firstRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/>'
    '<a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstRow></a:tblStyle></a:tblStyleLst>'
)


def cell(text="", span="", fill="", borders="", anchor=None, algn=None):
    """One a:tc. `span` = attributes (gridSpan, rowSpan, hMerge, vMerge); `borders` = lnL/lnR/lnT/lnB XML."""
    txt = para(text, algn=algn) if text else '<a:p><a:endParaRPr lang="en-US"/></a:p>'
    a = f' anchor="{anchor}"' if anchor else ""
    return f'<a:tc{span}><a:txBody><a:bodyPr/><a:lstStyle/>{txt}</a:txBody><a:tcPr{a}>{borders}{fill}</a:tcPr></a:tc>'


def table(cols, rows, tbl_pr='<a:tblPr firstRow="1" bandRow="1"/>', row_h=420000):
    grid = "".join(f'<a:gridCol w="{w}"/>' for w in cols)
    body = "".join(f'<a:tr h="{row_h}">{"".join(r)}</a:tr>' for r in rows)
    return f'<a:graphicData uri="{TABLE_URI}"><a:tbl>{tbl_pr}<a:tblGrid>{grid}</a:tblGrid>{body}</a:tbl></a:graphicData>'


def border(side, color="C00000", w=19050, dash="solid"):
    return f'<a:ln{side} w="{w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>{rgb(color)}</a:solidFill><a:prstDash val="{dash}"/></a:ln{side}>'


def basic_table(cols=4, rows=4):
    head = [cell(f"Column {c}") for c in range(1, cols + 1)]
    return table([1800000] * cols, [head] + [[cell(f"R{r}C{c}") for c in range(1, cols + 1)] for r in range(1, rows)])


def tables(path):
    d = X.Deck(title="Tables")
    d.parts_override["ppt/tableStyles.xml"] = TABLE_STYLES.encode()
    s = d.slide(3)
    s.add(title(2, "Plain table with a style, banding and a header row"))
    s.add(frame(3, "Styled table", 700000, 1500000, 7200000, 1680000, table([1800000] * 4, [[cell(f"Column {c}") for c in range(1, 5)]] + [[cell(f"R{r}C{c}") for c in range(1, 5)] for r in range(1, 4)],
                tbl_pr=f'<a:tblPr firstRow="1" firstCol="1" bandRow="1"><a:tableStyleId>{STYLE_ID}</a:tableStyleId></a:tblPr>')))
    s.add(frame(4, "Unstyled table", 700000, 3700000, 7200000, 1260000, table([2400000] * 3, [[cell(f"A{r}"), cell(f"B{r}"), cell(f"C{r}")] for r in range(1, 4)], tbl_pr="<a:tblPr/>")))
    s = d.slide(3)
    s.add(title(2, "Merged cells"))
    rows = [
        [cell("Spans two columns", span=' gridSpan="2"'), cell(span=' hMerge="1"'), cell("C1"), cell("D1")],
        [cell("Spans two rows", span=' rowSpan="2"'), cell("B2"), cell("C2"), cell("D2")],
        [cell(span=' vMerge="1"'), cell("Spans 2 x 2", span=' gridSpan="2" rowSpan="2"'), cell(span=' hMerge="1"'), cell("D3")],
        [cell("A4"), cell(span=' vMerge="1"'), cell(span=' vMerge="1" hMerge="1"'), cell("D4")],
    ]
    s.add(frame(3, "Merged table", 700000, 1500000, 7200000, 1680000, table([1800000] * 4, rows, tbl_pr='<a:tblPr firstRow="1" bandRow="1"/>')))
    s = d.slide(3)
    s.add(title(2, "Borders and cell fills"))
    fills = [solid(rgb("FFF2CC")), '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="9DC3E6"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>', solid(scheme("accent2", ("lumMod", 40000), ("lumOff", 60000))), "<a:noFill/>"]
    rows = [
        [cell("Thick red box", borders="".join(border(side, w=38100) for side in "LRTB"), fill=fills[0], anchor="ctr"),
         cell("Dashed bottom", borders=border("B", "0070C0", 25400, "dash"), fill=fills[1], algn="ctr"),
         cell("Diagonal", borders='<a:lnTlToBr w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnTlToBr>', fill=fills[2]),
         cell("No fill", fill=fills[3])],
        [cell("Left only", borders=border("L", "00B050", 38100)), cell("Top only", borders=border("T", "7030A0", 38100)), cell("Right only", borders=border("R", "ED7D31", 38100)), cell("Plain")],
    ]
    s.add(frame(3, "Bordered table", 700000, 1500000, 7200000, 840000, table([1800000] * 4, rows, tbl_pr="<a:tblPr/>", row_h=420000)))
    d.write(path)


def groups_connectors(path):
    d = X.Deck(title="Groups and connectors")
    s = d.slide(3)
    s.add(title(2, "Groups"))
    box = lambda i, name, x, y, t="": sp(i, name, x, y, 1200000, 700000, fill=solid(scheme("accent1")), paras=para(t, algn="ctr") if t else "")
    s.add(grp(10, "Simple group", 700000, 1500000, 3000000, 900000, box(11, "Left", 700000, 1500000, "A") + box(12, "Right", 2500000, 1500000, "B")))
    # the child coordinate space is twice the size of the group: the children are drawn at half scale
    s.add(grp(20, "Scaled group", 4400000, 1500000, 3000000, 900000, box(21, "Scaled left", 0, 0, "C") + box(22, "Scaled right", 3000000, 0, "D"), ch=(0, 0, 6000000, 1800000)))
    inner = grp(33, "Level 3", 1000000, 4100000, 1600000, 600000, box(34, "Innermost", 1000000, 4100000, "E"))
    mid = grp(32, "Level 2", 800000, 3900000, 2200000, 1000000, inner + box(35, "Level 2 shape", 800000, 4300000, "F"))
    s.add(grp(31, "Level 1", 700000, 3800000, 3200000, 1300000, mid + box(36, "Level 1 shape", 2500000, 3800000, "G")))
    s.add(grp(40, "Group with text and a line", 4600000, 3900000, 3600000, 1000000, textbox(41, "Caption", 4600000, 3900000, 3600000, 400000, para("A caption inside a group"))
             + sp(42, "Rule", 4600000, 4400000, 3600000, 0, prst="line", ln=line(19050))))
    s = d.slide(3)
    s.add(title(2, "Connectors"))
    w, h = 1500000, 700000
    a = [(700000, 1700000), (700000, 3000000), (700000, 4300000)]
    b = [(5000000, 1700000), (5000000, 2700000), (5000000, 4900000)]
    for n, ((ax, ay), (bx, by)) in enumerate(zip(a, b)):
        s.add(sp(10 + n * 2, f"Start {n + 1}", ax, ay, w, h, fill=solid(scheme("accent1")), paras=para("start", algn="ctr")))
        s.add(sp(11 + n * 2, f"End {n + 1}", bx, by, w, h, fill=solid(scheme("accent2")), paras=para("end", algn="ctr")))
    arrow = line(25400, tail=("triangle", "med"))
    # right-middle site (3) of the start shape to left-middle site (1) of the end shape
    s.add(cxn(20, "Straight connector", 700000 + w, 1700000 + h // 2, 5000000 - 700000 - w, 0, "straightConnector1", arrow, st=(10, 3), end=(11, 1)))
    s.add(cxn(21, "Elbow connector", 700000 + w, 2700000 + h // 2, 5000000 - 700000 - w, 300000, "bentConnector3", arrow, st=(12, 3), end=(13, 1), flipV=True))
    s.add(cxn(22, "Curved connector", 700000 + w, 4300000 + h // 2, 5000000 - 700000 - w, 4900000 - 4300000, "curvedConnector3", arrow, st=(14, 3), end=(15, 1)))
    s.add(cxn(23, "Unconnected elbow", 7500000, 1700000, 1500000, 900000, "bentConnector2", line(19050, tail=("oval", "med"))))
    s.add(cxn(24, "Unconnected curve", 7500000, 3200000, 1500000, 900000, "curvedConnector2", line(19050, head=("diamond", "med"))))
    s.add(cxn(25, "Unconnected line", 7500000, 4600000, 1500000, 0, "straightConnector1", line(19050)))
    d.write(path)
