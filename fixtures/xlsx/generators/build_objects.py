"""Drawing-layer, comment, link, table, filter, pivot, sparkline, protection, print and rich-text constructs; an inert .xlsm."""
import base64
import hashlib
import struct

import xlsxbuilder as X
from build_core import sample

JPEG = base64.b64decode(
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAA0JCgsKCA0LCgsODg0PEyAVExISEyccHhcgLikxMC4pLSwzOko+MzZGNywtQFdBRkxOUlNSMj5aYVpQYEpRUk//2wBDAQ4ODhMREyYVFSZPNS01T09PT09PT09P"
    "T09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0//wAARCAAIAAgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQID"
    "AAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLD"
    "xMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMz"
    "UvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq"
    "8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDuf+Qf/wBef/oj/wBh/wDQf937pRRXRTpqorsylJxdkf/Z"
)
SVG = b'<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16"><rect width="16" height="16" fill="#1e5ac8"/><circle cx="8" cy="8" r="5" fill="#f0f0f0"/></svg>'
GIF = base64.b64decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7")
D = "http://schemas.openxmlformats.org/drawingml/2006/main"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
TC = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
EXT = '<extLst><ext uri="{%s}" xmlns:x14="' + X.X14 + '" xmlns:xm="' + X.XM + '">%s</ext></extLst>'


def emf():
    """A header-only EMF (header record plus end-of-file record): a valid, empty, inert metafile."""
    header = struct.pack("<II4i4iIIIIHHIII2i2i", 1, 88, 0, 0, 15, 15, 0, 0, 400, 400, 0x464D4520, 0x10000, 108, 2, 1, 0, 0, 0, 0, 0, 0, 0, 0)
    return header + struct.pack("<IIIII", 14, 20, 0, 16, 20)


def _drawing(book, sh, anchors, rels):
    """Attach a drawing part to a sheet; anchors is the inner XML, rels a Rels for the drawing's relationships."""
    book.add("xl/drawings/drawing1.xml", f'{X.HEAD}<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{D}" xmlns:r="{X.R}" xmlns:c="{CHART}">{anchors}</xdr:wsDr>', "drawing")
    book.add("xl/drawings/_rels/drawing1.xml.rels", rels.xml())
    sh.x["drawing"] = f'<drawing r:id="{sh.rels.add("drawing", "../drawings/drawing1.xml")}"/>'


def _from(c, r, to=None):
    m = lambda t, c_, r_: f"<xdr:{t}><xdr:col>{c_}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r_}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:{t}>"
    return m("from", c, r) + (m("to", *to) if to else "")


def _pic(i, name, rid, crop="", svg=None):
    blip = f'<a:blip r:embed="{rid}"' + (">" + f'<a:extLst><a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="{svg}"/></a:ext></a:extLst></a:blip>' if svg else "/>")
    return (f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{i + 1}" name="{name}" descr="{name}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
            f'<xdr:blipFill>{blip}{crop}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="609600" cy="609600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/>')


def images(path):
    b = X.Book(title="Images")
    sh = b.sheet("Images")
    sh.set(1, 1, "Five pictures: one-cell, two-cell (cropped), absolute, SVG with PNG fallback, EMF")
    media = {"image1.png": X.png(24, 24), "image2.jpg": JPEG, "image3.gif": GIF, "image4.png": X.png(16, 16, fg=(200, 90, 30)), "image5.svg": SVG, "image6.emf": emf()}
    for n, d in media.items():
        b.add(f"xl/media/{n}", d)
    rels = X.Rels()
    ids = {n: rels.add("image", f"../media/{n}") for n in media}
    crop = '<a:srcRect l="10000" t="20000" r="10000" b="5000"/>'
    anchors = (f'<xdr:oneCellAnchor>{_from(1, 2)}<xdr:ext cx="609600" cy="609600"/>{_pic(1, "png", ids["image1.png"])}</xdr:oneCellAnchor>'
               f'<xdr:twoCellAnchor editAs="oneCell">{_from(4, 2, (6, 5))}{_pic(2, "jpeg cropped", ids["image2.jpg"], crop)}</xdr:twoCellAnchor>'
               f'<xdr:absoluteAnchor><xdr:pos x="381000" y="1905000"/><xdr:ext cx="609600" cy="609600"/>{_pic(3, "gif", ids["image3.gif"])}</xdr:absoluteAnchor>'
               f'<xdr:oneCellAnchor>{_from(8, 2)}<xdr:ext cx="609600" cy="609600"/>{_pic(4, "svg", ids["image4.png"], svg=ids["image5.svg"])}</xdr:oneCellAnchor>'
               f'<xdr:twoCellAnchor>{_from(1, 12, (3, 15))}{_pic(5, "emf", ids["image6.emf"])}</xdr:twoCellAnchor>')
    _drawing(b, sh, anchors, rels)
    b.write(path)


def _ser(kind, i, cat, val, xcol=None):
    tx = f'<c:tx><c:strRef><c:f>Charts!${"BCDE"[i]}$1</c:f></c:strRef></c:tx>'
    v = f"<c:val><c:numRef><c:f>Charts!${'BCDE'[i]}$2:${'BCDE'[i]}$6</c:f></c:numRef></c:val>"
    c = "<c:cat><c:strRef><c:f>Charts!$A$2:$A$6</c:f></c:strRef></c:cat>"
    if kind == "scatter":
        return f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>{tx}<c:xVal><c:numRef><c:f>Charts!$B$2:$B$6</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>Charts!$C$2:$C$6</c:f></c:numRef></c:yVal></c:ser>'
    return f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>{tx}{c}{v}</c:ser>'


def chart_xml(plots):
    axes = '<c:axId val="111"/><c:axId val="222"/>'
    body = ""
    for kind, series in plots:
        s = "".join(_ser(kind, i, 0, 0) for i in series)
        body += {"bar": f'<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>{s}{axes}</c:barChart>',
                 "line": f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{s}<c:marker val="1"/>{axes}</c:lineChart>',
                 "pie": f'<c:pieChart><c:varyColors val="1"/>{s}<c:firstSliceAng val="0"/></c:pieChart>',
                 "scatter": f'<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>{s}{axes}</c:scatterChart>'}[kind]
    if plots[0][0] == "scatter":
        ax = ('<c:valAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:valAx>'
              '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/></c:valAx>')
    elif plots[0][0] == "pie":
        ax = ""
    else:
        ax = ('<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>'
              '<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/></c:valAx>')
    return (f'{X.HEAD}<c:chartSpace xmlns:c="{CHART}" xmlns:a="{D}" xmlns:r="{X.R}"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>{body}{ax}</c:plotArea>'
            '<c:legend><c:legendPos val="r"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>')


def charts(path):
    b = X.Book(title="Charts")
    sh = b.sheet("Charts")
    sh.row(1, ["Region", "Q1", "Q2", "Q3"])
    for i, r in enumerate([("North", 10, 20, 30), ("South", 5, 15, 25), ("East", 7, 8, 9), ("West", 1, 2, 3), ("Centre", 4, 6, 8)], 2):
        sh.row(i, r)
    kinds = [[("bar", [0, 1])], [("line", [0, 1])], [("pie", [0])], [("scatter", [0])], [("bar", [0]), ("line", [1])]]
    rels = X.Rels()
    anchors = ""
    for n, plots in enumerate(kinds, 1):
        b.add(f"xl/charts/chart{n}.xml", chart_xml(plots), "chart")
        rid = rels.add("chart", f"../charts/chart{n}.xml")
        anchors += (f'<xdr:twoCellAnchor>{_from(5 + (n - 1) % 2 * 8, (n - 1) // 2 * 16, (12 + (n - 1) % 2 * 8, (n - 1) // 2 * 16 + 15))}'
                    f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{n + 1}" name="Chart {n}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
                    f'<a:graphic><a:graphicData uri="{CHART}"><c:chart r:id="{rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>')
    _drawing(b, sh, anchors, rels)
    b.write(path)


VML = ('<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'
       '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>%s</xml>')
VSHAPE = ('<v:shape id="_x0000_s%d" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:%dpt;width:108pt;height:60pt;z-index:%d;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">'
          '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>'
          '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>2, 15, %d, 10, 4, 15, %d, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>%d</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>')


def comments(path):
    b = X.Book(title="Comments")
    sh = sample(b.sheet("Notes"))
    guid = "{5A6A8F52-3C1D-4E0B-9D55-0A1B2C3D4E01}"
    notes = [("A2", "Reviewer One", "Plain note on the first region."), ("B3", "Reviewer Two", "Second note, with <markup> & an ampersand."), ("C4", "tc={5A6A8F52-3C1D-4E0B-9D55-0A1B2C3D4E01}", "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel.\n\nComment:\n    Please confirm this figure.\nReply:\n    Confirmed.\n")]
    authors = ["Reviewer One", "Reviewer Two", "tc={5A6A8F52-3C1D-4E0B-9D55-0A1B2C3D4E01}"]
    cm = "".join(f'<comment ref="{r}" authorId="{authors.index(a)}"><text><t xml:space="preserve">{X.esc(t)}</t></text></comment>' for r, a, t in notes)
    b.add("xl/comments1.xml", f'{X.HEAD}<comments xmlns="{X.M}"><authors>{"".join(f"<author>{a}</author>" for a in authors)}</authors><commentList>{cm}</commentList></comments>', "comments")
    b.add("xl/drawings/vmlDrawing1.vml", X.HEAD + VML % "".join(VSHAPE % (1025 + i, 10 + 20 * i, i + 1, r_, r_ + 3, r_) for i, r_ in enumerate([1, 2, 3])))
    sh.x["legacyDrawing"] = f'<legacyDrawing r:id="{sh.rels.add("vmlDrawing", "../drawings/vmlDrawing1.vml")}"/>'
    sh.rels.add("comments", "../comments1.xml")
    b.add("xl/threadedComments/threadedComment1.xml",
          f'{X.HEAD}<ThreadedComments xmlns="{TC}"><threadedComment ref="C4" dT="2026-01-15T09:00:00.00" personId="{{00000000-0000-0000-0000-000000000001}}" id="{guid}"><text>Please confirm this figure.</text></threadedComment>'
          f'<threadedComment ref="C4" dT="2026-01-15T09:30:00.00" personId="{{00000000-0000-0000-0000-000000000002}}" id="{{5A6A8F52-3C1D-4E0B-9D55-0A1B2C3D4E02}}" parentId="{guid}"><text>Confirmed.</text></threadedComment></ThreadedComments>',
          "application/vnd.ms-excel.threadedcomments+xml")
    sh.rels.add("http://schemas.microsoft.com/office/2017/10/relationships/threadedComment", "../threadedComments/threadedComment1.xml")
    b.add("xl/persons/person.xml", f'{X.HEAD}<personList xmlns="{TC}" xmlns:x="{X.M}"><person displayName="Reviewer One" id="{{00000000-0000-0000-0000-000000000001}}" userId="reviewer.one" providerId="None"/>'
          '<person displayName="Reviewer Two" id="{00000000-0000-0000-0000-000000000002}" userId="reviewer.two" providerId="None"/></personList>', "application/vnd.ms-excel.person+xml")
    b.wbrels.add("http://schemas.microsoft.com/office/2017/10/relationships/person", "persons/person.xml")
    b.write(path)


def links(path):
    b = X.Book(title="Hyperlinks")
    sh = b.sheet("Links")
    sh.set(1, 1, "Kind").set(1, 2, "Link")
    targets = [("https", "https://example.com/spindle-fixture", None), ("mailto", "mailto:someone@example.com?subject=Fixture", None), ("internal", None, "Links!A1"), ("javascript", "javascript:alert(1)", None)]
    hl = ""
    for i, (kind, target, loc) in enumerate(targets, 2):
        sh.set(i, 1, kind).set(i, 2, f"{kind} link")
        hl += f'<hyperlink ref="B{i}" ' + (f'r:id="{sh.rels.add("hyperlink", target, external=True)}"' if target else f'location="{loc}" display="Back to top"') + ' tooltip="fixture link"/>'
    sh.x["hyperlinks"] = f"<hyperlinks>{hl}</hyperlinks>"
    b.write(path)


def tables(path):
    b = X.Book(title="Tables")
    sh = b.sheet("Tables")
    sh.row(1, ["Region", "Q1", "Q2", "Q3", "Total"])
    for i, r in enumerate([("North", 10, 20, 30), ("South", 5, 15, 25), ("East", 7, 8, 9), ("West", 1, 2, 3)], 2):
        sh.row(i, r)
        sh.set(i, 5, ("f", "Sales[[#This Row],[Q1]]+Sales[[#This Row],[Q2]]+Sales[[#This Row],[Q3]]", sum(r[1:]), ""))
    sh.set(6, 1, "Total")
    for c, v in zip(range(2, 6), (23, 45, 67, 135)):
        sh.set(6, c, ("f", f"SUBTOTAL(109,Sales[{['Q1', 'Q2', 'Q3', 'Total'][c - 2]}])", v, ""))
    sh.set(8, 1, "Sum of Q1").set(8, 2, ("f", "SUM(Sales[Q1])", 23, ""))
    cols = '<tableColumn id="1" name="Region" totalsRowLabel="Total"/>' + "".join(f'<tableColumn id="{i}" name="Q{i - 1}" totalsRowFunction="sum"/>' for i in (2, 3, 4))
    cols += '<tableColumn id="5" name="Total" totalsRowFunction="sum"><calculatedColumnFormula>Sales[[#This Row],[Q1]]+Sales[[#This Row],[Q2]]+Sales[[#This Row],[Q3]]</calculatedColumnFormula></tableColumn>'
    b.add("xl/tables/table1.xml", f'{X.HEAD}<table xmlns="{X.M}" id="1" name="Sales" displayName="Sales" ref="A1:E6" totalsRowCount="1"><autoFilter ref="A1:E5"/><tableColumns count="5">{cols}</tableColumns>'
          '<tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>', "table")
    sh.x["tableParts"] = f'<tableParts count="1"><tablePart r:id="{sh.rels.add("table", "../tables/table1.xml")}"/></tableParts>'
    b.write(path)


def filters_sort(path):
    b = X.Book(title="Filters and sort")
    sh = b.sheet("Filters")
    sh.row(1, ["Region", "Q1", "Q2", "Q3"])
    for i, r in enumerate([("North", 10, 20, 30), ("East", 7, 8, 9), ("South", 5, 15, 25), ("Centre", 4, 6, 8), ("West", 1, 2, 3)], 2):
        sh.row(i, r)
    for r in (4, 5, 6):
        sh.rowattrs[r] = 'hidden="1"'
    sh.x["autoFilter"] = ('<autoFilter ref="A1:D6"><filterColumn colId="0"><filters><filter val="North"/><filter val="East"/></filters></filterColumn>'
                          '<filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="4"/></customFilters></filterColumn>'
                          '<sortState ref="A2:D6"><sortCondition descending="1" ref="B2:B6"/></sortState></autoFilter>')
    b.names = ['<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Filters!$A$1:$D$6</definedName>']
    b.write(path)


def pivots_slicers(path):
    b = X.Book(title="Pivot table and slicer")
    data = b.sheet("Data")
    data.row(1, ["Region", "Product", "Units", "Sales"])
    rows = [("North", "Pen", 3, 30), ("South", "Pen", 2, 20), ("North", "Ink", 1, 15), ("East", "Pen", 9, 90), ("South", "Ink", 4, 60)]
    for i, r in enumerate(rows, 2):
        data.row(i, r)
    pv = b.sheet("Pivot")
    pv.set(1, 1, "Sales by region")
    regions, products = ["East", "North", "South"], ["Ink", "Pen"]
    b.add("xl/pivotCache/pivotCacheDefinition1.xml",
          f'{X.HEAD}<pivotCacheDefinition xmlns="{X.M}" xmlns:r="{X.R}" r:id="rId1" refreshOnLoad="1" refreshedBy="Fixture" refreshedDate="46037" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="5">'
          '<cacheSource type="worksheet"><worksheetSource ref="A1:D6" sheet="Data"/></cacheSource><cacheFields count="4">'
          f'<cacheField name="Region" numFmtId="0"><sharedItems count="3">{"".join(f"<s v=\"{r}\"/>" for r in regions)}</sharedItems></cacheField>'
          f'<cacheField name="Product" numFmtId="0"><sharedItems count="2">{"".join(f"<s v=\"{p}\"/>" for p in products)}</sharedItems></cacheField>'
          '<cacheField name="Units" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="9"/></cacheField>'
          '<cacheField name="Sales" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="15" maxValue="90"/></cacheField></cacheFields></pivotCacheDefinition>',
          X.SS + "pivotCacheDefinition+xml")
    recs = "".join(f'<r><x v="{regions.index(r[0])}"/><x v="{products.index(r[1])}"/><n v="{r[2]}"/><n v="{r[3]}"/></r>' for r in rows)
    b.add("xl/pivotCache/pivotCacheRecords1.xml", f'{X.HEAD}<pivotCacheRecords xmlns="{X.M}" count="5">{recs}</pivotCacheRecords>', X.SS + "pivotCacheRecords+xml")
    cr = X.Rels()
    cr.add("pivotCacheRecords", "pivotCacheRecords1.xml")
    b.add("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", cr.xml())
    b.add("xl/pivotTables/pivotTable1.xml",
          f'{X.HEAD}<pivotTableDefinition xmlns="{X.M}" name="PivotTable1" cacheId="1" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">'
          '<location ref="A3:B7" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="4"><pivotField axis="axisRow" showAll="0"><items count="4"><item x="0"/><item x="1"/><item x="2"/><item t="default"/></items></pivotField>'
          '<pivotField showAll="0"/><pivotField showAll="0"/><pivotField dataField="1" showAll="0"/></pivotFields><rowFields count="1"><field x="0"/></rowFields>'
          '<rowItems count="4"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i><i t="grand"><x/></i></rowItems><colItems count="1"><i/></colItems>'
          '<dataFields count="1"><dataField name="Sum of Sales" fld="3" baseField="0" baseItem="0"/></dataFields><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>',
          X.SS + "pivotTable+xml")
    pr = X.Rels()
    pr.add("pivotCacheDefinition", "../pivotCache/pivotCacheDefinition1.xml")
    b.add("xl/pivotTables/_rels/pivotTable1.xml.rels", pr.xml())
    pv.rels.add("pivotTable", "../pivotTables/pivotTable1.xml")
    cache_rid = b.wbrels.add("pivotCacheDefinition", "pivotCache/pivotCacheDefinition1.xml")
    b.wb["pivotCaches"] = f'<pivotCaches><pivotCache cacheId="1" r:id="{cache_rid}"/></pivotCaches>'
    S = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/slicer"
    b.add("xl/slicerCaches/slicerCache1.xml", f'{X.HEAD}<slicerCacheDefinition xmlns="{S}" name="Slicer_Region" sourceName="Region"><pivotTables><pivotTable tabId="2" name="PivotTable1"/></pivotTables>'
          '<data><tabular pivotCacheId="1"><items count="3"><i x="0" s="1"/><i x="1" s="1"/><i x="2" s="1"/></items></tabular></data></slicerCacheDefinition>', "application/vnd.ms-excel.slicerCache+xml")
    b.add("xl/slicers/slicer1.xml", f'{X.HEAD}<slicers xmlns="{S}"><slicer name="Region" cache="Slicer_Region" caption="Region" rowHeight="241300"/></slicers>', "application/vnd.ms-excel.slicer+xml")
    sc = b.wbrels.add("http://schemas.microsoft.com/office/2007/relationships/slicerCache", "slicerCaches/slicerCache1.xml")
    b.wb["extLst"] = EXT % ("BBE1A952-AA13-448e-AADC-164F8A28A991", f'<x14:slicerCaches xmlns:x14="{X.X14}"><x14:slicerCache r:id="{sc}"/></x14:slicerCaches>')
    b.names = ['<definedName name="Slicer_Region">#N/A</definedName>']
    pv.x["extLst"] = EXT % ("A8765BA9-456A-4dab-B4F3-ACF838C121DE", f'<x14:slicerList><x14:slicer r:id="{pv.rels.add("http://schemas.microsoft.com/office/2007/relationships/slicer", "../slicers/slicer1.xml")}"/></x14:slicerList>')
    b.write(path)


def sparklines(path):
    b = X.Book(title="Sparklines")
    sh = b.sheet("Sparks")
    sh.row(1, ["Region", "Q1", "Q2", "Q3", "Q4", "Trend", "Column"])
    for i, r in enumerate([("North", 10, 20, 15, 30), ("South", 5, -15, 25, 10), ("East", 7, 8, 9, 6)], 2):
        sh.row(i, r)
    def group(kind, col):
        c = "".join(f'<x14:color{n} rgb="FF{v}"/>' for n, v in [("Series", "376092"), ("Negative", "D00000"), ("Axis", "000000"), ("Markers", "D00000"), ("First", "D00000"), ("Last", "D00000"), ("High", "D00000"), ("Low", "D00000")])
        items = "".join(f"<x14:sparkline><xm:f>Sparks!B{r}:E{r}</xm:f><xm:sqref>{col}{r}</xm:sqref></x14:sparkline>" for r in (2, 3, 4))
        return f'<x14:sparklineGroup {kind}displayEmptyCellsAs="gap" markers="1" high="1" low="1">{c}<x14:sparklines>{items}</x14:sparklines></x14:sparklineGroup>'
    sh.x["extLst"] = EXT % ("05C60535-1F16-4fd2-B633-F4F36F0B64E0", "<x14:sparklineGroups>" + group("", "F") + group('type="column" ', "G") + "</x14:sparklineGroups>")
    b.write(path)


def x14_cf(path):
    b = X.Book(title="Conditional formatting x14")
    sh = b.sheet("CFx14")
    sh.row(1, ["Bar", "Icons"])
    for i, v in enumerate([-5, 3, 8, -2, 10], 2):
        sh.set(i, 1, v).set(i, 2, v * 3)
    ids = ("{E1B1C2D3-0001-4A00-8000-000000000001}", "{E1B1C2D3-0002-4A00-8000-000000000002}")
    sh.x["conditionalFormatting"] = (
        f'<conditionalFormatting sqref="A2:A6"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{{B025F937-C7B1-47D3-B67F-A62EFF666E3E}}" xmlns:x14="{X.X14}"><x14:id>{ids[0]}</x14:id></ext></extLst></cfRule></conditionalFormatting>'
        f'<conditionalFormatting sqref="B2:B6"><cfRule type="iconSet" priority="2"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet><extLst><ext uri="{{B025F937-C7B1-47D3-B67F-A62EFF666E3E}}" xmlns:x14="{X.X14}"><x14:id>{ids[1]}</x14:id></ext></extLst></cfRule></conditionalFormatting>')
    bar = f'<x14:conditionalFormatting><x14:cfRule type="dataBar" id="{ids[0]}"><x14:dataBar minLength="0" maxLength="100" gradient="0" negativeBarColorSameAsPositive="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>A2:A6</xm:sqref></x14:conditionalFormatting>'
    icons = (f'<x14:conditionalFormatting><x14:cfRule type="iconSet" priority="2" id="{ids[1]}"><x14:iconSet iconSet="3Arrows" custom="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo>'
             '<x14:cfIcon iconSet="3Flags" iconId="0"/><x14:cfIcon iconSet="3Flags" iconId="1"/><x14:cfIcon iconSet="3Flags" iconId="2"/></x14:iconSet></x14:cfRule><xm:sqref>B2:B6</xm:sqref></x14:conditionalFormatting>')
    sh.x["extLst"] = EXT % ("78C0D931-6437-407d-A8EE-F0AAD7539E65", f"<x14:conditionalFormattings>{bar}{icons}</x14:conditionalFormattings>")
    b.write(path)


def legacy_hash(pw):
    h = 0
    for ch in reversed(pw):
        h = (((h >> 14) & 1) | ((h << 1) & 0x7FFF)) ^ ord(ch)
    h = ((h >> 14) & 1) | ((h << 1) & 0x7FFF)
    return "%04X" % (h ^ len(pw) ^ 0xCE4B)


def sha512_hash(pw, salt, spin=100000):
    h = hashlib.sha512(salt + pw.encode("utf-16-le")).digest()
    for i in range(spin):
        h = hashlib.sha512(h + struct.pack("<I", i)).digest()
    return base64.b64encode(h).decode()


def protection(path):
    b = X.Book(title="Protection")
    unlocked = b.styles.xf(prot='<protection locked="0"/>', fmt="0.00")
    salt = bytes(range(16))
    kinds = [("Legacy hash", f'<sheetProtection password="{legacy_hash("spindle")}" sheet="1" objects="1" scenarios="1" formatCells="0" selectLockedCells="1"/>'),
             ("SHA-512 hash", f'<sheetProtection algorithmName="SHA-512" hashValue="{sha512_hash("spindle", salt)}" saltValue="{base64.b64encode(salt).decode()}" spinCount="100000" sheet="1" objects="1" scenarios="1"/>'),
             ("No password", '<sheetProtection sheet="1" objects="1" scenarios="1" insertRows="0" sort="0"/>')]
    for n, xml in kinds:
        sh = b.sheet(n)
        sh.set(1, 1, "Locked cell").set(2, 1, 5, unlocked).set(2, 2, "unlocked input")
        sh.x["sheetProtection"] = xml
    b.wb["protection"] = '<workbookProtection lockStructure="1"/>'
    b.write(path)


def print_setup(path):
    b = X.Book(title="Print setup")
    sh = sample(b.sheet("Print"))
    sh.x["sheetPr"] = '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
    sh.x["printOptions"] = '<printOptions horizontalCentered="1" gridLines="1"/>'
    sh.x["pageMargins"] = '<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    sh.x["pageSetup"] = '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>'
    sh.x["headerFooter"] = '<headerFooter><oddHeader>&amp;C&amp;"Calibri,Bold"Quarterly figures</oddHeader><oddFooter>&amp;LPage &amp;P of &amp;N&amp;R&amp;D</oddFooter></headerFooter>'
    sh.x["rowBreaks"] = '<rowBreaks count="1" manualBreakCount="1"><brk id="3" max="16383" man="1"/></rowBreaks>'
    sh.x["colBreaks"] = '<colBreaks count="1" manualBreakCount="1"><brk id="2" max="1048575" man="1"/></colBreaks>'
    b.names = ['<definedName name="_xlnm.Print_Area" localSheetId="0">Print!$A$1:$D$6</definedName>', '<definedName name="_xlnm.Print_Titles" localSheetId="0">Print!$1:$1</definedName>']
    b.write(path)


def rich_text(path):
    b = X.Book(title="Rich text")
    sh = b.sheet("Rich")
    runs = ('<si><r><t xml:space="preserve">plain </t></r><r><rPr><b/><sz val="11"/><color rgb="FFFF0000"/><rFont val="Calibri"/><family val="2"/></rPr><t>bold red</t></r>'
            '<r><rPr><i/><u/><sz val="14"/><color theme="1"/><rFont val="Cambria"/><family val="1"/></rPr><t xml:space="preserve"> italic underlined</t></r><r><rPr><vertAlign val="superscript"/><sz val="11"/><rFont val="Calibri"/></rPr><t>2</t></r></si>')
    sh.set(1, 1, runs).set(2, 1, "plain text for contrast")
    sh.set(3, 1, ("raw", '<is><r><rPr><b/><sz val="11"/><rFont val="Calibri"/></rPr><t>inline bold</t></r><r><t xml:space="preserve"> and plain</t></r></is>'), t="inlineStr")
    b.write(path)


def macro_enabled(path):
    b = X.Book(title="Macro-enabled workbook")
    b.ext_ct["workbook"] = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
    sample(b.sheet("Macros"))
    b.add("xl/vbaProject.bin", b"SPINDLE-FIXTURE-INERT-DUMMY: not an OLE file and not VBA; it only exercises macro-part handling.\n")
    b.wbrels.add("http://schemas.microsoft.com/office/2006/relationships/vbaProject", "vbaProject.bin")
    b.write(path)
