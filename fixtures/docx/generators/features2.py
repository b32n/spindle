"""More feature blocks: images, shapes/text boxes/SmartArt, equations, content controls and forms."""
from docbuilder import Doc, paragraph_text
from features import H, body_p
from ooxml import CT, EMU, NS, NS_DECL, esc, fld, p, r, tbl, tc, tr, write_zip


# ================================================================ images
def images(doc: Doc):
    out = [H("Images", 1)]
    out.append(p(r("Inline images of different sizes: ") + doc.inline("rect", 96, 60, "blue rectangle") + r(" ") + doc.inline("circle", 72, 72, "blue circle", fg=(20, 140, 90)) + r(" ")
                 + doc.inline("diag", 120, 48, "diagonal band", fg=(180, 60, 140)) + r(" in running text."), style="BodyText"))
    out.append(p(doc.inline("checker", 240, 150, "checkerboard, block-level inline image", fg=(60, 60, 60)), ppr='<w:jc w:val="center"/>'))
    out.append(p(r("Figure ") + fld("SEQ Figure \\* ARABIC", "1") + r(": a centred inline image with a caption."), style="Caption", ppr='<w:jc w:val="center"/>'))
    # floating, wrapped
    out.append(p(doc.floating("circle", 130, 130, "circle wrapped square, right", wrap="square", hpos=("margin", "right"), vpos=("paragraph", 0))
                 + r(paragraph_text(210, 8)), style="BodyText"))
    out.append(p(doc.floating("stripes", 120, 90, "stripes wrapped square, left", wrap="square", hpos=("margin", "left"), vpos=("paragraph", 0), fg=(200, 120, 20))
                 + r(paragraph_text(211, 8)), style="BodyText"))
    out.append(p(doc.floating("rect", 110, 110, "tight wrap image", wrap="tight", hpos=("margin", "right"), vpos=("paragraph", 0), fg=(90, 40, 160))
                 + r(paragraph_text(212, 7)), style="BodyText"))
    out.append(p(doc.floating("diag", 300, 80, "top and bottom wrap banner", wrap="topAndBottom", hpos=("margin", "center"), vpos=("paragraph", 100000), fg=(20, 90, 140))
                 + r(paragraph_text(213, 4)), style="BodyText"))
    out.append(p(doc.floating("checker", 200, 200, "behind-text watermark-like image", wrap="none", hpos=("page", "center"), vpos=("page", 2000000), behind=True, fg=(225, 225, 225))
                 + r("Text over an image placed behind the text. " + paragraph_text(214, 3)), style="BodyText"))
    # cropped
    crops = [(0, 0, 0, 0), (25000, 0, 0, 0), (0, 30000, 30000, 0), (10000, 10000, 40000, 40000)]
    cells = []
    for i, c in enumerate(crops):
        cells.append(tc(p(doc.inline("checker", 160, 160, f"crop {c}", crop=c if any(c) else None, fg=(30, 90, 200)), ppr='<w:jc w:val="center"/>') + p(r(f"srcRect {c}" if any(c) else "uncropped"), style="Caption"), w=2300))
    out.append(p("Cropped images (a:srcRect) inside a table, next to an uncropped original:", style="BodyText"))
    out.append(tbl([tr(cells)], [2300] * 4))
    out.append(p(r("Image-only paragraph followed by alt-text-less image: ") + doc.inline("circle", 40, 40, ""), style="BodyText"))
    return "".join(out)


# ================================================================= shapes
def _wsp(name, wid, x, y, cx, cy, prst, fill, text_xml="", line="1F3864", wrap="wrapSquare"):
    emu = lambda v: v * EMU
    txbx = f"<wps:txbx><w:txbxContent>{text_xml}</w:txbxContent></wps:txbx>" if text_xml else ""
    return (f'<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu(cx)}" cy="{emu(cy)}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln></wps:spPr>{txbx}<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp>')


def _anchor_drawing(doc, inner_graphic, cx, cy, name, hpos="margin", halign="left", vpos=0, wrap="<wp:wrapSquare wrapText=\"bothSides\"/>"):
    pid = doc.nid()
    emu = lambda v: v * EMU
    return (f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="{hpos}"><wp:align>{halign}</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>{vpos}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{emu(cx)}" cy="{emu(cy)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrap}<wp:docPr id="{pid}" name="{name} {pid}"/><wp:cNvGraphicFramePr/>{inner_graphic}</wp:anchor>')


def shapes(doc: Doc):
    out = [H("Shapes, text boxes and SmartArt", 1)]
    # 1. DrawingML shape with text (wps), anchored, with AlternateContent + VML fallback
    boxtext = p(r("Text box inside a DrawingML shape", b=True)) + p(r(paragraph_text(230, 2))) + p(r("A bulleted line inside the box."), style="ListParagraph", ppr=num_ppr())
    wsp = _wsp("Rect", 1, 0, 0, 260, 120, "roundRect", "FFF2CC", boxtext)
    graphic = f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">{wsp}</a:graphicData></a:graphic>'
    anchor = _anchor_drawing(doc, graphic, 260, 120, "Text Box", halign="right")
    vml = ('<w:pict><v:roundrect id="_x0000_s1026" style="position:absolute;margin-left:0;margin-top:0;width:195pt;height:90pt;z-index:251659264;mso-position-horizontal:right" arcsize="10923f" fillcolor="#fff2cc" strokecolor="#1f3864">'
           f'<v:textbox><w:txbxContent>{boxtext}</w:txbxContent></v:textbox><w10:wrap type="square"/></v:roundrect></w:pict>')
    out.append(p(r("A paragraph that wraps around an anchored shape containing text. ") +
                 f'<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>{anchor}</w:drawing></mc:Choice><mc:Fallback>{vml}</mc:Fallback></mc:AlternateContent></w:r>'
                 + r(paragraph_text(231, 6)), style="BodyText"))
    # 2. inline shapes without text: ellipse, arrow, line (DrawingML only)
    row = ""
    for prst, fill, w, h in (("ellipse", "BDD7EE", 90, 60), ("rightArrow", "C5E0B4", 100, 50), ("triangle", "F8CBAD", 70, 70), ("star5", "FFE699", 70, 70), ("line", "7F7F7F", 120, 2)):
        g = f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">{_wsp(prst, 1, 0, 0, w, h, prst, fill)}</a:graphicData></a:graphic>'
        pid = doc.nid()
        row += f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{w * EMU}" cy="{h * EMU}"/><wp:docPr id="{pid}" name="Shape {pid}" descr="{prst} shape"/><wp:cNvGraphicFramePr/>{g}</wp:inline></w:drawing></w:r>' + r("  ")
    out.append(p(r("Inline shapes: ") + row, style="BodyText"))
    # 3. group of shapes
    grp = ('<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/><a:chOff x="0" y="0"/><a:chExt cx="%d" cy="%d"/></a:xfrm></wpg:grpSpPr>%s%s</wpg:wgp>'
           % (300 * EMU, 100 * EMU, 300 * EMU, 100 * EMU, _wsp("a", 2, 0, 0, 140, 100, "rect", "DDEBF7", p(r("Group child A"))), _wsp("b", 3, 160, 0, 140, 100, "ellipse", "E2EFDA", p(r("Group child B")))))
    out.append(p(r("A shape group with two children containing text: ") + f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{300 * EMU}" cy="{100 * EMU}"/><wp:docPr id="{doc.nid()}" name="Group"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">{grp}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>', style="BodyText"))
    # 4. legacy VML text box + VML rectangle with text, plus a text box holding a table and a list
    tb_inner = p(r("Legacy VML text box.", b=True)) + p(r(paragraph_text(232, 1))) + tbl([tr([tc("t1", w=1200), tc("t2", w=1200)])], [1200, 1200]) + p("")
    out.append(p(r("A legacy VML text box (Word 2003 style) that contains a table: ") +
                 f'<w:r><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
                 f'<v:shape id="_x0000_s1027" type="#_x0000_t202" style="width:200pt;height:110pt;visibility:visible" fillcolor="#e2efda" strokecolor="#548235"><v:textbox><w:txbxContent>{tb_inner}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>', style="BodyText"))
    out.append(smartart(doc))
    return "".join(out)


def num_ppr():
    return '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'


def smartart(doc: Doc):
    """A basic list SmartArt: data model, layout, style, colors parts and the dsp fallback drawing."""
    nsd = 'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"'
    labels = ["Survey", "Review", "Approve", "Publish"]
    pts = '<dgm:pt modelId="{00000000-0000-0000-0000-000000000001}" type="doc"><dgm:prSet loTypeId="urn:microsoft.com/office/officeart/2005/8/layout/default" loCatId="list" qsTypeId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1" qsCatId="simple" csTypeId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_2" csCatId="accent1"/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>'
    cx = ""
    for i, t in enumerate(labels):
        pts += f'<dgm:pt modelId="{{00000000-0000-0000-0000-0000000001{i:02d}}}"><dgm:prSet phldrT="[Text]"/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{t}</a:t></a:r></a:p></dgm:t></dgm:pt>'
        cx += f'<dgm:cxn modelId="{{00000000-0000-0000-0000-0000000002{i:02d}}}" srcId="{{00000000-0000-0000-0000-000000000001}}" destId="{{00000000-0000-0000-0000-0000000001{i:02d}}}" srcOrd="{i}" destOrd="0"/>'
    data = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:dataModel {nsd}><dgm:ptLst>{pts}</dgm:ptLst><dgm:cxnLst>{cx}</dgm:cxnLst><dgm:bg/><dgm:whole/>'
            '<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt relId="rIdDsp" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>')
    layout = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:layoutDef {nsd} uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/default"><dgm:title val="Basic Block List"/><dgm:desc val=""/>'
              '<dgm:catLst><dgm:cat type="list" pri="400"/></dgm:catLst><dgm:layoutNode name="diagram"><dgm:alg type="snake"/><dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/>'
              '<dgm:forEach name="Name0" axis="ch" ptType="node"><dgm:layoutNode name="node"><dgm:alg type="tx"/><dgm:shape type="rect" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf axis="desOrSelf" ptType="node"/></dgm:layoutNode></dgm:forEach></dgm:layoutNode></dgm:layoutDef>')
    style = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:styleDef {nsd} uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="simple" pri="10100"/></dgm:catLst><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:styleLbl name="node1"><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:sp3d/><dgm:txPr/><dgm:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dgm:style></dgm:styleLbl></dgm:styleDef>'
    colors = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dgm:colorsDef {nsd} uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_2"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="accent1" pri="11200"/></dgm:catLst><dgm:styleLbl name="node1"><dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:fillClrLst><dgm:linClrLst meth="repeat"><a:schemeClr val="lt1"/></dgm:linClrLst><dgm:effectClrLst/><dgm:txLinClrLst/><dgm:txFillClrLst/><dgm:txEffectClrLst/></dgm:styleLbl></dgm:colorsDef>'
    sps = ""
    for i, t in enumerate(labels):
        sps += (f'<dsp:sp modelId="{{00000000-0000-0000-0000-0000000003{i:02d}}}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>'
                f'<dsp:spPr><a:xfrm><a:off x="{i * 1000000}" y="0"/><a:ext cx="900000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></dsp:spPr>'
                f'<dsp:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>{t}</a:t></a:r></a:p></dsp:txBody>'
                f'<dsp:txXfrm><a:off x="{i * 1000000}" y="0"/><a:ext cx="900000" cy="600000"/></dsp:txXfrm></dsp:sp>')
    drawing = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dsp:drawing {nsd}><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>{sps}</dsp:spTree></dsp:drawing>'
    base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    rdm = doc.part("word/diagrams/data1.xml", data.replace("rIdDsp", "__DSP__"), "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml", base + "diagramData")
    rlo = doc.part("word/diagrams/layout1.xml", layout, "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml", base + "diagramLayout")
    rqs = doc.part("word/diagrams/quickStyle1.xml", style, "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml", base + "diagramQuickStyle")
    rcs = doc.part("word/diagrams/colors1.xml", colors, "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml", base + "diagramColors")
    rdsp = doc.part("word/diagrams/drawing1.xml", drawing, "application/vnd.ms-office.drawingml.diagramDrawing+xml", "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing")
    doc.parts["word/diagrams/data1.xml"] = doc.parts["word/diagrams/data1.xml"].replace(b"__DSP__", rdsp.encode())
    pid = doc.nid()
    g = (f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="{rdm}" r:lo="{rlo}" r:qs="{rqs}" r:cs="{rcs}"/></a:graphicData></a:graphic>')
    return p(r("A SmartArt diagram (basic block list) with its cached fallback drawing:", style=None), style="BodyText") + p(
        f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{4000 * 1000 * 9525 // 9525}" cy="{1200 * 1000}"/><wp:docPr id="{pid}" name="Diagram {pid}" descr="Four step process: Survey, Review, Approve, Publish"/><wp:cNvGraphicFramePr/>{g}</wp:inline></w:drawing></w:r>')


# ============================================================= equations
def equations(doc: Doc):
    mr = lambda t, sty=None: f'<m:r>{"<m:rPr><m:sty m:val=\"%s\"/></m:rPr>" % sty if sty else ""}<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>{esc(t)}</m:t></m:r>'
    frac = lambda n, d: f"<m:f><m:num>{n}</m:num><m:den>{d}</m:den></m:f>"
    sup = lambda b, e: f"<m:sSup><m:e>{b}</m:e><m:sup>{e}</m:sup></m:sSup>"
    sub = lambda b, s: f"<m:sSub><m:e>{b}</m:e><m:sub>{s}</m:sub></m:sSub>"
    nary = lambda ch, lo, hi, body: f'<m:nary><m:naryPr><m:chr m:val="{ch}"/><m:limLoc m:val="subSup"/></m:naryPr><m:sub>{lo}</m:sub><m:sup>{hi}</m:sup><m:e>{body}</m:e></m:nary>'
    delim = lambda inner, b="(", e=")": f'<m:d><m:dPr><m:begChr m:val="{b}"/><m:endChr m:val="{e}"/></m:dPr><m:e>{inner}</m:e></m:d>'
    rad = lambda inner, deg="": f'<m:rad><m:radPr>{"<m:degHide m:val=\"1\"/>" if not deg else ""}</m:radPr><m:deg>{deg}</m:deg><m:e>{inner}</m:e></m:rad>'
    matrix = '<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="3"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>' + "".join(
        "<m:mr>" + "".join(f"<m:e>{mr(v)}</m:e>" for v in row) + "</m:mr>" for row in (("1", "0", "2"), ("0", "1", "-1"), ("3", "4", "1"))) + "</m:m>"
    func = lambda name, arg: f"<m:func><m:fName>{mr(name, 'p')}</m:fName><m:e>{arg}</m:e></m:func>"
    out = [H("Equations (OMML)", 1)]
    disp = lambda inner: p(f"<m:oMathPara><m:oMathParaPr><m:jc m:val=\"center\"/></m:oMathParaPr><m:oMath>{inner}</m:oMath></m:oMathPara>")
    out.append(p("Display equations follow: the quadratic formula, a definite integral, a summation, a matrix product, a limit and a system of equations.", style="BodyText"))
    out.append(disp(mr("x=") + frac(mr("-b±") + rad(sup(mr("b"), mr("2")) + mr("-4ac")), mr("2a"))))
    out.append(disp(nary("∫", mr("0"), mr("∞"), sup(mr("e"), mr("-x") + sup(mr(""), mr("2"))) + mr("dx")) + mr("=") + frac(rad(mr("π")), mr("2"))))
    out.append(disp(nary("∑", mr("k=1"), mr("n"), sup(mr("k"), mr("2"))) + mr("=") + frac(mr("n") + delim(mr("n+1")) + delim(mr("2n+1")), mr("6"))))
    out.append(disp(mr("A=") + delim(matrix, "[", "]") + mr(",  ") + mr("det(A)=") + mr("-9")))
    out.append(disp(func("lim", sub(mr(""), mr("x→0")) + frac(func("sin", mr("x")), mr("x"))) + mr("=1")))
    eqarr = '<m:eqArr><m:eqArrPr><m:maxDist m:val="1"/></m:eqArrPr>' + "".join(f"<m:e>{mr(t)}</m:e>" for t in ("2x+y=5", "x-y=1")) + "</m:eqArr>"
    out.append(disp(delim(eqarr, "{", "")))
    out.append(disp(sub(mr("E"), mr("k")) + mr("=") + frac(mr("1"), mr("2")) + mr("m") + sup(mr("v"), mr("2")) + mr(",  ") + f'<m:acc><m:accPr><m:chr m:val="⃗"/></m:accPr><m:e>{mr("F")}</m:e></m:acc>' + mr("=m") + f'<m:acc><m:accPr><m:chr m:val="⃗"/></m:accPr><m:e>{mr("a")}</m:e></m:acc>'))
    out.append(p(r("Inline math: the area of a circle is ") + f"<m:oMath>{mr('A=π')}{sup(mr('r'), mr('2'))}</m:oMath>" + r(", and the golden ratio is ") + f"<m:oMath>{mr('φ=')}{frac(mr('1+') + rad(mr('5')), mr('2'))}</m:oMath>" + r("."), style="BodyText"))
    out.append(p(r("Equation inside a table cell and a text box is also common:"), style="BodyText"))
    out.append(tbl([tr([tc(p(f"<m:oMath>{frac(mr('a'), mr('b'))}{mr('+')}{frac(mr('c'), mr('d'))}</m:oMath>"), w=3000), tc("a/b + c/d", w=3000)])], [3000, 3000]))
    return "".join(out)


# ================================================ content controls / forms
def _sdt_pr(alias, tag, kind, sid):
    return f'<w:sdtPr><w:alias w:val="{alias}"/><w:tag w:val="{tag}"/><w:id w:val="{sid}"/>{kind}</w:sdtPr>'


def sdt_forms(doc: Doc):
    ph = lambda t: r(t, color="808080")
    out = [H("Content controls and form fields", 1)]
    # block-level rich text control
    out.append(f'<w:sdt>{_sdt_pr("Summary", "summary", "<w:showingPlcHdr/>", 1001)}<w:sdtContent>{p(ph("Click here to enter a summary."), style="BodyText")}</w:sdtContent></w:sdt>')
    # rich text with content
    out.append(f'<w:sdt>{_sdt_pr("Notes", "notes", "", 1002)}<w:sdtContent>{p(r("Rich text control with ") + r("bold", b=True) + r(" and ") + r("italic", i=True) + r(" runs."), style="BodyText")}{p(r("And a second paragraph inside the same control."), style="BodyText")}</w:sdtContent></w:sdt>')
    # inline controls in a paragraph
    plain = f'<w:sdt>{_sdt_pr("Name", "name", "<w:text/>", 1003)}<w:sdtContent>{r("Ada Marlowe")}</w:sdtContent></w:sdt>'
    drop = (f'<w:sdt>{_sdt_pr("Berth", "berth", "<w:dropDownList><w:listItem w:displayText=\"North\" w:value=\"N\"/><w:listItem w:displayText=\"East\" w:value=\"E\"/><w:listItem w:displayText=\"South\" w:value=\"S\"/></w:dropDownList>", 1004)}'
            f'<w:sdtContent>{r("East")}</w:sdtContent></w:sdt>')
    combo = (f'<w:sdt>{_sdt_pr("Vessel", "vessel", "<w:comboBox><w:listItem w:displayText=\"Marguerite\" w:value=\"m\"/><w:listItem w:displayText=\"Thornback\" w:value=\"t\"/></w:comboBox>", 1005)}<w:sdtContent>{r("Thornback")}</w:sdtContent></w:sdt>')
    date = (f'<w:sdt>{_sdt_pr("Arrival", "arrival", "<w:date w:fullDate=\"2026-03-02T00:00:00Z\"><w:dateFormat w:val=\"d MMMM yyyy\"/><w:lid w:val=\"en-GB\"/><w:storeMappedDataAs w:val=\"dateTime\"/><w:calendar w:val=\"gregorian\"/></w:date>", 1006)}<w:sdtContent>{r("2 March 2026")}</w:sdtContent></w:sdt>')
    check = (f'<w:sdt>{_sdt_pr("Inspected", "inspected", "<w14:checkbox><w14:checked w14:val=\"1\"/><w14:checkedState w14:val=\"2612\" w14:font=\"MS Gothic\"/><w14:uncheckedState w14:val=\"2610\" w14:font=\"MS Gothic\"/></w14:checkbox>", 1007)}<w:sdtContent>{r("☒", font="DejaVu Sans")}</w:sdtContent></w:sdt>')
    out.append(p(r("Inline controls: name ") + plain + r(", berth ") + drop + r(", vessel ") + combo + r(", arrival ") + date + r(", inspected ") + check + r("."), style="BodyText"))
    # picture control
    rid, _ = doc.image_rid("circle", 80, 80, (60, 60, 160))
    from ooxml import inline_image
    out.append(p(f'<w:sdt>{_sdt_pr("Logo", "logo", "<w:picture/>", 1008)}<w:sdtContent><w:r>{inline_image(rid, doc.nid(), 80, 80, "logo placeholder")[5:-6]}</w:r></w:sdtContent></w:sdt>'))
    # repeating section with two items, inside a table-less body
    item = lambda n, t: (f'<w:sdt><w:sdtPr><w:id w:val="{2100 + n}"/><w15:repeatingSectionItem/></w:sdtPr><w:sdtContent>'
                         f'{p(r(f"Line item {n}: ") + r(t), style="ListParagraph", ppr=num_ppr())}</w:sdtContent></w:sdt>')
    out.append(f'<w:sdt><w:sdtPr><w:alias w:val="Line items"/><w:tag w:val="items"/><w:id w:val="1009"/><w15:repeatingSection><w15:sectionTitle w:val="Item"/></w15:repeatingSection></w:sdtPr><w:sdtContent>{item(1, "rope, 40 m")}{item(2, "lamp glass, 6 pieces")}{item(3, "tar, 2 barrels")}</w:sdtContent></w:sdt>')
    # row-level SDT in a table and cell-level control; data-bound control + docPartObj
    cell_sdt = f'<w:sdt>{_sdt_pr("Qty", "qty", "<w:text/>", 1010)}<w:sdtContent>{r("12")}</w:sdtContent></w:sdt>'
    out.append(tbl([tr([tc(p(r("Quantity")), w=3000, shade="D9E2F3"), tc(p(cell_sdt), w=3000)]),
                    f'<w:sdt><w:sdtPr><w:id w:val="1011"/><w:tag w:val="rowctl"/></w:sdtPr><w:sdtContent>{tr([tc("Row inside a content control", w=3000), tc("yes", w=3000)])}</w:sdtContent></w:sdt>'], [3000, 3000]))
    bound = (f'<w:sdt><w:sdtPr><w:alias w:val="Title"/><w:tag w:val="doc-title"/><w:id w:val="1012"/><w:dataBinding w:prefixMappings="xmlns:ns0=\'urn:spindle:fixture\'" w:xpath="/ns0:doc[1]/ns0:title[1]" w:storeItemID="{{5B9B6C3E-8F5A-4C7E-9A32-0A1B2C3D4E5F}}"/><w:text/></w:sdtPr><w:sdtContent>{r("Bound title")}</w:sdtContent></w:sdt>')
    out.append(p(r("Data-bound control: ") + bound, style="BodyText"))
    cover = (f'<w:sdt><w:sdtPr><w:id w:val="1013"/><w:docPartObj><w:docPartGallery w:val="Cover Pages"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>{p(r("Cover page building block", b=True, sz=32), style="Title")}{p(r("Wrapped in a docPartObj content control."), style="Subtitle")}</w:sdtContent></w:sdt>')
    out.append(cover)
    # deeply nested
    inner = p(r("Innermost paragraph, 5 controls deep."), style="BodyText")
    for d in range(5):
        inner = f'<w:sdt>{_sdt_pr(f"Level {5 - d}", f"nest{5 - d}", "", 1100 + d)}<w:sdtContent>{inner}</w:sdtContent></w:sdt>'
    out.append(inner)
    # legacy form fields
    def ff(name, body):
        return (f'<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="{name}"/><w:enabled/><w:calcOnExit w:val="0"/>{body}</w:ffData></w:fldChar></w:r>')
    text_ff = ff("Captain", '<w:textInput><w:default w:val="Enter name"/><w:maxLength w:val="30"/></w:textInput>') + '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + r("Ada Marlowe") + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    chk_ff = ff("Cleared", '<w:checkBox><w:sizeAuto/><w:default w:val="0"/><w:checked/></w:checkBox>') + '<w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
    dd_ff = ff("Flag", '<w:ddList><w:result w:val="1"/><w:default w:val="0"/><w:listEntry w:val="Red"/><w:listEntry w:val="Green"/><w:listEntry w:val="Blue"/></w:ddList>') + '<w:r><w:instrText xml:space="preserve"> FORMDROPDOWN </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + r("Green") + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    out.append(p("Legacy (Word 97) form fields:", style="BodyText"))
    out.append(p(r("Captain: ") + text_ff + r("   Cleared: ") + chk_ff + r("   Flag colour: ") + dd_ff, style="BodyText"))
    # custom XML part the control binds to
    doc.parts["customXml/item1.xml"] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><ns0:doc xmlns:ns0="urn:spindle:fixture"><ns0:title>Bound title</ns0:title></ns0:doc>'
    doc.parts["customXml/itemProps1.xml"] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><ds:datastoreItem ds:itemID="{5B9B6C3E-8F5A-4C7E-9A32-0A1B2C3D4E5F}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="urn:spindle:fixture"/></ds:schemaRefs></ds:datastoreItem>'
    doc.parts["customXml/_rels/item1.xml.rels"] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>'
    doc.overrides["/customXml/itemProps1.xml"] = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
    doc.rel("customXml", "../customXml/item1.xml")
    return "".join(out)
