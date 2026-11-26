"""Charts with an embedded workbook, SmartArt, and WordArt, ink, OLE objects, equations and actions."""
import assets as A
import pptxbuilder as X
from dml import body_pr, frame, grad, line, para, rgb, run, scheme, solid, sp, textbox, title

C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
DSP_NS = "http://schemas.microsoft.com/office/drawing/2008/diagram"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
CATS = ["North", "South", "East", "West"]
SERIES = [("2024", [10, 5, 7, 1]), ("2025", [20, 15, 8, 2])]
GUID = lambda n: "{%08X-0000-4000-8000-%012X}" % (0xA0000000 + n, n)


def _cache_str(ref, values):
    pts = "".join(f'<c:pt idx="{i}"><c:v>{X.esc(v)}</c:v></c:pt>' for i, v in enumerate(values))
    return f'<c:strRef><c:f>{ref}</c:f><c:strCache><c:ptCount val="{len(values)}"/>{pts}</c:strCache></c:strRef>'


def _cache_num(ref, values):
    pts = "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(values))
    return f'<c:numRef><c:f>{ref}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{len(values)}"/>{pts}</c:numCache></c:numRef>'


def _ser(kind, i, name, values):
    col = chr(66 + i)
    color = X.COLORS[2 + i][1]
    tx = f"<c:tx>{_cache_str(f'Sheet1!${col}$1', [name])}</c:tx>"
    sppr = (f'<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></c:spPr>' if kind in ("line", "scatter")
            else f'<c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></c:spPr>')
    if kind == "scatter":
        return (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>{tx}{sppr}<c:xVal>{_cache_num("Sheet1!$A$2:$A$5", [1, 2, 3, 4])}</c:xVal>'
                f'<c:yVal>{_cache_num(f"Sheet1!${col}$2:${col}$5", values)}</c:yVal><c:smooth val="0"/></c:ser>')
    cat = f"<c:cat>{_cache_str('Sheet1!$A$2:$A$5', CATS)}</c:cat>"
    val = f'<c:val>{_cache_num(f"Sheet1!${col}$2:${col}$5", values)}</c:val>'
    marker = '<c:marker><c:symbol val="circle"/><c:size val="6"/></c:marker>' if kind == "line" else ""
    smooth = '<c:smooth val="0"/>' if kind == "line" else ""
    return f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>{tx}{sppr}{marker}{cat}{val}{smooth}</c:ser>'


def chart_xml(kind, title_text):
    series = SERIES[:1] if kind == "pie" else SERIES
    sers = "".join(_ser(kind, i, n, v) for i, (n, v) in enumerate(series))
    axes = '<c:axId val="111"/><c:axId val="112"/>'
    plot = {
        "bar": f'<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>{sers}<c:gapWidth val="150"/>{axes}</c:barChart>',
        "line": f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{sers}<c:marker val="1"/>{axes}</c:lineChart>',
        "pie": f'<c:pieChart><c:varyColors val="1"/>{sers}<c:firstSliceAng val="0"/></c:pieChart>',
        "scatter": f'<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>{sers}{axes}</c:scatterChart>',
    }[kind]
    tick = '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
    cat_ax = (f'<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>{tick}<c:crossAx val="112"/><c:crosses val="autoZero"/>'
              '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>')
    x_ax = (f'<c:valAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>{tick}<c:crossAx val="112"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>')
    val_ax = (f'<c:valAx><c:axId val="112"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>{tick}<c:crossAx val="111"/><c:crosses val="autoZero"/>'
              f'<c:crossBetween val="{"midCat" if kind == "scatter" else "between"}"/></c:valAx>')
    axes_xml = "" if kind == "pie" else (x_ax if kind == "scatter" else cat_ax) + val_ax
    return (f'{X.HEAD}<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{X.A}" xmlns:r="{X.R}"><c:date1904 val="0"/><c:roundedCorners val="0"/><c:chart>'
            f'<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{X.esc(title_text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
            f'<c:plotArea><c:layout/>{plot}{axes_xml}</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
            '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>')


def add_chart(deck, slide, kind, sid, x, y, cx, cy, title_text=None):
    """Add a chart part, its embedded workbook and the graphic frame to `slide`; returns nothing (the frame is added)."""
    n = sum(1 for k in deck.extra if k.startswith("ppt/charts/chart") and k.endswith(".xml") and "_rels" not in k) + 1
    rows = [["", *[name for name, _ in SERIES]]] + [[c, *[v[i] for _, v in SERIES]] for i, c in enumerate(CATS)]
    deck.extra[f"ppt/embeddings/Microsoft_Excel_Sheet{n}.xlsx"] = (A.mini_xlsx(rows), None)
    rels = X.Rels()
    rels.add("package", f"../embeddings/Microsoft_Excel_Sheet{n}.xlsx")
    deck.extra[f"ppt/charts/_rels/chart{n}.xml.rels"] = (rels.xml().encode(), None)
    rid = slide.part(f"ppt/charts/chart{n}.xml", chart_xml(kind, title_text or f"{kind.title()} chart"), "chart", "chart")
    slide.add(frame(sid, f"{kind.title()} chart", x, y, cx, cy, f'<a:graphicData uri="{C_NS}"><c:chart xmlns:c="{C_NS}" r:id="{rid}"/></a:graphicData>'))


def charts(path):
    d = X.Deck(title="Charts")
    for kind in ("bar", "line", "pie", "scatter"):
        s = d.slide(3)
        s.add(title(2, f"{kind.title()} chart with an embedded workbook"))
        add_chart(d, s, kind, 3, 1500000, 1500000, 8000000, 4300000)
    d.write(path)


# -- SmartArt -------------------------------------------------------------------------------------------------------------

def _text_pt(i, text):
    return (f'<dgm:pt modelId="{GUID(i)}"><dgm:prSet phldrT="[Text]"/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></dgm:t></dgm:pt>')


def _trans_pt(i, kind, cxn):
    return f'<dgm:pt modelId="{GUID(i)}" type="{kind}" cxnId="{GUID(cxn)}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>'


LAYOUT_ID = "urn:microsoft.com/office/officeart/2005/8/layout/default"
STYLE_ID = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1"
COLORS_ID = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2"
NODES = ["Plan", "Build", "Review"]


def smartart(path):
    d = X.Deck(title="SmartArt")
    s = d.slide(3)
    s.add(title(2, "Basic block list"))
    pts = (f'<dgm:pt modelId="{GUID(1)}" type="doc"><dgm:prSet loTypeId="{LAYOUT_ID}" loCatId="list" qsTypeId="{STYLE_ID}" qsCatId="simple" csTypeId="{COLORS_ID}" csCatId="accent1" phldr="0"/>'
           '<dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>')
    cxns = ""
    for n, text in enumerate(NODES):
        node, par, sib, cx = 10 + n, 20 + n, 30 + n, 40 + n
        pts += _text_pt(node, text) + _trans_pt(par, "parTrans", cx) + _trans_pt(sib, "sibTrans", cx)
        cxns += f'<dgm:cxn modelId="{GUID(cx)}" srcId="{GUID(1)}" destId="{GUID(node)}" srcOrd="{n}" destOrd="0" parTransId="{GUID(par)}" sibTransId="{GUID(sib)}"/>'
    node_layout = ('<dgm:layoutNode name="node"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"/><dgm:shape type="rect" r:blip=""><dgm:adjLst/></dgm:shape>'
                   '<dgm:presOf axis="desOrSelf" ptType="node"/><dgm:constrLst><dgm:constr type="h" refType="w" fact="0.6"/></dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode>')
    layout = (f'{X.HEAD}<dgm:layoutDef xmlns:dgm="{DGM_NS}" xmlns:a="{X.A}" xmlns:r="{X.R}" uniqueId="{LAYOUT_ID}"><dgm:title val=""/><dgm:desc val=""/>'
              '<dgm:catLst><dgm:cat type="list" pri="400"/></dgm:catLst><dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst>'
              '<dgm:alg type="snake"><dgm:param type="grDir" val="tL"/><dgm:param type="flowDir" val="row"/><dgm:param type="contDir" val="sameDir"/><dgm:param type="off" val="ctr"/></dgm:alg>'
              '<dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst><dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="0.6"/>'
              '<dgm:constr type="w" for="ch" forName="sibTrans" refType="w" refFor="ch" refForName="node" fact="0.1"/></dgm:constrLst><dgm:ruleLst/>'
              f'<dgm:forEach name="Name0" axis="ch" ptType="node">{node_layout}<dgm:forEach name="Name1" axis="followSib" ptType="sibTrans" cnt="1"><dgm:layoutNode name="sibTrans"><dgm:alg type="sp"/><dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:forEach></dgm:forEach>'
              '</dgm:layoutNode></dgm:layoutDef>')
    style = (f'{X.HEAD}<dgm:styleDef xmlns:dgm="{DGM_NS}" xmlns:a="{X.A}" uniqueId="{STYLE_ID}"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="simple" pri="10100"/></dgm:catLst>'
             '<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:styleLbl name="node1"><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:sp3d/><dgm:txPr/>'
             '<dgm:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dgm:style></dgm:styleLbl></dgm:styleDef>')
    colors = (f'{X.HEAD}<dgm:colorsDef xmlns:dgm="{DGM_NS}" xmlns:a="{X.A}" uniqueId="{COLORS_ID}"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="accent1" pri="11200"/></dgm:catLst>'
              '<dgm:styleLbl name="node1"><dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:fillClrLst><dgm:linClrLst meth="repeat"><a:schemeClr val="lt1"/></dgm:linClrLst>'
              '<dgm:effectClrLst/><dgm:txLinClrLst/><dgm:txFillClrLst/><dgm:txEffectClrLst/></dgm:styleLbl></dgm:colorsDef>')
    shapes = ""
    for n, text in enumerate(NODES):
        x, y, w, h = n * 2900000, 0, 2600000, 1560000
        shapes += (f'<dsp:sp modelId="{GUID(10 + n)}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
                   '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></dsp:spPr>'
                   '<dsp:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style>'
                   f'<dsp:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="2400"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:rPr><a:t>{text}</a:t></a:r></a:p></dsp:txBody>'
                   f'<dsp:txXfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></dsp:txXfrm></dsp:sp>')
    drawing = (f'{X.HEAD}<dsp:drawing xmlns:dgm="{DGM_NS}" xmlns:dsp="{DSP_NS}" xmlns:a="{X.A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>{shapes}</dsp:spTree></dsp:drawing>')
    dr = s.part("ppt/diagrams/drawing1.xml", drawing, "diagramDrawing", "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing")
    data = (f'{X.HEAD}<dgm:dataModel xmlns:dgm="{DGM_NS}" xmlns:a="{X.A}"><dgm:ptLst>{pts}</dgm:ptLst><dgm:cxnLst>{cxns}</dgm:cxnLst><dgm:bg/><dgm:whole/>'
            f'<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="{DSP_NS}" relId="{dr}" minVer="{DGM_NS}"/></a:ext></dgm:extLst></dgm:dataModel>')
    dm = s.part("ppt/diagrams/data1.xml", data, "diagramData", "diagramData")
    lo = s.part("ppt/diagrams/layout1.xml", layout, "diagramLayout", "diagramLayout")
    qs = s.part("ppt/diagrams/quickStyle1.xml", style, "diagramStyle", "diagramQuickStyle")
    cs = s.part("ppt/diagrams/colors1.xml", colors, "diagramColors", "diagramColors")
    s.add(frame(3, "Diagram", 1000000, 1800000, 8700000, 1560000,
                f'<a:graphicData uri="{DGM_NS}"><dgm:relIds xmlns:dgm="{DGM_NS}" r:dm="{dm}" r:lo="{lo}" r:qs="{qs}" r:cs="{cs}"/></a:graphicData>'))
    d.write(path)


# -- WordArt, ink, OLE, equations, actions --------------------------------------------------------------------------------

INK = ('<?xml version="1.0" encoding="UTF-8"?>\n<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:definitions><inkml:context xml:id="ctx0"><inkml:inkSource xml:id="inkSrc0">'
       '<inkml:traceFormat><inkml:channel name="X" type="integer" max="3840" units="cm"/><inkml:channel name="Y" type="integer" max="2160" units="cm"/></inkml:traceFormat>'
       '<inkml:channelProperties><inkml:channelProperty channel="X" name="resolution" value="100" units="1/cm"/><inkml:channelProperty channel="Y" name="resolution" value="100" units="1/cm"/></inkml:channelProperties></inkml:inkSource></inkml:context>'
       '<inkml:brush xml:id="br0"><inkml:brushProperty name="width" value="0.05" units="cm"/><inkml:brushProperty name="height" value="0.05" units="cm"/><inkml:brushProperty name="color" value="#E71224"/></inkml:brush></inkml:definitions>'
       '<inkml:trace xml:id="tr0" contextRef="#ctx0" brushRef="#br0">100 100, 180 160, 260 100, 340 160, 420 100</inkml:trace></inkml:ink>')


def objects(path):
    d = X.Deck(title="WordArt, ink, OLE objects and equations")
    s = d.slide(3)
    s.add(title(2, "WordArt"))
    warp = lambda prst: body_pr(wrap="none", rtlCol="0", kids=f'<a:prstTxWarp prst="{prst}"><a:avLst/></a:prstTxWarp><a:noAutofit/>')
    fancy = run("Spindle", sz=6000, b=1, fill=grad([(0, rgb("C00000")), (100000, rgb("FFC000"))], lin=0), ln=line(12700, solid(rgb("7F6000"))))
    s.add(sp(3, "Arch WordArt", 700000, 1500000, 5000000, 2200000, fill="<a:noFill/>", paras=para(fancy, algn="ctr"), bpr=warp("textArchUp"), txbox=True))
    s.add(sp(4, "Wave WordArt", 6200000, 1500000, 5000000, 2200000, fill="<a:noFill/>", paras=para(run("Fixture", sz=6000, b=1, fill=solid(scheme("accent1"))), algn="ctr"), bpr=warp("textWave1"), txbox=True))
    s = d.slide(3)
    s.add(title(2, "Ink and equation"))
    ink_rid = s.part("ppt/ink/ink1.xml", INK, "application/inkml+xml", "customXml")
    s.add(f'<mc:AlternateContent xmlns:mc="{MC_NS}"><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">'
          f'<p:contentPart bwMode="auto" r:id="{ink_rid}"><p14:nvContentPartPr><p14:cNvPr id="3" name="Ink 2"/><p14:cNvContentPartPr/><p14:nvPr/></p14:nvContentPartPr>'
          '<p14:xfrm><a:off x="700000" y="1500000"/><a:ext cx="3000000" cy="900000"/></p14:xfrm></p:contentPart></mc:Choice></mc:AlternateContent>')
    eq = ('<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>x=</m:t></m:r><m:f><m:fPr><m:ctrlPr><a:rPr lang="en-US"/></m:ctrlPr></m:fPr>'
          '<m:num><m:r><m:t>-b±</m:t></m:r><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:sSup><m:e><m:r><m:t>b</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t>-4ac</m:t></m:r></m:e></m:rad></m:num>'
          '<m:den><m:r><m:t>2a</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara>')
    shape = lambda body: (f'<p:sp><p:nvSpPr><p:cNvPr id="4" name="Equation"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="4500000" y="1500000"/><a:ext cx="5000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
                          f'<p:txBody><a:bodyPr wrap="square" rtlCol="0"/><a:lstStyle/><a:p>{body}</a:p></p:txBody></p:sp>')
    s.add(f'<mc:AlternateContent xmlns:mc="{MC_NS}"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">{shape(f"<a14:m>{eq}</a14:m>")}</mc:Choice>'
          f'<mc:Fallback>{shape(run("x = (-b ± sqrt(b^2 - 4ac)) / 2a"))}</mc:Fallback></mc:AlternateContent>')
    s = d.slide(3)
    s.add(title(2, "Embedded object and actions"))
    s.media("ole-preview.png", X.png(48, 32, fg=(0, 128, 64)))
    prev = s.rels.items[0][0]
    ole_rid = s.part("ppt/embeddings/oleObject1.bin", A.opaque_blob("an embedded OLE object"), None, "oleObject")
    s.add(frame(3, "Embedded worksheet", 700000, 1500000, 2400000, 1600000,
                f'<a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="Worksheet" r:id="{ole_rid}" imgW="2400000" imgH="1600000" progId="Excel.Sheet.12"><p:embed/>'
                f'<p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{prev}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
                '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2400000" cy="1600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData>'))
    s.add(sp(4, "Next slide button", 4000000, 1500000, 2400000, 700000, prst="roundRect", fill=solid(scheme("accent1")), paras=para("Go to the next slide", algn="ctr"),
             hlink='<a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=nextslide"/>'))
    s.add(sp(5, "Macro button", 4000000, 2500000, 2400000, 700000, prst="roundRect", fill=solid(scheme("accent2")), paras=para("Run a macro (never run)", algn="ctr"),
             hlink='<a:hlinkClick r:id="" action="ppaction://macro?name=Module1.Hello"/>'))
    d.write(path)
