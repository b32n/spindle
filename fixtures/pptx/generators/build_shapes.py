"""Shapes: every preset geometry, custom geometry, fills, lines and effects."""
import pptxbuilder as X
from dml import (blip_fill, body_pr, cust_geom, grad, line, para, patt, rgb, run, scheme, solid, sp, textbox, title, xfrm)

# ST_ShapeType (ECMA-376 20.1.10.56): the full list of preset geometries
PRESETS = """line lineInv triangle rtTriangle rect diamond parallelogram trapezoid nonIsoscelesTrapezoid pentagon hexagon heptagon octagon decagon dodecagon
star4 star5 star6 star7 star8 star10 star12 star16 star24 star32 roundRect round1Rect round2SameRect round2DiagRect snipRoundRect snip1Rect snip2SameRect snip2DiagRect
plaque ellipse teardrop homePlate chevron pieWedge pie blockArc donut noSmoking rightArrow leftArrow upArrow downArrow stripedRightArrow notchedRightArrow bentUpArrow
leftRightArrow upDownArrow leftUpArrow leftRightUpArrow quadArrow leftArrowCallout rightArrowCallout upArrowCallout downArrowCallout leftRightArrowCallout
upDownArrowCallout quadArrowCallout bentArrow uturnArrow circularArrow leftCircularArrow leftRightCircularArrow curvedRightArrow curvedLeftArrow curvedUpArrow
curvedDownArrow swooshArrow cube can lightningBolt heart sun moon smileyFace irregularSeal1 irregularSeal2 foldedCorner bevel frame halfFrame corner diagStripe chord arc
leftBracket rightBracket leftBrace rightBrace bracketPair bracePair straightConnector1 bentConnector2 bentConnector3 bentConnector4 bentConnector5 curvedConnector2
curvedConnector3 curvedConnector4 curvedConnector5 callout1 callout2 callout3 accentCallout1 accentCallout2 accentCallout3 borderCallout1 borderCallout2 borderCallout3
accentBorderCallout1 accentBorderCallout2 accentBorderCallout3 wedgeRectCallout wedgeRoundRectCallout wedgeEllipseCallout cloudCallout cloud ribbon ribbon2 ellipseRibbon
ellipseRibbon2 leftRightRibbon verticalScroll horizontalScroll wave doubleWave plus flowChartProcess flowChartAlternateProcess flowChartDecision flowChartInputOutput
flowChartPredefinedProcess flowChartInternalStorage flowChartDocument flowChartMultidocument flowChartTerminator flowChartPreparation flowChartManualInput
flowChartManualOperation flowChartConnector flowChartOffpageConnector flowChartPunchedCard flowChartPunchedTape flowChartSummingJunction flowChartOr flowChartCollate
flowChartSort flowChartExtract flowChartMerge flowChartOnlineStorage flowChartOfflineStorage flowChartDelay flowChartMagneticTape flowChartMagneticDisk flowChartMagneticDrum flowChartDisplay
actionButtonBlank actionButtonHome actionButtonHelp actionButtonInformation actionButtonForwardNext actionButtonBackPrevious actionButtonEnd actionButtonBeginning
actionButtonReturn actionButtonDocument actionButtonSound actionButtonMovie gear6 gear9 funnel mathPlus mathMinus mathMultiply mathDivide mathEqual mathNotEqual
cornerTabs squareTabs plaqueTabs chartX chartStar chartPlus""".split()
PER_SLIDE = 24


def shapes_presets(path):
    assert len(PRESETS) == len(set(PRESETS))
    d = X.Deck(title="Shape presets")
    for start in range(0, len(PRESETS), PER_SLIDE):
        s = d.slide(4)
        s.add(title(2, f"Presets {start + 1} to {min(start + PER_SLIDE, len(PRESETS))}"))
        for n, prst in enumerate(PRESETS[start:start + PER_SLIDE]):
            col, row = n % 6, n // 6
            x, y = 450000 + col * 1850000, 1350000 + row * 1350000
            s.add(sp(3 + n, prst, x, y, 1100000, 800000, prst=prst, fill=solid(scheme("accent1")), ln=line(12700, solid(scheme("accent1", ("lumMod", 50000))))))
            s.add(textbox(100 + n, f"Label {prst}", x - 250000, y + 830000, 1600000, 300000, para(run(prst, sz=900), algn="ctr"), bpr=body_pr(wrap="square", lIns=0, rIns=0, tIns=0, bIns=0, kids="<a:noAutofit/>")))
    s = d.slide(4)
    s.add(title(2, "Adjustments, flips, rotation and text"))
    cases = [("roundRect", {"adj": 40000}), ("star5", {"adj": 20000, "hf": 105146, "vf": 110557}), ("rightArrow", {"adj1": 30000, "adj2": 80000}),
             ("wedgeRectCallout", {"adj1": -20833, "adj2": 100000}), ("pie", {"adj1": 0, "adj2": 16200000}), ("chevron", {"adj": 80000})]
    for n, (prst, adj) in enumerate(cases):
        s.add(sp(3 + n, f"{prst} adjusted", 450000 + n * 1900000, 1400000, 1500000, 1000000, prst=prst, adj=adj, fill=solid(scheme("accent2")), paras=para("adj", algn="ctr")))
    s.add(sp(10, "Rotated 30 degrees", 450000, 3000000, 1800000, 900000, rot=1800000, fill=solid(scheme("accent3")), paras=para("rot 30", algn="ctr")))
    s.add(sp(11, "Flipped horizontally", 2900000, 3000000, 1800000, 900000, prst="rightArrow", flipH=True, fill=solid(scheme("accent4")), paras=para("flipH", algn="ctr")))
    s.add(sp(12, "Flipped vertically", 5300000, 3000000, 1800000, 900000, prst="rtTriangle", flipV=True, fill=solid(scheme("accent5"))))
    s.add(sp(13, "Rotated and flipped", 7700000, 3000000, 1800000, 900000, prst="homePlate", rot=19800000, flipH=True, fill=solid(scheme("accent6"))))
    s.add(textbox(14, "Plain text box", 450000, 4400000, 4200000, 700000, para("A text box has no geometry fill.")))
    s.add(sp(15, "Text in a shape", 5000000, 4400000, 4200000, 900000, prst="ellipse", fill=solid(scheme("accent1")), paras=para(run("Text inside an ellipse", fill=solid(scheme("bg1"))), algn="ctr")))
    d.write(path)


def shapes_custgeom(path):
    d = X.Deck(title="Custom geometry")
    s = d.slide(3)
    s.add(title(2, "Freeform shapes"))
    tri = '<a:path w="100" h="100"><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="50" y="0"/></a:lnTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path>'
    s.add(sp(3, "Triangle path", 600000, 1500000, 1800000, 1500000, custom=cust_geom(tri), fill=solid(scheme("accent1"))))
    bez = ('<a:path w="200" h="100"><a:moveTo><a:pt x="0" y="50"/></a:moveTo><a:cubicBezTo><a:pt x="40" y="-20"/><a:pt x="80" y="-20"/><a:pt x="100" y="50"/></a:cubicBezTo>'
           '<a:quadBezTo><a:pt x="150" y="120"/><a:pt x="200" y="50"/></a:quadBezTo><a:lnTo><a:pt x="200" y="100"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/></a:path>')
    s.add(sp(4, "Bezier curves", 2800000, 1500000, 2400000, 1200000, custom=cust_geom(bez), fill=solid(scheme("accent2"))))
    arc = ('<a:path w="100" h="100" fill="none"><a:moveTo><a:pt x="0" y="50"/></a:moveTo><a:arcTo wR="50" hR="50" stAng="10800000" swAng="10800000"/></a:path>'
           '<a:path w="100" h="100" stroke="0"><a:moveTo><a:pt x="25" y="50"/></a:moveTo><a:arcTo wR="25" hR="25" stAng="10800000" swAng="10800000"/><a:close/></a:path>')
    s.add(sp(5, "Arcs, two paths", 5600000, 1500000, 1800000, 1500000, custom=cust_geom(arc), fill=solid(scheme("accent3")), ln=line(19050)))
    guides = ('<a:custGeom><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst><a:gdLst><a:gd name="x1" fmla="*/ w adj 100000"/><a:gd name="y1" fmla="*/ h 1 2"/></a:gdLst><a:ahLst/>'
              '<a:cxnLst><a:cxn ang="0"><a:pos x="x1" y="y1"/></a:cxn><a:cxn ang="5400000"><a:pos x="hc" y="b"/></a:cxn></a:cxnLst><a:rect l="x1" t="t" r="r" b="b"/>'
              '<a:pathLst><a:path><a:moveTo><a:pt x="x1" y="t"/></a:moveTo><a:lnTo><a:pt x="r" y="t"/></a:lnTo><a:lnTo><a:pt x="r" y="b"/></a:lnTo><a:lnTo><a:pt x="l" y="y1"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>')
    s.add(sp(6, "Guides and connection sites", 7800000, 1500000, 2400000, 1500000, custom=guides, fill=solid(scheme("accent4")), paras=para("text rect", algn="ctr")))
    star = "".join(f'<a:lnTo><a:pt x="{x}" y="{y}"/></a:lnTo>' for x, y in ((62, 38), (98, 38), (69, 60), (80, 95), (50, 73), (20, 95), (31, 60), (2, 38), (38, 38)))
    s.add(sp(7, "Hand-drawn star", 600000, 3500000, 1800000, 1800000, custom=cust_geom(f'<a:path w="100" h="100"><a:moveTo><a:pt x="50" y="0"/></a:moveTo>{star}<a:close/></a:path>'), fill=solid(scheme("accent5"))))
    d.write(path)


def fills_lines(path):
    d = X.Deck(title="Fills and lines")
    s = d.slide(3)
    s.add(title(2, "Fills"))
    rid = s.media("tile.png", X.png(24, 24, fg=(200, 60, 60)))
    fills = [("Solid RGB", solid(rgb("4472C4"))), ("Solid scheme, darker", solid(scheme("accent2", ("lumMod", 75000)))), ("Linear gradient", grad([(0, rgb("FFFFFF")), (100000, rgb("4472C4"))], lin=2700000)),
             ("Three stop gradient", grad([(0, rgb("C00000")), (50000, rgb("FFC000")), (100000, rgb("00B050"))], lin=0)),
             ("Circle gradient", grad([(0, rgb("FFFFFF")), (100000, rgb("7030A0"))], path="circle")), ("Rectangle gradient", grad([(0, rgb("FFFFFF")), (100000, rgb("0070C0"))], path="rect")),
             ("Pattern dkDnDiag", patt("dkDnDiag")), ("Pattern pct50", patt("pct50", "C00000")), ("Pattern smCheck", patt("smCheck", "00B050", "FFF2CC")),
             ("Picture stretch", blip_fill(rid, tag="a:blipFill")), ("Picture tile", blip_fill(rid, tile=True)), ("No fill", "<a:noFill/>"),
             ("Alpha 50 percent", solid(rgb("C00000", ("alpha", 50000)))), ("Alpha 20 percent", solid(scheme("accent1", ("alpha", 20000))))]
    for n, (label, fill) in enumerate(fills):
        s.add(sp(3 + n, label, 500000 + (n % 5) * 2300000, 1400000 + (n // 5) * 1500000, 2000000, 1200000, fill=fill, ln=line(9525), paras=para(run(label, sz=1100), algn="ctr")))
    s2 = d.slide(3)
    s2.add(title(2, "Lines"))
    dashes = ["solid", "dot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot", "sysDash", "sysDot", "sysDashDot", "sysDashDotDot"]
    for n, dash in enumerate(dashes):
        s2.add(sp(3 + n, f"Dash {dash}", 500000, 1300000 + n * 300000, 3000000, 0, prst="line", ln=line(28575, dash=dash)))
    arrows = [("triangle", "med"), ("stealth", "lg"), ("diamond", "med"), ("oval", "sm"), ("arrow", "med")]
    for n, (a, size) in enumerate(arrows):
        s2.add(sp(20 + n, f"Arrow {a}", 4000000, 1400000 + n * 450000, 3000000, 0, prst="line", ln=line(28575, head=("none", "med"), tail=(a, size))))
        s2.add(sp(30 + n, f"Head arrow {a}", 4000000, 1400000 + n * 450000 + 200000, 3000000, 0, prst="line", ln=line(12700, head=(a, size))))
    for n, (w, cap, join) in enumerate(((0, "flat", "round"), (12700, "rnd", "round"), (76200, "sq", "bevel"), (152400, "flat", "miter"))):
        s2.add(sp(40 + n, f"Width {w} cap {cap}", 7500000, 1400000 + n * 500000, 2400000, 0, prst="line", ln=line(w, cap=cap, join=join)))
    for n, cmpd in enumerate(("dbl", "thickThin", "thinThick", "tri")):
        s2.add(sp(50 + n, f"Compound {cmpd}", 7500000, 3600000 + n * 400000, 2400000, 0, prst="line", ln=line(76200, cmpd=cmpd)))
    s2.add(sp(60, "Gradient line", 500000, 4800000, 3000000, 900000, prst="rect", fill="<a:noFill/>", ln=line(76200, grad([(0, rgb("C00000")), (100000, rgb("0070C0"))]))))
    s2.add(sp(61, "Elbow connector line", 4000000, 4800000, 3000000, 900000, prst="bentConnector3", fill="<a:noFill/>", ln=line(25400, tail=("triangle", "med"))))
    s3 = d.slide(3)
    s3.bg = "<p:bgPr>" + grad([(0, rgb("DEEAF6")), (100000, rgb("9DC3E6"))], lin=5400000) + "<a:effectLst/></p:bgPr>"
    s3.add(title(2, "Gradient slide background"))
    s4 = d.slide(3)
    bg = s4.media("background.png", X.png(32, 32, fg=(120, 160, 220), bg=(230, 240, 250)))
    s4.bg = "<p:bgPr>" + blip_fill(bg, tile=True) + "<a:effectLst/></p:bgPr>"
    s4.add(title(2, "Picture slide background"))
    s5 = d.slide(3)
    s5.bg = "<p:bgPr>" + solid(rgb("FFF2CC")) + "<a:effectLst/></p:bgPr>"
    s5.add(title(2, "Solid slide background"))
    d.write(path)


def effects(path):
    d = X.Deck(title="Effects")
    s = d.slide(3)
    s.add(title(2, "Shadows, glow, reflection, soft edges"))
    base = dict(fill=solid(scheme("accent1")))
    shdw = {
        "Outer shadow": '<a:effectLst><a:outerShdw blurRad="63500" dist="76200" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="45000"/></a:srgbClr></a:outerShdw></a:effectLst>',
        "Inner shadow": '<a:effectLst><a:innerShdw blurRad="63500" dist="50800" dir="13500000"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:innerShdw></a:effectLst>',
        "Preset shadow": '<a:effectLst><a:prstShdw prst="shdw5" dist="50800" dir="2700000"><a:srgbClr val="000000"/></a:prstShdw></a:effectLst>',
        "Glow": '<a:effectLst><a:glow rad="101600"><a:srgbClr val="FFC000"><a:alpha val="60000"/></a:srgbClr></a:glow></a:effectLst>',
        "Reflection": '<a:effectLst><a:reflection blurRad="6350" stA="50000" endA="300" endPos="35000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/></a:effectLst>',
        "Soft edges": '<a:effectLst><a:softEdge rad="63500"/></a:effectLst>',
        "Blur": '<a:effectLst><a:blur rad="38100" grow="1"/></a:effectLst>',
        "Shadow and glow": '<a:effectLst><a:glow rad="63500"><a:srgbClr val="00B0F0"><a:alpha val="40000"/></a:srgbClr></a:glow><a:outerShdw blurRad="50800" dist="38100" dir="5400000"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>',
    }
    for n, (label, eff) in enumerate(shdw.items()):
        s.add(sp(3 + n, label, 600000 + (n % 4) * 2800000, 1500000 + (n // 4) * 1900000, 2000000, 1100000, effects=eff, paras=para(label, algn="ctr"), **base))
    s2 = d.slide(3)
    s2.add(title(2, "Three-dimensional formatting"))
    scene = '<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="18900000" lon="19800000" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>'
    sp3d = '<a:sp3d extrusionH="152400" contourW="12700" prstMaterial="plastic"><a:bevelT w="101600" h="101600" prst="circle"/><a:bevelB w="50800" h="50800" prst="angle"/><a:extrusionClr><a:srgbClr val="2F5597"/></a:extrusionClr></a:sp3d>'
    s2.add(sp(3, "Extruded and bevelled", 800000, 1600000, 2600000, 1600000, effects=scene + sp3d, paras=para("3-D shape", algn="ctr"), **base))
    s2.add(sp(4, "Rotated in 3-D", 4200000, 1600000, 2600000, 1600000, effects='<a:scene3d><a:camera prst="perspectiveRelaxed"><a:rot lat="1200000" lon="20400000" rev="0"/></a:camera><a:lightRig rig="balanced" dir="t"/></a:scene3d><a:sp3d/>',
             paras=para("Perspective", algn="ctr"), **base))
    s2.add(sp(5, "Text with shadow", 7600000, 1600000, 3600000, 1600000, fill="<a:noFill/>", ln=line(9525),
              paras=para(run("Shadowed text", sz=3200, fill=solid(scheme("accent2")), ln="") .replace("</a:rPr>", '<a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000"><a:srgbClr val="000000"><a:alpha val="43000"/></a:srgbClr></a:outerShdw></a:effectLst></a:rPr>'))))
    s2.add(sp(6, "Style reference effect", 800000, 3800000, 2600000, 1400000, fill=solid(scheme("accent3")),
              style='<p:style><a:lnRef idx="2"><a:schemeClr val="accent3"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent3"/></a:fillRef><a:effectRef idx="3"><a:schemeClr val="accent3"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>',
              paras=para("effectRef idx 3", algn="ctr")))
    d.write(path)
