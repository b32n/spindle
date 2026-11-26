"""DrawingML snippets for the PPTX fixture generators: colours, fills, lines, text, shapes and tables as XML strings."""
from pptxbuilder import esc


def rgb(hex_, *mods):
    inner = "".join(f'<a:{m[0]} val="{m[1]}"/>' for m in mods)
    return f'<a:srgbClr val="{hex_}">{inner}</a:srgbClr>' if inner else f'<a:srgbClr val="{hex_}"/>'


def scheme(name, *mods):
    inner = "".join(f'<a:{m[0]} val="{m[1]}"/>' for m in mods)
    return f'<a:schemeClr val="{name}">{inner}</a:schemeClr>' if inner else f'<a:schemeClr val="{name}"/>'


def solid(color):
    return f"<a:solidFill>{color}</a:solidFill>"


def grad(stops, lin=5400000, path=None):
    gs = "".join(f'<a:gs pos="{p}">{c}</a:gs>' for p, c in stops)
    kind = f'<a:path path="{path}"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>' if path else f'<a:lin ang="{lin}" scaled="0"/>'
    return f'<a:gradFill rotWithShape="1"><a:gsLst>{gs}</a:gsLst>{kind}</a:gradFill>'


def patt(prst, fg="4472C4", bg="FFFFFF"):
    return f'<a:pattFill prst="{prst}"><a:fgClr>{rgb(fg)}</a:fgClr><a:bgClr>{rgb(bg)}</a:bgClr></a:pattFill>'


def blip_fill(rid, tile=False, tag="a:blipFill", src=""):
    mode = '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>' if tile else "<a:stretch><a:fillRect/></a:stretch>"
    return f'<{tag}><a:blip r:embed="{rid}"/>{src}{mode}</{tag}>'


def line(width=12700, fill=None, dash=None, head=None, tail=None, cap=None, join="round", cmpd=None):
    attrs = f' w="{width}"' + (f' cap="{cap}"' if cap else "") + (f' cmpd="{cmpd}"' if cmpd else "")
    inner = (fill if fill is not None else solid(rgb("000000"))) + (f'<a:prstDash val="{dash}"/>' if dash else "")
    inner += {"round": "<a:round/>", "bevel": "<a:bevel/>", "miter": '<a:miter lim="800000"/>'}.get(join, "")
    inner += (f'<a:headEnd type="{head[0]}" w="{head[1]}" len="{head[1]}"/>' if head else "") + (f'<a:tailEnd type="{tail[0]}" w="{tail[1]}" len="{tail[1]}"/>' if tail else "")
    return f"<a:ln{attrs}>{inner}</a:ln>"


def rpr(tag="a:rPr", text="", **kw):
    """Run properties: attributes from kw (b, i, u, strike, sz, baseline, cap, spc, lang, kumimoji, altLang) and optional children via `fill`, `font`, `hlink`, `highlight`, `ea`, `cs`."""
    kids = {k: kw.pop(k) for k in ("fill", "font", "hlink", "highlight", "ea", "cs", "ln") if k in kw}
    attrs = "".join(f' {k}="{v}"' for k, v in kw.items())
    inner = kids.get("ln", "") + kids.get("fill", "") + (f'<a:highlight>{kids["highlight"]}</a:highlight>' if "highlight" in kids else "")
    inner += (f'<a:latin typeface="{esc(kids["font"])}"/>' if "font" in kids else "") + (f'<a:ea typeface="{esc(kids["ea"])}"/>' if "ea" in kids else "")
    inner += (f'<a:cs typeface="{esc(kids["cs"])}"/>' if "cs" in kids else "") + kids.get("hlink", "")
    return f"<{tag}{attrs}>{inner}</{tag}>" if inner else f"<{tag}{attrs}/>"


def run(text, **kw):
    kw.setdefault("lang", "en-US")
    return f"<a:r>{rpr(**kw)}<a:t>{esc(text)}</a:t></a:r>"


def para(content="", algn=None, lvl=None, marL=None, indent=None, rtl=None, ppr_children="", end=True, **attrs):
    """A paragraph: `content` is run XML (or plain text, wrapped in one run); ppr_children are pPr child elements (spacing, bullets, tabs)."""
    if content and not content.startswith("<"):
        content = run(content)
    a = "".join(f' {k}="{v}"' for k, v in (("algn", algn), ("lvl", lvl), ("marL", marL), ("indent", indent), ("rtl", rtl)) if v is not None)
    a += "".join(f' {k}="{v}"' for k, v in attrs.items())
    ppr = f"<a:pPr{a}>{ppr_children}</a:pPr>" if ppr_children else (f"<a:pPr{a}/>" if a else "")
    tail = '<a:endParaRPr lang="en-US"/>' if end and not content else ""
    return f"<a:p>{ppr}{content}{tail}</a:p>"


def spacing(before=None, after=None, line_pct=None, line_pts=None):
    out = ""
    if line_pct is not None:
        out += f'<a:lnSpc><a:spcPct val="{line_pct}"/></a:lnSpc>'
    if line_pts is not None:
        out += f'<a:lnSpc><a:spcPts val="{line_pts}"/></a:lnSpc>'
    if before is not None:
        out += f'<a:spcBef><a:spcPts val="{before}"/></a:spcBef>'
    if after is not None:
        out += f'<a:spcAft><a:spcPts val="{after}"/></a:spcAft>'
    return out


def bullet_char(ch="&#8226;", font="Arial"):
    return f'<a:buFont typeface="{font}"/><a:buChar char="{ch}"/>'


def bullet_num(kind="arabicPeriod", start=None):
    return f'<a:buFont typeface="+mj-lt"/><a:buAutoNum type="{kind}"' + (f' startAt="{start}"' if start else "") + "/>"


def tabs(*stops):
    return "<a:tabLst>" + "".join(f'<a:tab pos="{p}" algn="{a}"/>' for p, a in stops) + "</a:tabLst>"


def body_pr(**kw):
    kids = kw.pop("kids", "")
    attrs = "".join(f' {k}="{v}"' for k, v in kw.items())
    return f"<a:bodyPr{attrs}>{kids}</a:bodyPr>" if kids else f"<a:bodyPr{attrs}/>"


def tx_body(paras, bpr=None):
    return f'<p:txBody>{bpr or body_pr(rtlCol="0")}<a:lstStyle/>{paras or "<a:p><a:endParaRPr lang=" + chr(34) + "en-US" + chr(34) + "/></a:p>"}</p:txBody>'


def geom(prst="rect", adj=None):
    av = "".join(f'<a:gd name="{n}" fmla="val {v}"/>' for n, v in (adj or {}).items())
    return f'<a:prstGeom prst="{prst}"><a:avLst>{av}</a:avLst></a:prstGeom>'


def xfrm(x, y, cx, cy, rot=0, flipH=False, flipV=False, tag="a:xfrm"):
    attrs = (f' rot="{rot}"' if rot else "") + (' flipH="1"' if flipH else "") + (' flipV="1"' if flipV else "")
    return f'<{tag}{attrs}><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></{tag}>'


def sp(sid, name, x, y, cx, cy, prst="rect", fill=None, ln=None, paras="", bpr=None, rot=0, flipH=False, flipV=False, txbox=False, adj=None, effects="", custom=None, descr="", style=""):
    """A p:sp with its own position. `fill` / `ln` are fill XML (or None for none given); `custom` replaces the preset geometry."""
    sppr = xfrm(x, y, cx, cy, rot, flipH, flipV) + (custom if custom else geom(prst, adj)) + (fill or "") + (ln or "") + effects
    d = f' descr="{esc(descr)}"' if descr else ""
    tx = tx_body(paras, bpr) if (paras or txbox) else ""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{esc(name)}"{d}/><p:cNvSpPr{" txBox=" + chr(34) + "1" + chr(34) if txbox else ""}/><p:nvPr/></p:nvSpPr>'
            f"<p:spPr>{sppr}</p:spPr>{style}{tx}</p:sp>")


def textbox(sid, name, x, y, cx, cy, paras, bpr=None, **kw):
    kw.setdefault("fill", "<a:noFill/>")
    return sp(sid, name, x, y, cx, cy, paras=paras, bpr=bpr or body_pr(wrap="square", rtlCol="0", kids="<a:noAutofit/>"), txbox=True, **kw)


def pic(sid, name, rid, x, y, cx, cy, crop="", rot=0, flipH=False, flipV=False, descr="", ln="", extra_blip="", link_rid=None):
    d = f' descr="{esc(descr)}"' if descr else ""
    blip = f'<a:blip r:embed="{rid}"/>' if not link_rid else f'<a:blip r:link="{link_rid}"/>'
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{sid}" name="{esc(name)}"{d}/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
            f"<p:blipFill>{blip}{extra_blip}{crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{xfrm(x, y, cx, cy, rot, flipH, flipV)}{geom()}{ln}</p:spPr></p:pic>")


def grp(sid, name, x, y, cx, cy, children, ch=None):
    ox, oy, ecx, ecy = ch or (x, y, cx, cy)
    return (f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{sid}" name="{esc(name)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/>'
            f'<a:chOff x="{ox}" y="{oy}"/><a:chExt cx="{ecx}" cy="{ecy}"/></a:xfrm></p:grpSpPr>{children}</p:grpSp>')


def cxn(sid, name, x, y, cx, cy, prst="straightConnector1", ln=None, st=None, end=None, flipH=False, flipV=False, rot=0):
    """A connector; st / end are (shape id, connection-site index)."""
    c = (f'<a:stCxn id="{st[0]}" idx="{st[1]}"/>' if st else "") + (f'<a:endCxn id="{end[0]}" idx="{end[1]}"/>' if end else "")
    return (f'<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{sid}" name="{esc(name)}"/><p:cNvCxnSpPr>{c}</p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>'
            f"<p:spPr>{xfrm(x, y, cx, cy, rot, flipH, flipV)}{geom(prst)}{ln or line()}</p:spPr></p:cxnSp>")


def frame(sid, name, x, y, cx, cy, graphic):
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{sid}" name="{esc(name)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
            f'{xfrm(x, y, cx, cy, tag="p:xfrm")}<a:graphic>{graphic}</a:graphic></p:graphicFrame>')


def title(sid, text, **kw):
    """A title placeholder inheriting its position from the layout."""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="Title {sid}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>'
            f"<p:txBody><a:bodyPr/><a:lstStyle/>{para(text, **kw)}</p:txBody></p:sp>")


def body_ph(sid, paras, idx=1, typ=None, name=None):
    t = f' type="{typ}"' if typ else ""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name or "Content Placeholder " + str(sid)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph{t} idx="{idx}"/></p:nvPr></p:nvSpPr><p:spPr/>'
            f"<p:txBody><a:bodyPr/><a:lstStyle/>{paras}</p:txBody></p:sp>")


def cust_geom(paths, rect=(0, 0, 0, 0)):
    return (f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="{rect[0]}" t="{rect[1]}" r="r" b="b"/><a:pathLst>{paths}</a:pathLst></a:custGeom>')
