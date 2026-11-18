"""Tiny OOXML (WordprocessingML) package builder used by the fixture generators.

Everything the corpus contains is authored here or in the build_*.py scripts that
import this module; there is no third-party document content. Output is
deterministic: fixed zip timestamps, fixed core-properties dates, no randomness.
"""
import io
import struct
import zlib
import zipfile
from xml.sax.saxutils import escape as _esc

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "w10": "urn:schemas-microsoft-com:office:word",
}
NS_DECL = " ".join(f'xmlns:{k}="{v}"' for k, v in NS.items()) + ' mc:Ignorable="w14 w15"'

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT_MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
CT = {
    "styles": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    "numbering": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    "settings": "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
    "fontTable": "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml",
    "theme": "application/vnd.openxmlformats-officedocument.theme+xml",
    "footnotes": "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
    "endnotes": "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
    "comments": "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    "commentsExtended": "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
    "header": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    "footer": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    "chart": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
}


def esc(s):
    return _esc(s, {'"': "&quot;"})


# ---------------------------------------------------------------- text helpers
_SENTENCES = [
    "The harbour committee met on the first Tuesday of every month to review the tide tables.",
    "Each lantern on the quay was numbered, oiled and logged by the keeper before dusk.",
    "A careful reader will notice that the second ledger disagrees with the first by seven coins.",
    "Wind, salt and patience shaped the town more than any decree ever did.",
    "Nobody remembered who had painted the clock face blue, but everybody agreed it suited the square.",
    "Measurements were taken twice, once at dawn and once after the fishing boats returned.",
    "The orchard on the northern slope produced more pears than the town could eat or trade.",
    "Rain arrived in short, orderly bursts, as if it had been scheduled by the same committee.",
    "In the margin of the survey, someone had sketched a heron standing on one leg.",
    "Letters addressed to the old mill were still delivered, although the mill had burned down years ago.",
    "Bread was baked in three sizes: small for travellers, medium for families, large for feast days.",
    "The bridge keeper kept a tally of crossings using a string of coloured beads.",
    "Every spring the children repainted the boundary stones, each in a different colour.",
    "Maps of the estuary changed so often that the cartographer stopped signing them.",
    "A single bell, rung at noon, was enough to pause the market for exactly one breath.",
    "The apprentice learned that a good index is simply a promise kept in alphabetical order.",
    "Winter inventories listed rope, tar, lamp glass, and a surprising quantity of buttons.",
    "Visitors were invited to leave a pebble at the lighthouse, and the pile grew into a small hill.",
]


def prose(n=3, seed=0):
    """n deterministic original sentences joined into a paragraph."""
    return " ".join(_SENTENCES[(seed + i * 5) % len(_SENTENCES)] for i in range(n))


# ----------------------------------------------------------------- run / para
def rpr(**kw):
    """Run properties from keywords: b,i,u,strike,color,sz(half-points),font,style,vanish,caps,smallCaps,highlight,lang,rtl."""
    out = []
    if kw.get("style"):
        out.append(f'<w:rStyle w:val="{kw["style"]}"/>')
    if kw.get("font"):
        f = kw["font"]
        out.append(f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}" w:eastAsia="{kw.get("eastAsia", f)}"/>')
    for k, tag in (("b", "b"), ("i", "i"), ("strike", "strike"), ("caps", "caps"), ("smallCaps", "smallCaps"), ("vanish", "vanish")):
        if kw.get(k):
            out.append(f"<w:{tag}/>")
    if kw.get("color"):
        out.append(f'<w:color w:val="{kw["color"]}"/>')
    if kw.get("sz"):
        out.append(f'<w:sz w:val="{kw["sz"]}"/><w:szCs w:val="{kw["sz"]}"/>')
    if kw.get("highlight"):
        out.append(f'<w:highlight w:val="{kw["highlight"]}"/>')
    if kw.get("u"):
        out.append('<w:u w:val="single"/>')
    if kw.get("vert"):
        out.append(f'<w:vertAlign w:val="{kw["vert"]}"/>')
    if kw.get("rtl"):
        out.append("<w:rtl/>")
    if kw.get("lang"):
        out.append(f'<w:lang w:val="{kw["lang"]}" w:bidi="{kw.get("bidi", kw["lang"])}" w:eastAsia="{kw.get("ea", kw["lang"])}"/>')
    return f"<w:rPr>{''.join(out)}</w:rPr>" if out else ""


def r(text, **kw):
    return f'<w:r>{rpr(**kw)}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def tab():
    return "<w:r><w:tab/></w:r>"


def br(kind=None):
    return f'<w:r><w:br w:type="{kind}"/></w:r>' if kind else "<w:r><w:br/></w:r>"


def p(content="", style=None, ppr="", **rkw):
    """Paragraph. content: plain str (becomes one run) or pre-built run XML (starts with '<')."""
    if isinstance(content, str) and not content.startswith("<"):
        content = r(content, **rkw) if content else ""
    st = f'<w:pStyle w:val="{style}"/>' if style else ""
    pp = f"<w:pPr>{st}{ppr}</w:pPr>" if (st or ppr) else ""
    return f"<w:p>{pp}{content}</w:p>"


def num(num_id, ilvl=0):
    return f'<w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr>'


def fld(instr, result="", **kw):
    """Complex field: begin / instrText / separate / cached result / end."""
    rp = rpr(**kw)
    return (
        f'<w:r>{rp}<w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r>{rp}<w:instrText xml:space="preserve"> {esc(instr)} </w:instrText></w:r>'
        f'<w:r>{rp}<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r>{rp}<w:t xml:space="preserve">{esc(result)}</w:t></w:r>'
        f'<w:r>{rp}<w:fldChar w:fldCharType="end"/></w:r>'
    )


def bookmark(bid, name, inner):
    return f'<w:bookmarkStart w:id="{bid}" w:name="{name}"/>{inner}<w:bookmarkEnd w:id="{bid}"/>'


def hyperlink_ext(rid, text, **kw):
    return f'<w:hyperlink r:id="{rid}" w:history="1">{r(text, style="Hyperlink", **kw)}</w:hyperlink>'


def hyperlink_int(anchor, text):
    return f'<w:hyperlink w:anchor="{anchor}" w:history="1">{r(text, style="Hyperlink")}</w:hyperlink>'


# ----------------------------------------------------------------------- tables
def tc(content, w=None, span=None, vmerge=None, shade=None, borders="", valign=None, extra=""):
    pr = ""
    if w:
        pr += f'<w:tcW w:w="{w}" w:type="dxa"/>'
    if span:
        pr += f'<w:gridSpan w:val="{span}"/>'
    if vmerge is not None:
        pr += '<w:vMerge w:val="restart"/>' if vmerge == "restart" else "<w:vMerge/>"
    if borders:
        pr += f"<w:tcBorders>{borders}</w:tcBorders>"
    if shade:
        pr += f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>'
    if valign:
        pr += f'<w:vAlign w:val="{valign}"/>'
    pr += extra
    if not content.startswith("<w:p") and not content.startswith("<w:tbl"):
        content = p(content)
    if "<w:p>" not in content.split("</w:tc>")[-1] and not content.rstrip().endswith("</w:p>"):
        content += p("")  # a cell must end with a paragraph
    return f"<w:tc><w:tcPr>{pr}</w:tcPr>{content}</w:tc>"


def tr(cells, header=False, cant_split=False, height=None):
    pr = ""
    if cant_split:
        pr += "<w:cantSplit/>"
    if height:
        pr += f'<w:trHeight w:val="{height}"/>'
    if header:
        pr += "<w:tblHeader/>"
    return f'<w:tr>{"<w:trPr>" + pr + "</w:trPr>" if pr else ""}{"".join(cells)}</w:tr>'


def tbl(rows, grid, style="TableGrid", layout="fixed", borders="", width=None, look="04A0", extra=""):
    tw = width or sum(grid)
    tblw = '<w:tblW w:w="0" w:type="auto"/>' if layout == "autofit" else f'<w:tblW w:w="{tw}" w:type="dxa"/>'
    lay = '<w:tblLayout w:type="fixed"/>' if layout == "fixed" else ""
    bd = f"<w:tblBorders>{borders}</w:tblBorders>" if borders else ""
    pr = f'<w:tblPr><w:tblStyle w:val="{style}"/>{tblw}{bd}{lay}<w:tblLook w:val="{look}"/>{extra}</w:tblPr>'
    g = "".join(f'<w:gridCol w:w="{c}"/>' for c in grid)
    return f"<w:tbl>{pr}<w:tblGrid>{g}</w:tblGrid>{''.join(rows)}</w:tbl>"


def border(side, val="single", sz=4, color="auto", space=0):
    return f'<w:{side} w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'


# -------------------------------------------------------------- section props
def sect(w=12240, h=15840, orient="portrait", margins=(1440, 1440, 1440, 1440), cols=None, refs="", extra="", kind=None):
    t, rgt, b, lft = margins
    o = ' w:orient="landscape"' if orient == "landscape" else ""
    sz = f'<w:pgSz w:w="{w}" w:h="{h}"{o}/>'
    mar = f'<w:pgMar w:top="{t}" w:right="{rgt}" w:bottom="{b}" w:left="{lft}" w:header="720" w:footer="720" w:gutter="0"/>'
    c = ""
    if cols:
        n, space, sep = cols
        c = f'<w:cols w:num="{n}" w:space="{space}"{" w:sep=\"1\"" if sep else ""}/>'
    ty = f'<w:type w:val="{kind}"/>' if kind else ""
    return f"<w:sectPr>{refs}{ty}{sz}{mar}{c}{extra}</w:sectPr>"


def landscape_sect(**kw):
    return sect(w=15840, h=12240, orient="landscape", **kw)


# ----------------------------------------------------------------------- images
def png(width, height, shape="rect", fg=(30, 90, 200), bg=(240, 240, 240)):
    """Original geometric PNG (no third-party pixels): rect, circle, diag, stripes, checker."""
    rows = []
    cx, cy = width / 2, height / 2
    rad = min(width, height) * 0.4
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            on = False
            if shape == "rect":
                on = width * 0.2 < x < width * 0.8 and height * 0.2 < y < height * 0.8
            elif shape == "circle":
                on = (x - cx) ** 2 + (y - cy) ** 2 < rad**2
            elif shape == "diag":
                on = abs(x - y * width / height) < width * 0.12
            elif shape == "stripes":
                on = (x // max(1, width // 8)) % 2 == 0
            elif shape == "checker":
                on = ((x // max(1, width // 8)) + (y // max(1, height // 8))) % 2 == 0
            row += bytes(fg if on else bg)
        rows.append(bytes(row))
    raw = b"".join(rows)

    def chunk(tag, data):
        c = struct.pack(">I", len(data)) + tag + data
        return c + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b"")


EMU = 9525  # per px


def _pic(rid, pid, name, cx, cy, descr, crop=None):
    src = ""
    if crop:
        l, t, rr, b = crop
        src = f'<a:srcRect l="{l}" t="{t}" r="{rr}" b="{b}"/>'
    return (
        f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{pid}" name="{name}" descr="{esc(descr)}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}"/>{src}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        f"</pic:pic></a:graphicData></a:graphic>"
    )


def inline_image(rid, pid, wpx, hpx, descr="", crop=None):
    cx, cy = wpx * EMU, hpx * EMU
    return (
        f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:docPr id="{pid}" name="Picture {pid}" descr="{esc(descr)}"/><wp:cNvGraphicFramePr/>'
        f"{_pic(rid, pid, f'img{pid}', cx, cy, descr, crop)}</wp:inline></w:drawing></w:r>"
    )


def anchor_image(rid, pid, wpx, hpx, descr="", wrap="square", hpos=("margin", "right"), vpos=("paragraph", 0), crop=None, behind=False):
    cx, cy = wpx * EMU, hpx * EMU
    ph = f'<wp:positionH relativeFrom="{hpos[0]}"><wp:align>{hpos[1]}</wp:align></wp:positionH>'
    pv = f'<wp:positionV relativeFrom="{vpos[0]}"><wp:posOffset>{vpos[1]}</wp:posOffset></wp:positionV>'
    wraps = {
        "square": '<wp:wrapSquare wrapText="bothSides"/>',
        "tight": '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>',
        "topAndBottom": "<wp:wrapTopAndBottom/>",
        "none": "<wp:wrapNone/>",
    }[wrap]
    return (
        f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" '
        f'behindDoc="{1 if behind else 0}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>{ph}{pv}'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wraps}'
        f'<wp:docPr id="{pid}" name="Picture {pid}" descr="{esc(descr)}"/><wp:cNvGraphicFramePr/>'
        f"{_pic(rid, pid, f'img{pid}', cx, cy, descr, crop)}</wp:anchor></w:drawing></w:r>"
    )


# ---------------------------------------------------------------- shared parts
def _rfonts_default():
    return '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>'


def _style(kind, sid, name, based=None, nxt=None, link=None, ppr="", rp="", default=False, extra="", ui=None, q=True):
    d = ' w:default="1"' if default else ""
    s = f'<w:style w:type="{kind}"{d} w:styleId="{sid}"><w:name w:val="{name}"/>'
    if based:
        s += f'<w:basedOn w:val="{based}"/>'
    if nxt:
        s += f'<w:next w:val="{nxt}"/>'
    if link:
        s += f'<w:link w:val="{link}"/>'
    if ui is not None:
        s += f'<w:uiPriority w:val="{ui}"/>'
    if q:
        s += "<w:qFormat/>"
    s += extra
    if ppr:
        s += f"<w:pPr>{ppr}</w:pPr>"
    if rp:
        s += f"<w:rPr>{rp}</w:rPr>"
    return s + "</w:style>"


def styles_xml():
    S = []
    S.append(_style("paragraph", "Normal", "Normal", default=True, ppr='<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>'))
    S.append('<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>')
    S.append('<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>')
    S.append('<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/></w:style>')
    sizes = {1: (32, "2F5496"), 2: (26, "2F5496"), 3: (24, "1F3763")}
    for lvl, (sz, col) in sizes.items():
        S.append(
            _style("paragraph", f"Heading{lvl}", f"heading {lvl}", based="Normal", nxt="BodyText", link=f"Heading{lvl}Char", ui=9,
                   ppr=f'<w:keepNext/><w:keepLines/><w:spacing w:before="{240 if lvl == 1 else 160}" w:after="80"/><w:outlineLvl w:val="{lvl - 1}"/>',
                   rp=f'<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:b/><w:color w:val="{col}"/><w:sz w:val="{sz}"/>')
        )
        S.append(_style("character", f"Heading{lvl}Char", f"Heading {lvl} Char", based="DefaultParagraphFont", link=f"Heading{lvl}", ui=9, q=False,
                        rp=f'<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:b/><w:color w:val="{col}"/><w:sz w:val="{sz}"/>'))
    S.append(_style("paragraph", "Title", "Title", based="Normal", nxt="Subtitle", ppr='<w:spacing w:after="120"/><w:contextualSpacing/>', rp='<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:sz w:val="56"/><w:spacing w:val="-10"/>'))
    S.append(_style("paragraph", "Subtitle", "Subtitle", based="Normal", nxt="BodyText", rp='<w:i/><w:color w:val="595959"/><w:sz w:val="28"/>'))
    S.append(_style("paragraph", "BodyText", "Body Text", based="Normal", link="BodyTextChar", ppr='<w:spacing w:after="120" w:line="276" w:lineRule="auto"/><w:jc w:val="both"/>'))
    S.append(_style("character", "BodyTextChar", "Body Text Char", based="DefaultParagraphFont", link="BodyText", q=False))
    S.append(_style("paragraph", "Quote", "Quote", based="BodyText", nxt="BodyText", ppr='<w:ind w:left="720" w:right="720"/><w:jc w:val="center"/>', rp="<w:i/><w:color w:val=\"404040\"/>"))
    S.append(_style("paragraph", "IntenseQuote", "Intense Quote", based="Quote", nxt="BodyText", ppr='<w:pBdr><w:top w:val="single" w:sz="4" w:space="10" w:color="4472C4"/><w:bottom w:val="single" w:sz="4" w:space="10" w:color="4472C4"/></w:pBdr>', rp='<w:b/><w:color w:val="4472C4"/>'))
    S.append(_style("paragraph", "CodeBlock", "Code Block", based="BodyText", link="CodeBlockChar", ppr='<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/>', rp='<w:rFonts w:ascii="DejaVu Sans Mono" w:hAnsi="DejaVu Sans Mono" w:cs="DejaVu Sans Mono"/><w:sz w:val="18"/>'))
    S.append(_style("character", "CodeBlockChar", "Code Block Char", based="BodyTextChar", link="CodeBlock", q=False, rp='<w:rFonts w:ascii="DejaVu Sans Mono" w:hAnsi="DejaVu Sans Mono"/><w:sz w:val="18"/>'))
    S.append(_style("paragraph", "ListParagraph", "List Paragraph", based="Normal", ppr='<w:ind w:left="720"/><w:contextualSpacing/>'))
    S.append(_style("paragraph", "Caption", "caption", based="Normal", nxt="Normal", rp='<w:i/><w:color w:val="44546A"/><w:sz w:val="18"/>', ppr='<w:spacing w:after="200" w:line="240" w:lineRule="auto"/>'))
    for i in (1, 2, 3):
        S.append(_style("paragraph", f"TOC{i}", f"toc {i}", based="Normal", nxt="Normal", q=False, ppr=f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs><w:spacing w:after="100"/><w:ind w:left="{(i - 1) * 220}"/>'))
    S.append(_style("paragraph", "TOCHeading", "TOC Heading", based="Heading1", nxt="Normal", ppr='<w:outlineLvl w:val="9"/>'))
    S.append(_style("paragraph", "Header", "header", based="Normal", link="HeaderChar", q=False, ppr='<w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'))
    S.append(_style("character", "HeaderChar", "Header Char", based="DefaultParagraphFont", link="Header", q=False))
    S.append(_style("paragraph", "Footer", "footer", based="Normal", link="FooterChar", q=False, ppr='<w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'))
    S.append(_style("character", "FooterChar", "Footer Char", based="DefaultParagraphFont", link="Footer", q=False))
    S.append(_style("character", "Hyperlink", "Hyperlink", based="DefaultParagraphFont", q=False, ui=99, rp='<w:color w:val="0563C1"/><w:u w:val="single"/>'))
    S.append(_style("character", "Strong", "Strong", based="DefaultParagraphFont", rp="<w:b/>"))
    S.append(_style("character", "Emphasis", "Emphasis", based="DefaultParagraphFont", rp="<w:i/>"))
    S.append(_style("character", "Accent", "Accent Word", based="Strong", rp='<w:color w:val="C55A11"/><w:smallCaps/>'))
    S.append(_style("paragraph", "FootnoteText", "footnote text", based="Normal", link="FootnoteTextChar", q=False, ppr='<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', rp='<w:sz w:val="20"/>'))
    S.append(_style("character", "FootnoteTextChar", "Footnote Text Char", based="DefaultParagraphFont", link="FootnoteText", q=False, rp='<w:sz w:val="20"/>'))
    S.append(_style("character", "FootnoteReference", "footnote reference", based="DefaultParagraphFont", q=False, rp='<w:vertAlign w:val="superscript"/>'))
    S.append(_style("paragraph", "EndnoteText", "endnote text", based="Normal", link="EndnoteTextChar", q=False, ppr='<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', rp='<w:sz w:val="20"/>'))
    S.append(_style("character", "EndnoteTextChar", "Endnote Text Char", based="DefaultParagraphFont", link="EndnoteText", q=False, rp='<w:sz w:val="20"/>'))
    S.append(_style("character", "EndnoteReference", "endnote reference", based="DefaultParagraphFont", q=False, rp='<w:vertAlign w:val="superscript"/>'))
    S.append(_style("character", "CommentReference", "annotation reference", based="DefaultParagraphFont", q=False, rp='<w:sz w:val="16"/>'))
    S.append(_style("paragraph", "CommentText", "annotation text", based="Normal", link="CommentTextChar", q=False, ppr='<w:spacing w:line="240" w:lineRule="auto"/>', rp='<w:sz w:val="20"/>'))
    S.append(_style("character", "CommentTextChar", "Comment Text Char", based="DefaultParagraphFont", link="CommentText", q=False, rp='<w:sz w:val="20"/>'))
    S.append(_style("paragraph", "Placeholder", "Placeholder Para", based="Normal", rp='<w:color w:val="808080"/>', q=False))
    S.append(_style("table", "TableGrid", "Table Grid", based="TableNormal", ui=39, ppr='<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>',
                    extra='<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>'))
    S.append(
        '<w:style w:type="table" w:styleId="BandedBlue"><w:name w:val="Banded Blue"/><w:basedOn w:val="TableGrid"/><w:uiPriority w:val="40"/><w:qFormat/>'
        '<w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="4472C4"/><w:bottom w:val="single" w:sz="8" w:space="0" w:color="4472C4"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="B4C6E7"/></w:tblBorders></w:tblPr>'
        '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>'
        '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr></w:style>'
    )
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS_DECL}>'
        f'<w:docDefaults><w:rPrDefault><w:rPr>{_rfonts_default()}<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
        f'<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
        f'{"".join(S)}</w:styles>'
    )


def _lvl(i, fmt, text, left, hanging=360, font=None, start=1, suff=None):
    f = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/></w:rPr>' if font else ""
    s = f'<w:suff w:val="{suff}"/>' if suff else ""
    return (f'<w:lvl w:ilvl="{i}"><w:start w:val="{start}"/><w:numFmt w:val="{fmt}"/>{s}<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr>{f}</w:lvl>')


def numbering_xml(extra=""):
    """Abstract lists: 0 outline decimal (1. 1.1 1.1.1), 1 bullets, 2 legal-letter (a) (i), 3 headings-linked."""
    a0 = "".join(_lvl(i, "decimal", ".".join(f"%{j + 1}" for j in range(i + 1)) + ("." if i == 0 else ""), 720 + 360 * i, 360 + 120 * i) for i in range(9))
    bul = ["•", "o", "▪"]
    a1 = "".join(_lvl(i, "bullet", bul[i % 3], 720 + 360 * i, 360, font="Liberation Sans" if i % 3 == 1 else "DejaVu Sans") for i in range(9))
    fmts = ["lowerLetter", "lowerRoman", "decimal", "upperLetter", "upperRoman"]
    a2 = "".join(_lvl(i, fmts[i % 5], f"({chr(0x25)}{i + 1})" if False else f"(%{i + 1})", 720 + 360 * i, 360) for i in range(9))
    a3 = "".join(_lvl(i, "decimal", ".".join(f"%{j + 1}" for j in range(i + 1)), 432 + 144 * i, 432 + 144 * i, suff="space") for i in range(9))
    absn = "".join(
        f'<w:abstractNum w:abstractNumId="{n}"><w:nsid w:val="0000000{n}"/><w:multiLevelType w:val="{"multilevel" if n != 1 else "hybridMultilevel"}"/>{body}</w:abstractNum>'
        for n, body in enumerate([a0, a1, a2, a3])
    )
    nums = (
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
        '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>'
        # restart: same abstract list, numbering starts over at 1 / at 10
        '<w:num w:numId="4"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>'
        '<w:num w:numId="5"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="10"/></w:lvlOverride><w:lvlOverride w:ilvl="1"><w:startOverride w:val="3"/></w:lvlOverride></w:num>'
        '<w:num w:numId="6"><w:abstractNumId w:val="3"/></w:num>'
    )
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {NS_DECL}>{absn}{nums}{extra}</w:numbering>'


def settings_xml(extra="", footnote_pr=True):
    fn = ""
    if footnote_pr:
        fn = '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>'
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS_DECL}>{extra}'
            f'<w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>{fn}'
            f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')


def font_table_xml(fonts=("Liberation Sans", "Liberation Serif", "DejaVu Sans", "DejaVu Sans Mono"), extra=""):
    body = "".join(f'<w:font w:name="{f}"><w:charset w:val="00"/><w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>' for f in fonts)
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts {NS_DECL}>{body}{extra}</w:fonts>'


def theme_xml(major="Liberation Serif", minor="Liberation Sans", name="Spindle Test Theme", accents=("4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47")):
    acc = "".join(f'<a:accent{i + 1}><a:srgbClr val="{c}"/></a:accent{i + 1}>' for i, c in enumerate(accents))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<a:theme xmlns:a="{NS["a"]}" name="{name}"><a:themeElements>'
        f'<a:clrScheme name="{name} Colors"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        f'<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>{acc}'
        '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>'
        f'<a:fontScheme name="{name} Fonts"><a:majorFont><a:latin typeface="{major}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
        f'<a:minorFont><a:latin typeface="{minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
        '<a:fmtScheme name="Plain"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
        '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
        '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
        '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
        "</a:themeElements></a:theme>"
    )


# ---------------------------------------------------------------------- package
class Package:
    """In-memory OPC package with a main document and relationship bookkeeping."""

    def __init__(self, title="", creator="Spindle fixture generator", subject=""):
        self.parts = {}
        self.defaults = {"rels": "application/vnd.openxmlformats-package.relationships+xml", "xml": "application/xml"}
        self.overrides = {"/word/document.xml": CT_MAIN}
        self.doc_rels = []  # (rid, type, target, mode)
        self._rid = 0
        self.title, self.creator, self.subject = title, creator, subject

    def rel(self, rtype, target, external=False, rid=None):
        self._rid += 1
        rid = rid or f"rId{self._rid}"
        self.doc_rels.append((rid, rtype if rtype.startswith("http") else REL + rtype, target, external))
        return rid

    def part(self, name, data, ctype=None, rtype=None, rel=True):
        """Add a part (str/bytes). With rtype, also relate it from word/document.xml and return its rId."""
        self.parts[name] = data.encode("utf-8") if isinstance(data, str) else data
        if ctype:
            self.overrides["/" + name] = ctype
        if rtype and rel:
            return self.rel(rtype, name[len("word/"):] if name.startswith("word/") else "../" + name)
        return None

    def media(self, name, data, ext="png"):
        self.defaults.setdefault(ext, {"png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg", "gif": "image/gif", "ttf": "application/x-font-ttf", "odttf": "application/vnd.openxmlformats-officedocument.obfuscatedFont", "bin": "application/vnd.openxmlformats-officedocument.oleObject"}[ext])
        return self.part("word/media/" + name, data, rtype="image")

    def standard(self, styles=None, numbering=None, settings=None, theme=None, fonts=None):
        self.part("word/styles.xml", styles or styles_xml(), CT["styles"], "styles")
        self.part("word/numbering.xml", numbering or numbering_xml(), CT["numbering"], "numbering")
        self.part("word/settings.xml", settings or settings_xml(), CT["settings"], "settings")
        self.part("word/fontTable.xml", fonts or font_table_xml(), CT["fontTable"], "fontTable")
        self.part("word/theme/theme1.xml", theme or theme_xml(), CT["theme"], "theme")

    def document(self, body, final_sect=None, root_extra=""):
        sp = final_sect if final_sect is not None else sect()
        self.parts["word/document.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS_DECL}{root_extra}><w:body>{body}{sp}</w:body></w:document>'
        ).encode("utf-8")

    def notes(self, kind, items):
        """kind footnote|endnote; items: list of (id, inner paragraphs xml). Adds separators."""
        k = kind
        sep = (f'<w:{k} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:{k}>'
               f'<w:{k} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:{k}>')
        body = "".join(f'<w:{k} w:id="{i}">{x}</w:{k}>' for i, x in items)
        return self.part(f"word/{k}s.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:{k}s {NS_DECL}>{sep}{body}</w:{k}s>', CT[f"{k}s"], f"{k}s")

    def note_ref(self, kind, nid):
        st = "FootnoteReference" if kind == "footnote" else "EndnoteReference"
        return f'<w:r><w:rPr><w:rStyle w:val="{st}"/></w:rPr><w:{kind}Reference w:id="{nid}"/></w:r>'

    def note_para(self, kind, text, **kw):
        st, ref = ("FootnoteText", "footnoteRef") if kind == "footnote" else ("EndnoteText", "endnoteRef")
        rs = "FootnoteReference" if kind == "footnote" else "EndnoteReference"
        return p(f'<w:r><w:rPr><w:rStyle w:val="{rs}"/></w:rPr><w:{ref}/></w:r>' + r(" " + text, **kw), style=st)

    def header_footer(self, kind, inner, name):
        k = "hdr" if kind == "header" else "ftr"
        return self.part(f"word/{name}.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:{k} {NS_DECL}>{inner}</w:{k}>', CT[kind], kind)

    def save(self, path):
        ct = "".join(f'<Default Extension="{e}" ContentType="{c}"/>' for e, c in self.defaults.items())
        ct += "".join(f'<Override PartName="{n}" ContentType="{c}"/>' for n, c in sorted(self.overrides.items()))
        ct += '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        ct += '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        content_types = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">{ct}</Types>'
        root_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL}officeDocument" Target="word/document.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            f'<Relationship Id="rId3" Type="{REL}extended-properties" Target="docProps/app.xml"/></Relationships>'
        )
        rels = "".join(
            f'<Relationship Id="{rid}" Type="{t}" Target="{esc(tg)}"{" TargetMode=\"External\"" if ext else ""}/>' for rid, t, tg, ext in self.doc_rels
        )
        doc_rels = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>'
        core = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f"<dc:title>{esc(self.title)}</dc:title><dc:subject>{esc(self.subject)}</dc:subject><dc:creator>{esc(self.creator)}</dc:creator>"
            '<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-15T09:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-15T09:00:00Z</dcterms:modified></cp:coreProperties>'
        )
        app = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Spindle fixture generator</Application></Properties>'
        out = {"[Content_Types].xml": content_types.encode(), "_rels/.rels": root_rels.encode(), "word/_rels/document.xml.rels": doc_rels.encode(),
               "docProps/core.xml": core.encode(), "docProps/app.xml": app.encode()}
        out.update(self.parts)
        write_zip(path, out)


def write_zip(path, parts):
    """Deterministic zip: fixed timestamps, [Content_Types].xml first, sorted otherwise."""
    names = sorted(parts, key=lambda n: (n != "[Content_Types].xml", n))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for n in names:
            zi = zipfile.ZipInfo(n, date_time=(2026, 1, 15, 9, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o644 << 16
            z.writestr(zi, parts[n])


def read_zip(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}
