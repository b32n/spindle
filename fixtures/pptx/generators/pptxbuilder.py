"""Tiny PresentationML package builder used by the PPTX fixture generators.

Everything in the hand-authored corpus is written here or in the build_*.py scripts that import this module; there is
no third-party content. Output is deterministic: fixed zip timestamps, fixed core-properties dates, no author names,
no randomness. `Deck` assembles masters, layouts, slides, notes and every relationship and content-type entry.
"""
import re
import struct
import zipfile
import zlib
from xml.sax.saxutils import escape as _esc, unescape as _unesc

STAMP = (2026, 1, 15, 9, 0, 0)
ISO = "2026-01-15T09:00:00Z"
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = R + "/"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = f'xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"'
PML = "application/vnd.openxmlformats-officedocument.presentationml."
DML = "application/vnd.openxmlformats-officedocument.drawingml."
CT = {
    "presentation": PML + "presentation.main+xml", "macro": "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
    "master": PML + "slideMaster+xml", "layout": PML + "slideLayout+xml", "slide": PML + "slide+xml", "notes": PML + "notesSlide+xml",
    "notesMaster": PML + "notesMaster+xml", "handoutMaster": PML + "handoutMaster+xml", "presProps": PML + "presProps+xml",
    "viewProps": PML + "viewProps+xml", "tableStyles": PML + "tableStyles+xml", "commentAuthors": PML + "commentAuthors+xml",
    "comments": PML + "comments+xml", "theme": "application/vnd.openxmlformats-officedocument.theme+xml",
    "chart": DML + "chart+xml", "diagramData": DML + "diagramData+xml", "diagramLayout": DML + "diagramLayout+xml",
    "diagramStyle": DML + "diagramStyle+xml", "diagramColors": DML + "diagramColors+xml",
    "diagramDrawing": "application/vnd.ms-office.drawingml.diagramDrawing+xml",
    "core": "application/vnd.openxmlformats-package.core-properties+xml",
    "app": "application/vnd.openxmlformats-officedocument.extended-properties+xml",
}
DEFAULT_TYPES = {
    "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif", "svg": "image/svg+xml", "emf": "image/x-emf",
    "wmf": "image/x-wmf", "bin": "application/vnd.openxmlformats-officedocument.oleObject", "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "wav": "audio/wav", "mp4": "video/mp4", "ink": "application/inkml+xml", "vml": "application/vnd.openxmlformats-officedocument.vmlDrawing",
    "mp3": "audio/mpeg", "xml": "application/xml", "rels": "application/vnd.openxmlformats-package.relationships+xml",
}
SLIDE_W, SLIDE_H = 12192000, 6858000
COLORS = [("dk2", "44546A"), ("lt2", "E7E6E6"), ("accent1", "4472C4"), ("accent2", "ED7D31"), ("accent3", "A5A5A5"), ("accent4", "FFC000"),
          ("accent5", "5B9BD5"), ("accent6", "70AD47"), ("hlink", "0563C1"), ("folHlink", "954F72")]


def esc(s):
    return _esc(str(s), {'"': "&quot;"})


class Rels:
    def __init__(self):
        self.items = []

    def add(self, typ, target, external=False):
        rid = f"rId{len(self.items) + 1}"
        self.items.append((rid, typ if typ.startswith("http") else REL + typ, target, external))
        return rid

    def xml(self):
        body = "".join(f'<Relationship Id="{i}" Type="{t}" Target="{esc(g)}"' + (' TargetMode="External"' if e else "") + "/>" for i, t, g, e in self.items)
        return f'{HEAD}<Relationships xmlns="{PKG_REL}">{body}</Relationships>'


def core_xml(title="Fixture"):
    return (f'{HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" '
            f'xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>{esc(title)}</dc:title>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{ISO}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{ISO}</dcterms:modified></cp:coreProperties>')


def app_xml(application="Spindle fixture generator"):
    return f'{HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>{esc(application)}</Application></Properties>'


def content_types(parts, overrides):
    """Defaults for the extensions in `parts` that have one, an Override for every named part."""
    exts = {n.rsplit(".", 1)[-1].lower() for n in parts if "." in n.rsplit("/", 1)[-1]} | {"rels", "xml"}
    d = "".join(f'<Default Extension="{e}" ContentType="{DEFAULT_TYPES[e]}"/>' for e in sorted(exts) if e in DEFAULT_TYPES)
    o = "".join(f'<Override PartName="/{n}" ContentType="{c}"/>' for n, c in sorted(overrides.items()))
    return f'{HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">{d}{o}</Types>'


def write_zip(path, parts):
    """Deterministic zip: fixed timestamps, [Content_Types].xml first, sorted otherwise."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for n in sorted(parts, key=lambda n: (n != "[Content_Types].xml", n)):
            zi = zipfile.ZipInfo(n, date_time=STAMP)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o644 << 16
            z.writestr(zi, parts[n])


def read_zip(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def normalise(path):
    """Producer-variety files: replace docProps with the fixed ones (keeping the producer's Application name) and re-zip deterministically."""
    parts = read_zip(path)
    if "docProps/core.xml" in parts:
        parts["docProps/core.xml"] = core_xml("Fixture").encode()
    if "docProps/app.xml" in parts:
        m = re.search(rb"<Application>([^<]*)</Application>", parts["docProps/app.xml"])
        parts["docProps/app.xml"] = app_xml(_unesc(m.group(1).decode()) if m else "Spindle fixture generator").encode()
    write_zip(path, parts)


def chunk(tag, data):
    return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)


def png(w=16, h=16, fg=(30, 90, 200), bg=(240, 240, 240)):
    """Original two-colour PNG (diagonal), no third-party pixels."""
    raw = b"".join(b"\x00" + b"".join(bytes(fg if abs(x - y) < 2 else bg) for x in range(w)) for y in range(h))
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b"")


def theme_xml(name="Spindle Test Theme", colors=None, major="Calibri Light", minor="Calibri", ea=("", ""), effects="plain", font_scheme="Spindle", clr_scheme="Spindle"):
    """A complete theme part. `effects` = 'plain' (three empty effect styles) or 'shadow' (outer shadow, glow and soft edge in styles 2 and 3)."""
    clr = "".join(f'<a:{n}><a:srgbClr val="{v}"/></a:{n}>' for n, v in (colors or COLORS))
    ln = "".join(f'<a:ln w="{w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' for w in (6350, 12700, 19050))
    fills = ('<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
             '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
             '<a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill>')
    plain = "<a:effectStyle><a:effectLst/></a:effectStyle>"
    rich = ('<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>'
            '<a:effectStyle><a:effectLst><a:glow rad="63500"><a:schemeClr val="accent1"><a:alpha val="40000"/></a:schemeClr></a:glow><a:softEdge rad="12700"/></a:effectLst></a:effectStyle>')
    est = plain + rich if effects == "shadow" else plain * 3
    return (f'{HEAD}<a:theme xmlns:a="{A}" name="{esc(name)}"><a:themeElements>'
            f'<a:clrScheme name="{esc(clr_scheme)}"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>{clr}</a:clrScheme>'
            f'<a:fontScheme name="{esc(font_scheme)}"><a:majorFont><a:latin typeface="{esc(major)}"/><a:ea typeface="{esc(ea[0])}"/><a:cs typeface=""/></a:majorFont>'
            f'<a:minorFont><a:latin typeface="{esc(minor)}"/><a:ea typeface="{esc(ea[1])}"/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
            f'<a:fmtScheme name="Spindle"><a:fillStyleLst>{fills}</a:fillStyleLst><a:lnStyleLst>{ln}</a:lnStyleLst>'
            f'<a:effectStyleLst>{est}</a:effectStyleLst><a:bgFillStyleLst>{fills}</a:bgFillStyleLst></a:fmtScheme>'
            '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>')


# --- placeholder geometry: fractions of the slide so every slide size gets a sensible layout -----------------------------

def _geo(cx, cy):
    m = int(cx * 0.07)
    return {
        "title": (m, int(cy * 0.05), cx - 2 * m, int(cy * 0.18)),
        "body": (m, int(cy * 0.26), cx - 2 * m, int(cy * 0.6)),
        "ctrTitle": (m, int(cy * 0.2), cx - 2 * m, int(cy * 0.25)),
        "subTitle": (m, int(cy * 0.48), cx - 2 * m, int(cy * 0.2)),
        "dt": (m, int(cy * 0.92), int(cx * 0.25), int(cy * 0.06)),
        "ftr": (int(cx * 0.34), int(cy * 0.92), int(cx * 0.32), int(cy * 0.06)),
        "sldNum": (int(cx * 0.72), int(cy * 0.92), int(cx * 0.21), int(cy * 0.06)),
        "half1": (m, int(cy * 0.26), int(cx * 0.42), int(cy * 0.6)),
        "half2": (int(cx * 0.51), int(cy * 0.26), int(cx * 0.42), int(cy * 0.6)),
    }


def xfrm(x, y, cx, cy, rot=0, flipH=False, flipV=False, tag="a:xfrm"):
    attrs = (f' rot="{rot}"' if rot else "") + (' flipH="1"' if flipH else "") + (' flipV="1"' if flipV else "")
    return f'<{tag}{attrs}><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></{tag}>'


def ph_shape(sid, name, ph, geo=None, text="", body_pr="<a:bodyPr/>", lst="", prst=None, layout_text=True):
    """A placeholder shape. `ph` = (type or None, idx or None, extra attributes); `geo` None inherits the position."""
    typ, idx, extra = ph
    attrs = (f' type="{typ}"' if typ else "") + (f' idx="{idx}"' if idx is not None else "") + extra
    sppr = (xfrm(*geo) if geo else "") + (f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>' if prst else "")
    tx = f"<p:txBody>{body_pr}{lst or '<a:lstStyle/>'}<a:p>{text}</a:p></p:txBody>" if text is not None else ""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{esc(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph{attrs}/></p:nvPr></p:nvSpPr>'
            f"<p:spPr>{sppr}</p:spPr>{tx}</p:sp>")


FLD = {"dt": '<a:fld id="{B1000000-0000-4000-8000-000000000001}" type="datetimeFigureOut"><a:rPr lang="en-US"/><a:t>1/15/2026</a:t></a:fld>',
       "sldNum": '<a:fld id="{B1000000-0000-4000-8000-000000000002}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld>',
       "ftr": ""}


def _run(t):
    return f'<a:r><a:rPr lang="en-US"/><a:t>{esc(t)}</a:t></a:r>'


GRP = ('<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
       '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>')


def _level_styles(tag, size, color="tx1", font="+mn-lt", bullets=True):
    out = f"<p:{tag}>"
    for lvl in range(1, 10):
        marl = 0 if not bullets else 228600 + (lvl - 1) * 457200
        ind = 0 if not bullets else -228600
        bu = '<a:buFont typeface="Arial"/><a:buChar char="&#8226;"/>' if bullets else "<a:buNone/>"
        out += (f'<a:lvl{lvl}pPr marL="{marl}" indent="{ind}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
                f'<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef>{bu}'
                f'<a:defRPr sz="{max(1000, size - (lvl - 1) * 200)}" kern="1200"><a:solidFill><a:schemeClr val="{color}"/></a:solidFill>'
                f'<a:latin typeface="{font}"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl{lvl}pPr>')
    return out + f"</p:{tag}>"


class Layout:
    def __init__(self, name, kind, shapes="", bg="", show_master_sp=True):
        self.name, self.kind, self.shapes, self.bg, self.show_master_sp = name, kind, shapes, bg, show_master_sp


class Master:
    def __init__(self, name="Office Theme", theme=None, clr_map=None, bg=None, shapes="", layouts=None, hf=True):
        self.name, self.theme, self.shapes, self.hf = name, theme, shapes, hf
        self.clr_map = clr_map or {}
        self.bg = bg
        self.layouts = layouts


CLR_MAP = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2", "accent1": "accent1", "accent2": "accent2", "accent3": "accent3", "accent4": "accent4",
           "accent5": "accent5", "accent6": "accent6", "hlink": "hlink", "folHlink": "folHlink"}


def standard_layouts(cx, cy, extra=()):
    g = _geo(cx, cy)
    foot = lambda s0: (ph_shape(s0, "Date Placeholder", ("dt", 10, ' sz="half"'), None, FLD["dt"]) + ph_shape(s0 + 1, "Footer Placeholder", ("ftr", 11, ' sz="quarter"'), None, "")
                       + ph_shape(s0 + 2, "Slide Number Placeholder", ("sldNum", 12, ' sz="quarter"'), None, FLD["sldNum"]))
    L = [
        Layout("Title Slide", "title", ph_shape(2, "Title 1", ("ctrTitle", None, ""), g["ctrTitle"], _run("Click to edit the title"))
               + ph_shape(3, "Subtitle 2", ("subTitle", 1, ""), g["subTitle"], _run("Click to edit the subtitle")) + foot(4)),
        Layout("Title and Content", "obj", ph_shape(2, "Title 1", ("title", None, ""), None, _run("Click to edit the title"))
               + ph_shape(3, "Content Placeholder 2", (None, 1, ""), None, _run("Click to edit the text")) + foot(4)),
        Layout("Two Content", "twoObj", ph_shape(2, "Title 1", ("title", None, ""), None, _run("Click to edit the title"))
               + ph_shape(3, "Content Placeholder 2", (None, 1, ' sz="half"'), g["half1"], _run("Left"))
               + ph_shape(4, "Content Placeholder 3", (None, 2, ' sz="half"'), g["half2"], _run("Right")) + foot(5)),
        Layout("Title Only", "titleOnly", ph_shape(2, "Title 1", ("title", None, ""), None, _run("Click to edit the title")) + foot(3)),
        Layout("Blank", "blank", foot(2)),
    ] + list(extra)
    return L


class Slide:
    def __init__(self, deck, master=0, layout=1):
        self.deck, self.master, self.layout = deck, master, layout
        self.shapes, self.rels, self.parts, self.rel_ids = [], Rels(), {}, {}
        self.bg = ""
        self.hidden = False
        self.transition = ""
        self.timing = ""
        self.notes = None
        self.clr_map_ovr = '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        self.show_master_sp = True
        self.next_id = 2
        self.ext = ""

    def sid(self):
        self.next_id += 1
        return self.next_id - 1

    def add(self, xml):
        self.shapes.append(xml)
        return self

    def media(self, name, data, kind="image"):
        """Add (or reuse) ppt/media/<name> and return the rId from this slide; one copy per deck."""
        self.deck.media[name] = data
        key = ("media", name)
        if key not in self.rel_ids:
            self.rel_ids[key] = self.rels.add(kind, f"../media/{name}")
        return self.rel_ids[key]

    def link(self, target, typ="hyperlink"):
        key = ("link", target)
        if key not in self.rel_ids:
            self.rel_ids[key] = self.rels.add(typ, target, external=not target.startswith("#") and not target.startswith("slide"))
        return self.rel_ids[key]

    def part(self, name, data, ct=None, typ=None, rel_target=None):
        """Attach a part (path inside the package) and relate the slide to it; returns the rId."""
        self.deck.extra[name] = (data.encode() if isinstance(data, str) else data, CT.get(ct, ct))
        return self.rels.add(typ, rel_target or "../" + name[len("ppt/"):])

    def xml(self):
        bg = f"<p:bg>{self.bg}</p:bg>" if self.bg else ""
        attrs = (' show="0"' if self.hidden else "") + ("" if self.show_master_sp else ' showMasterSp="0"')
        return (f'{HEAD}<p:sld {NS}{attrs}><p:cSld>{bg}<p:spTree>{GRP}{"".join(self.shapes)}</p:spTree></p:cSld>{self.clr_map_ovr}'
                f"{self.transition}{self.timing}{self.ext}</p:sld>")


class Deck:
    def __init__(self, title="Fixture", size=(SLIDE_W, SLIDE_H), size_type=None, masters=None, macro=False, notes_master=False, handout_master=False, extra_layouts=()):
        self.title, self.size, self.size_type, self.macro = title, size, size_type, macro
        self.masters = masters or [Master()]
        for m in self.masters:
            m.layouts = m.layouts or standard_layouts(*size, extra=extra_layouts)
        self.slides = []
        self.media, self.extra, self.parts_override = {}, {}, {}
        self.notes_master, self.handout_master = notes_master, handout_master
        self.pres_extra = {"sections": "", "custShows": "", "defaultTextStyle": True, "ext": ""}
        self.root_rels = Rels()
        self.pres_rels_extra = []  # (type, target) added after the standard ones
        self.comment_authors = None  # list of (id, name, initials) for legacy comments
        self.slide_comments = {}  # slide index -> list of (author, x, y, text)

    def slide(self, layout=1, master=0):
        s = Slide(self, master, layout)
        self.slides.append(s)
        return s

    # -- part writers ---------------------------------------------------------------------------------------------------

    def _master_xml(self, mi, m, layout_ids):
        cx, cy = self.size
        g = _geo(cx, cy)
        title = ph_shape(2, "Title Placeholder 1", ("title", None, ""), g["title"], _run("Click to edit Master title style"), '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr>', prst="rect")
        body = ph_shape(3, "Text Placeholder 2", ("body", 1, ""), g["body"], _run("Click to edit Master text styles"), '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr>', prst="rect")
        foot = ""
        if m.hf:
            foot = (ph_shape(4, "Date Placeholder 3", ("dt", 2, ' sz="half"'), g["dt"], FLD["dt"], prst="rect") + ph_shape(5, "Footer Placeholder 4", ("ftr", 3, ' sz="quarter"'), g["ftr"], "", prst="rect")
                    + ph_shape(6, "Slide Number Placeholder 5", ("sldNum", 4, ' sz="quarter"'), g["sldNum"], FLD["sldNum"], prst="rect"))
        bg = f"<p:bg>{m.bg}</p:bg>" if m.bg else '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
        clr = dict(CLR_MAP, **m.clr_map)
        cm = "<p:clrMap " + " ".join(f'{k}="{v}"' for k, v in clr.items()) + "/>"
        ids = "".join(f'<p:sldLayoutId id="{i}" r:id="rId{n}"/>' for n, i in enumerate(layout_ids, 1))
        tx = ("<p:txStyles>" + _level_styles("titleStyle", 4400, bullets=False).replace("+mn-lt", "+mj-lt").replace("<a:spcBef><a:spcPts val=\"600\"/></a:spcBef>", "<a:spcBef><a:spcPct val=\"0\"/></a:spcBef>")
              + _level_styles("bodyStyle", 2800) + _level_styles("otherStyle", 1800, bullets=False) + "</p:txStyles>")
        return (f'{HEAD}<p:sldMaster {NS}><p:cSld name="{esc(m.name)}">{bg}<p:spTree>{GRP}{title}{body}{foot}{m.shapes}</p:spTree></p:cSld>{cm}'
                f"<p:sldLayoutIdLst>{ids}</p:sldLayoutIdLst>{tx}</p:sldMaster>")

    def _layout_xml(self, lay):
        bg = f"<p:bg>{lay.bg}</p:bg>" if lay.bg else ""
        sm = "" if lay.show_master_sp else ' showMasterSp="0"'
        return (f'{HEAD}<p:sldLayout {NS} type="{lay.kind}" preserve="1"{sm}><p:cSld name="{esc(lay.name)}">{bg}<p:spTree>{GRP}{lay.shapes}</p:spTree></p:cSld>'
                '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>')

    def _aux_master(self, kind, root):
        if kind == "notesMaster":
            sl = ('<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr>'
                  '<p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>'
                  + ph_shape(3, "Notes Placeholder 2", ("body", 3, ' sz="quarter"'), (685800, 4400550, 5486400, 3600450), _run("Click to edit Master text styles"), prst="rect")
                  + ph_shape(4, "Header Placeholder 3", ("hdr", None, ' sz="quarter"'), (0, 0, 2971800, 458788), "", prst="rect")
                  + ph_shape(5, "Slide Number Placeholder 4", ("sldNum", 5, ' sz="quarter"'), (3884613, 8685213, 2971800, 458787), FLD["sldNum"], prst="rect"))
            body_end = _level_styles("notesStyle", 1200, bullets=False)
        else:
            sl = (ph_shape(2, "Header Placeholder 1", ("hdr", None, ' sz="quarter"'), (0, 0, 2971800, 458788), "", prst="rect")
                  + ph_shape(3, "Date Placeholder 2", ("dt", 1, ' sz="quarter"'), (3884613, 0, 2971800, 458788), FLD["dt"], prst="rect")
                  + ph_shape(4, "Footer Placeholder 3", ("ftr", 2, ' sz="quarter"'), (0, 8685213, 2971800, 458787), "", prst="rect")
                  + ph_shape(5, "Slide Number Placeholder 4", ("sldNum", 3, ' sz="quarter"'), (3884613, 8685213, 2971800, 458787), FLD["sldNum"], prst="rect"))
            body_end = ""
        clr = "<p:clrMap " + " ".join(f'{k}="{v}"' for k, v in CLR_MAP.items()) + "/>"
        bg = '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
        return f'{HEAD}<p:{root} {NS}><p:cSld>{bg}<p:spTree>{GRP}{sl}</p:spTree></p:cSld>{clr}{body_end}</p:{root}>'

    def build(self):
        P_ = {}
        ct = {}
        cx, cy = self.size
        # themes: one per master, then the notes and handout masters
        pres_rels = Rels()
        master_ids, layout_no, theme_no = [], 0, 0
        next_id = 2147483648
        sld_master_lst = ""
        layout_paths = {}  # (master index, layout index) -> part name
        for mi, m in enumerate(self.masters, 1):
            theme_no += 1
            theme_name = f"ppt/theme/theme{theme_no}.xml"
            P_[theme_name] = (m.theme or theme_xml(name=m.name)).encode()
            ct[theme_name] = CT["theme"]
            mrels = Rels()
            next_id += 1
            master_rid = pres_rels.add("slideMaster", f"slideMasters/slideMaster{mi}.xml")
            sld_master_lst += f'<p:sldMasterId id="{next_id}" r:id="{master_rid}"/>'
            lids = []
            for li, lay in enumerate(m.layouts):
                layout_no += 1
                next_id += 1
                lids.append(next_id)
                lname = f"ppt/slideLayouts/slideLayout{layout_no}.xml"
                layout_paths[(mi - 1, li)] = lname
                mrels.add("slideLayout", f"../slideLayouts/slideLayout{layout_no}.xml")
                P_[lname] = self._layout_xml(lay).encode()
                ct[lname] = CT["layout"]
                lr = Rels()
                lr.add("slideMaster", f"../slideMasters/slideMaster{mi}.xml")
                P_[f"ppt/slideLayouts/_rels/slideLayout{layout_no}.xml.rels"] = lr.xml().encode()
            mrels.add("theme", f"../theme/theme{theme_no}.xml")
            P_[f"ppt/slideMasters/slideMaster{mi}.xml"] = self._master_xml(mi, m, lids).encode()
            ct[f"ppt/slideMasters/slideMaster{mi}.xml"] = CT["master"]
            P_[f"ppt/slideMasters/_rels/slideMaster{mi}.xml.rels"] = mrels.xml().encode()
        notes_rid = handout_rid = None
        for kind, wanted, ctk in (("notesMaster", self.notes_master, "notesMaster"), ("handoutMaster", self.handout_master, "handoutMaster")):
            if not wanted:
                continue
            theme_no += 1
            P_[f"ppt/theme/theme{theme_no}.xml"] = theme_xml(name="Office Theme").encode()
            ct[f"ppt/theme/theme{theme_no}.xml"] = CT["theme"]
            P_[f"ppt/{kind}s/{kind}1.xml"] = self._aux_master(kind, kind).encode()
            ct[f"ppt/{kind}s/{kind}1.xml"] = CT[ctk]
            r = Rels()
            r.add("theme", f"../theme/theme{theme_no}.xml")
            P_[f"ppt/{kind}s/_rels/{kind}1.xml.rels"] = r.xml().encode()
            rid = pres_rels.add(kind, f"{kind}s/{kind}1.xml")
            if kind == "notesMaster":
                notes_rid = rid
            else:
                handout_rid = rid
        sld_ids, notes_no = "", 0
        for si, s in enumerate(self.slides, 1):
            srels = Rels()
            srels.items = list(s.rels.items)  # media, links and parts added first keep their rIds
            srels.add("slideLayout", "../slideLayouts/" + layout_paths[(s.master, s.layout)].rsplit("/", 1)[-1])
            P_[f"ppt/slides/slide{si}.xml"] = s.xml().encode()
            ct[f"ppt/slides/slide{si}.xml"] = CT["slide"]
            if s.notes is not None:
                notes_no += 1
                srels.add("notesSlide", f"../notesSlides/notesSlide{notes_no}.xml")
                nr = Rels()
                nr.add("notesMaster", "../notesMasters/notesMaster1.xml")
                nr.add("slide", f"../slides/slide{si}.xml")
                P_[f"ppt/notesSlides/_rels/notesSlide{notes_no}.xml.rels"] = nr.xml().encode()
                P_[f"ppt/notesSlides/notesSlide{notes_no}.xml"] = (
                    f'{HEAD}<p:notes {NS}><p:cSld><p:spTree>{GRP}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>'
                    f'<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>'
                    f"<p:txBody><a:bodyPr/><a:lstStyle/>{s.notes}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>").encode()
                ct[f"ppt/notesSlides/notesSlide{notes_no}.xml"] = CT["notes"]
            if si - 1 in self.slide_comments:
                srels.add("comments", f"../comments/comment{si}.xml")
                body = "".join(f'<p:cm authorId="{a}" dt="{ISO[:-1]}.000" idx="{n}"><p:pos x="{x}" y="{y}"/><p:text>{esc(t)}</p:text></p:cm>' for n, (a, x, y, t) in enumerate(self.slide_comments[si - 1], 1))
                P_[f"ppt/comments/comment{si}.xml"] = f'{HEAD}<p:cmLst {NS}>{body}</p:cmLst>'.encode()
                ct[f"ppt/comments/comment{si}.xml"] = CT["comments"]
            P_[f"ppt/slides/_rels/slide{si}.xml.rels"] = srels.xml().encode()
            sid = pres_rels.add("slide", f"slides/slide{si}.xml")
            sld_ids += f'<p:sldId id="{255 + si}" r:id="{sid}"/>'
        for name, data in self.media.items():
            P_[f"ppt/media/{name}"] = data
        for name, (data, c) in self.extra.items():
            P_[name] = data
            if c:
                ct[name] = c
        if self.comment_authors:
            pres_rels.add("commentAuthors", "commentAuthors.xml")
            au = "".join(f'<p:cmAuthor id="{i}" name="{esc(n)}" initials="{esc(ini)}" lastIdx="1" clrIdx="{i}"/>' for i, n, ini in self.comment_authors)
            P_["ppt/commentAuthors.xml"] = f'{HEAD}<p:cmAuthorLst {NS}>{au}</p:cmAuthorLst>'.encode()
            ct["ppt/commentAuthors.xml"] = CT["commentAuthors"]
        for typ, target in self.pres_rels_extra:
            pres_rels.add(typ, target)
        for name, typ, data, c in (("presProps", "presProps", f'{HEAD}<p:presentationPr {NS}/>', "presProps"),
                                   ("viewProps", "viewProps", f'{HEAD}<p:viewPr {NS}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>', "viewProps"),
                                   ("tableStyles", "tableStyles", f'{HEAD}<a:tblStyleLst xmlns:a="{A}" def="{{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}}"/>', "tableStyles")):
            pres_rels.add(typ, f"{name}.xml")
            P_[f"ppt/{name}.xml"] = data.encode()
            ct[f"ppt/{name}.xml"] = CT[c]
        pres_rels.add("theme", "theme/theme1.xml")
        pe = self.pres_extra
        nm = f'<p:notesMasterIdLst><p:notesMasterId r:id="{notes_rid}"/></p:notesMasterIdLst>' if notes_rid else ""
        hm = f'<p:handoutMasterIdLst><p:handoutMasterId r:id="{handout_rid}"/></p:handoutMasterIdLst>' if handout_rid else ""
        typ = f' type="{self.size_type}"' if self.size_type else ""
        dts = ("<p:defaultTextStyle><a:defPPr><a:defRPr lang=\"en-US\"/></a:defPPr>" + "".join(
            f'<a:lvl{n}pPr marL="{(n - 1) * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl{n}pPr>'
            for n in range(1, 10)) + "</p:defaultTextStyle>") if pe["defaultTextStyle"] else ""
        slst = f"<p:sldIdLst>{sld_ids}</p:sldIdLst>" if self.slides else ""
        P_["ppt/presentation.xml"] = (f'{HEAD}<p:presentation {NS} saveSubsetFonts="1"><p:sldMasterIdLst>{sld_master_lst}</p:sldMasterIdLst>{nm}{hm}{slst}'
                                      f'<p:sldSz cx="{cx}" cy="{cy}"{typ}/><p:notesSz cx="6858000" cy="9144000"/>{pe["custShows"]}{dts}{pe["ext"]}</p:presentation>').encode()
        ct["ppt/presentation.xml"] = CT["macro" if self.macro else "presentation"]
        P_["ppt/_rels/presentation.xml.rels"] = pres_rels.xml().encode()
        self.root_rels.add("officeDocument", "ppt/presentation.xml")
        self.root_rels.add("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml")
        self.root_rels.add(REL + "extended-properties", "docProps/app.xml")
        P_["_rels/.rels"] = self.root_rels.xml().encode()
        P_["docProps/core.xml"], ct["docProps/core.xml"] = core_xml(self.title).encode(), CT["core"]
        P_["docProps/app.xml"], ct["docProps/app.xml"] = app_xml().encode(), CT["app"]
        P_.update(self.parts_override)
        P_["[Content_Types].xml"] = content_types(P_, ct).encode()
        return P_

    def write(self, path):
        write_zip(path, self.build())
