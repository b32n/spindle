"""Tiny SpreadsheetML package builder used by the XLSX fixture generators.

Everything in the hand-authored corpus is written here or in the build_*.py scripts that import
this module; there is no third-party content. Output is deterministic: fixed zip timestamps,
fixed core-properties dates, no author names, no randomness.
"""
import io
import re
import struct
import zipfile
import zlib
from xml.sax.saxutils import escape as _esc, unescape as _unesc

STAMP = (2026, 1, 15, 9, 0, 0)
ISO = "2026-01-15T09:00:00Z"
M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = R + "/"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
SS = "application/vnd.openxmlformats-officedocument.spreadsheetml."
CT = {
    "workbook": SS + "sheet.main+xml", "worksheet": SS + "worksheet+xml", "styles": SS + "styles+xml",
    "sst": SS + "sharedStrings+xml", "theme": "application/vnd.openxmlformats-officedocument.theme+xml",
    "comments": SS + "comments+xml", "table": SS + "table+xml", "drawing": "application/vnd.openxmlformats-officedocument.drawing+xml",
    "chart": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
    "core": "application/vnd.openxmlformats-package.core-properties+xml",
    "app": "application/vnd.openxmlformats-officedocument.extended-properties+xml",
}
ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges",
         "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting",
         "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
         "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture",
         "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"]
HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def esc(s):
    return _esc(str(s), {'"': "&quot;"})


def col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def ref(row, c):
    return f"{col(c)}{row}"


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


class Styles:
    """Deduplicating styles.xml builder; index 0 of every pool is the default."""

    def __init__(self):
        self.fonts = ['<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>']
        self.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>']
        self.borders = ["<border><left/><right/><top/><bottom/><diagonal/></border>"]
        self.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
        self.fmts = {}
        self.dxfs = []
        self.cell_style_xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>']
        self.cell_styles = ['<cellStyle name="Normal" xfId="0" builtinId="0"/>']
        self.colors = ""

    @staticmethod
    def _pool(pool, xml):
        if xml not in pool:
            pool.append(xml)
        return pool.index(xml)

    def font(self, xml):
        return self._pool(self.fonts, xml)

    def fill(self, xml):
        return self._pool(self.fills, xml)

    def border(self, xml):
        return self._pool(self.borders, xml)

    @staticmethod
    def solid(rgb):
        return f'<fill><patternFill patternType="solid"><fgColor rgb="FF{rgb}"/><bgColor indexed="64"/></patternFill></fill>'

    def fmt(self, code):
        """Number-format id for a code: built-in ids for the common ones, else a custom id from 164."""
        builtin = {"General": 0, "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10, "0.00E+00": 11, "# ?/?": 12, "# ??/??": 13,
                   "mm-dd-yy": 14, "d-mmm-yy": 15, "d-mmm": 16, "mmm-yy": 17, "h:mm AM/PM": 18, "h:mm:ss AM/PM": 19, "h:mm": 20, "h:mm:ss": 21,
                   "m/d/yy h:mm": 22, "@": 49}
        if code in builtin:
            return builtin[code]
        return self.fmts.setdefault(code, 164 + len(self.fmts))

    def xf(self, fmt=None, font=None, fill=None, border=None, align="", prot="", xf_id=0, extra=""):
        n = self.fmt(fmt) if fmt else 0
        f = self.font(font) if font else 0
        fl = self.fill(fill) if fill else 0
        b = self.border(border) if border else 0
        attrs = f'numFmtId="{n}" fontId="{f}" fillId="{fl}" borderId="{b}" xfId="{xf_id}"'
        attrs += "".join(f' apply{k}="1"' for k, v in (("NumberFormat", n), ("Font", f), ("Fill", fl), ("Border", b)) if v)
        inner = (f"<alignment {align}/>" if align else "") + prot
        attrs += (' applyAlignment="1"' if align else "") + extra
        return self._pool(self.xfs, f"<xf {attrs}>{inner}</xf>" if inner else f"<xf {attrs}/>")

    def dxf(self, xml):
        self.dxfs.append(f"<dxf>{xml}</dxf>")
        return len(self.dxfs) - 1

    def xml(self):
        nf = "".join(f'<numFmt numFmtId="{i}" formatCode="{esc(c)}"/>' for c, i in self.fmts.items())
        out = f'{HEAD}<styleSheet xmlns="{M}">'
        out += f'<numFmts count="{len(self.fmts)}">{nf}</numFmts>' if nf else ""
        for tag, pool in (("fonts", self.fonts), ("fills", self.fills), ("borders", self.borders), ("cellStyleXfs", self.cell_style_xfs), ("cellXfs", self.xfs), ("cellStyles", self.cell_styles)):
            out += f'<{tag} count="{len(pool)}">{"".join(pool)}</{tag}>'
        out += f'<dxfs count="{len(self.dxfs)}">{"".join(self.dxfs)}</dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
        return out + self.colors + "</styleSheet>"


class Sheet:
    def __init__(self, name, state=None, book=None):
        self.name, self.state, self.book = name, state, book
        self.cells, self.rowattrs, self.x, self.rels, self.parts = {}, {}, {}, Rels(), {}
        self.ns = ""

    def set(self, r, c, v, s=0, **kw):
        """v: int/float (number), bool, str (shared string), ('e', '#N/A'), ('i', text), ('f', formula, cached, attrs), ('raw', xml)."""
        self.cells[(r, c)] = (v, s, kw)
        return self

    def row(self, r, values, c=1, s=0):
        for i, v in enumerate(values):
            self.set(r, c + i, v, s)
        return self

    def _cell(self, r, c, v, s, kw):
        a = f'r="{ref(r, c)}"' + (f' s="{s}"' if s else "") + "".join(f' {k}="{x}"' for k, x in kw.items())
        if isinstance(v, bool):
            return f'<c {a} t="b"><v>{int(v)}</v></c>'
        if isinstance(v, (int, float)):
            return f"<c {a}><v>{v}</v></c>"
        if isinstance(v, str):
            return f'<c {a} t="s"><v>{self.book.sst(v)}</v></c>'
        if v[0] == "e":
            return f'<c {a} t="e"><v>{v[1]}</v></c>'
        if v[0] == "i":
            return f'<c {a} t="inlineStr"><is><t>{esc(v[1])}</t></is></c>'
        if v[0] == "raw":
            return f"<c {a}>{v[1]}</c>"
        f, cached, attrs = v[1], (v[2] if len(v) > 2 else None), (v[3] if len(v) > 3 else "")
        t = ""
        if isinstance(cached, bool):
            t, cv = ' t="b"', int(cached)
        elif isinstance(cached, str):
            t, cv = ' t="str"', esc(cached)
        else:
            cv = cached
        fx = f"<f{attrs}>{esc(f)}</f>" if f else f"<f{attrs}/>"
        return f"<c {a}{t}>{fx}" + (f"<v>{cv}</v>" if cached is not None else "") + "</c>"

    def xml(self):
        rows = {}
        for (r, c), (v, s, kw) in sorted(self.cells.items()):
            rows.setdefault(r, []).append(self._cell(r, c, v, s, kw))
        for r in self.rowattrs:
            rows.setdefault(r, [])
        data = "".join(f'<row r="{r}"' + (f" {self.rowattrs[r]}" if r in self.rowattrs else "") + (">" + "".join(cs) + "</row>" if cs else "/>") for r, cs in sorted(rows.items()))
        x = dict(self.x)
        x["sheetData"] = f"<sheetData>{data}</sheetData>" if data else "<sheetData/>"
        body = "".join(x[k] for k in ORDER if k in x)
        return f'{HEAD}<worksheet xmlns="{M}" xmlns:r="{R}"{self.ns}>{body}</worksheet>'


class Book:
    def __init__(self, styles=None, minimal=False, title="Fixture"):
        self.sheets, self.strings, self.styles = [], [], styles or Styles()
        self.parts, self.wbrels, self.ct, self.names = {}, Rels(), {}, []
        self.wb = {"bookViews": '<bookViews><workbookView/></bookViews>', "calcPr": '<calcPr calcId="191029"/>'}
        self.minimal, self.title, self.ns, self.ext_ct = minimal, title, "", {}
        self.root_rels = Rels()

    def sheet(self, name, state=None):
        sh = Sheet(name, state, self)
        self.sheets.append(sh)
        return sh

    def sst(self, text):
        if text not in self.strings:
            self.strings.append(text)
        return self.strings.index(text)

    def add(self, name, data, ct=None):
        self.parts[name] = data.encode("utf-8") if isinstance(data, str) else data
        if ct:
            self.ct[name] = CT.get(ct, ct)

    def build(self):
        P = dict(self.parts)
        ct = dict(self.ct)
        sheets = ""
        for i, sh in enumerate(self.sheets, 1):
            rid = self.wbrels.add("worksheet", f"worksheets/sheet{i}.xml")
            sheets += f'<sheet name="{esc(sh.name)}" sheetId="{i}"' + (f' state="{sh.state}"' if sh.state else "") + f' r:id="{rid}"/>'
            P[f"xl/worksheets/sheet{i}.xml"] = sh.xml().encode()
            ct[f"xl/worksheets/sheet{i}.xml"] = CT["worksheet"]
            if sh.rels.items:
                P[f"xl/worksheets/_rels/sheet{i}.xml.rels"] = sh.rels.xml().encode()
            for n, (d, c) in sh.parts.items():
                P[n] = d.encode() if isinstance(d, str) else d
                if c:
                    ct[n] = CT.get(c, c)
        if not self.minimal:
            self.wbrels.add("styles", "styles.xml")
            P["xl/styles.xml"], ct["xl/styles.xml"] = self.styles.xml().encode(), CT["styles"]
            if self.strings:
                self.wbrels.add("sharedStrings", "sharedStrings.xml")
                si = "".join(s if s.startswith("<si>") else f'<si><t xml:space="preserve">{esc(s)}</t></si>' for s in self.strings)
                P["xl/sharedStrings.xml"] = f'{HEAD}<sst xmlns="{M}" count="{len(self.strings)}" uniqueCount="{len(self.strings)}">{si}</sst>'.encode()
                ct["xl/sharedStrings.xml"] = CT["sst"]
        names = f"<definedNames>{''.join(self.names)}</definedNames>" if self.names else ""
        wb = self.wb
        body = wb.get("pr", "") + wb.get("protection", "") + (wb["bookViews"] if not self.minimal else "") + f"<sheets>{sheets}</sheets>" + wb.get("externalReferences", "") + names + (wb["calcPr"] if not self.minimal else "") + wb.get("pivotCaches", "") + wb.get("extLst", "")
        P["xl/workbook.xml"] = f'{HEAD}<workbook xmlns="{M}" xmlns:r="{R}"{self.ns}>{body}</workbook>'.encode()
        ct["xl/workbook.xml"] = self.ext_ct.get("workbook", CT["workbook"])
        P["xl/_rels/workbook.xml.rels"] = self.wbrels.xml().encode()
        self.root_rels.add("officeDocument", "xl/workbook.xml")
        self.root_rels.add("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml")
        self.root_rels.add(REL + "extended-properties", "docProps/app.xml")
        P["_rels/.rels"] = self.root_rels.xml().encode()
        P["docProps/core.xml"], ct["docProps/core.xml"] = core_xml(self.title).encode(), CT["core"]
        P["docProps/app.xml"], ct["docProps/app.xml"] = app_xml().encode(), CT["app"]
        P["[Content_Types].xml"] = content_types(P, ct).encode()
        return P

    def write(self, path):
        write_zip(path, self.build())


def dynamic_array_metadata(book):
    """xl/metadata.xml with the XLDAPR dynamic-array flag; cells opt in with cm="1"."""
    book.add("xl/metadata.xml", f'{HEAD}<metadata xmlns="{M}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
             '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
             '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
             '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>', SS + "sheetMetadata+xml")
    book.wbrels.add("sheetMetadata", "metadata.xml")


def core_xml(title="Fixture"):
    return (f'{HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" '
            f'xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>{esc(title)}</dc:title>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{ISO}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{ISO}</dcterms:modified></cp:coreProperties>')


def app_xml(application="Spindle fixture generator"):
    return f'{HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>{esc(application)}</Application></Properties>'


def content_types(parts, overrides):
    d = '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    exts = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif", "svg": "image/svg+xml", "emf": "image/x-emf",
            "bin": "application/vnd.ms-office.vbaProject", "vml": "application/vnd.openxmlformats-officedocument.vmlDrawing"}
    for e in sorted({n.rsplit(".", 1)[-1].lower() for n in parts if "." in n.rsplit("/", 1)[-1]} & set(exts)):
        d += f'<Default Extension="{e}" ContentType="{exts[e]}"/>'
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


def png(w=8, h=8, fg=(30, 90, 200), bg=(240, 240, 240)):
    """Original two-colour PNG (diagonal), no third-party pixels."""
    raw = b"".join(b"\x00" + b"".join(bytes(fg if abs(x - y) < 2 else bg) for x in range(w)) for y in range(h))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b"")


def theme_xml():
    clr = "".join(f'<a:{n}><a:srgbClr val="{v}"/></a:{n}>' for n, v in [("dk2", "44546A"), ("lt2", "E7E6E6"), ("accent1", "4472C4"), ("accent2", "ED7D31"), ("accent3", "A5A5A5"),
                                                                       ("accent4", "FFC000"), ("accent5", "5B9BD5"), ("accent6", "70AD47"), ("hlink", "0563C1"), ("folHlink", "954F72")])
    ln = '<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    fills = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' * 3
    return (f'{HEAD}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Spindle Test Theme"><a:themeElements>'
            f'<a:clrScheme name="Spindle"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>{clr}</a:clrScheme>'
            '<a:fontScheme name="Spindle"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
            f'<a:fmtScheme name="Spindle"><a:fillStyleLst>{fills}</a:fillStyleLst><a:lnStyleLst>{ln * 3}</a:lnStyleLst>'
            f'<a:effectStyleLst>{"<a:effectStyle><a:effectLst/></a:effectStyle>" * 3}</a:effectStyleLst><a:bgFillStyleLst>{fills}</a:bgFillStyleLst></a:fmtScheme>'
            '</a:themeElements></a:theme>')
