"""Edge-case packages: broken, odd and hostile containers (all tiny; the large ones are build_runtime.py)."""
import struct

import xlsxbuilder as X


def truncated_zip(path):
    import io
    import zipfile

    b = X.Book(title="Truncation source")
    sh = b.sheet("Sheet1")
    for r in range(1, 60):
        sh.row(r, [r, f"row {r}", r * 1.5, f"text {r} " * 4])
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in sorted(b.build().items(), key=lambda kv: (kv[0] != "[Content_Types].xml", kv[0])):
            zi = zipfile.ZipInfo(n, date_time=X.STAMP)
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, d)
    data = buf.getvalue()
    open(path, "wb").write(data[: int(len(data) * 0.6)])


def no_workbook(path):
    b = X.Book(title="No workbook")
    P = b.build()
    for n in [n for n in P if n.startswith("xl/")]:
        del P[n]
    P["_rels/.rels"] = ('<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns="' + X.PKG_REL + '"><Relationship Id="rId1" Type="' + X.REL
                        + 'officeDocument" Target="xl/workbook.xml"/></Relationships>').encode()
    P["[Content_Types].xml"] = X.content_types(P, {"docProps/core.xml": X.CT["core"], "docProps/app.xml": X.CT["app"]}).encode()
    X.write_zip(path, P)


def bad_rels(path):
    b = X.Book(title="Bad relationships")
    b.sheet("Present").set(1, 1, "this sheet exists")
    b.wbrels.add("worksheet", "worksheets/missing.xml")
    b.wbrels.add("theme", "theme/theme9.xml")
    b.wbrels.add("sharedStrings", "../outside/strings.xml")
    P = b.build()
    wb = P["xl/workbook.xml"].decode().replace("</sheets>", '<sheet name="Dangling" sheetId="2" r:id="rId99"/></sheets>')
    P["xl/workbook.xml"] = wb.encode()
    X.write_zip(path, P)


def xml_entities(path):
    b = X.Book(title="XML entities")
    sh = b.sheet("Entities")
    sh.set(1, 1, ("i", "&lol3;")).set(2, 1, ("i", "&ext;")).set(3, 1, "plain")
    P = b.build()
    doc = ('<!DOCTYPE worksheet [<!ENTITY lol "lol"><!ENTITY lol1 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">'
           '<!ENTITY lol2 "&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;"><!ENTITY lol3 "&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;">'
           '<!ENTITY ext SYSTEM "file:///nonexistent/spindle-fixture-entity.txt">]>')
    s = P["xl/worksheets/sheet1.xml"].decode()
    P["xl/worksheets/sheet1.xml"] = s.replace("<worksheet ", doc + "<worksheet ", 1).encode()
    X.write_zip(path, P)


def _ole(path, streams):
    """An OLE compound-file header plus a directory sector naming the streams: a stub, not a usable compound file."""
    head = bytes([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]) + bytes(8) + struct.pack("<HHH", 0x3E, 3, 0xFFFE) + bytes(488)
    entries = b""
    for name in ["Root Entry"] + streams:
        raw = name.encode("utf-16-le")
        entries += raw.ljust(64, b"\0") + struct.pack("<H", len(raw) + 2) + bytes(62)
    open(path, "wb").write(head + entries.ljust(512, b"\0"))


def encrypted(path):
    _ole(path, ["EncryptionInfo", "EncryptedPackage"])


def legacy_xls(path):
    _ole(path, ["Workbook", "SummaryInformation"])


def duplicate_sheet_names(path):
    b = X.Book(title="Duplicate sheet names")
    for n in ("Data", "Data", "DATA"):
        b.sheet(n).set(1, 1, f"sheet {len(b.sheets)}")
    b.write(path)


def index_out_of_range(path):
    b = X.Book(title="Indexes out of range")
    sh = b.sheet("Indexes")
    sh.set(1, 1, "valid string").set(1, 2, ("raw", "<v>7</v>"), t="s").set(1, 3, 5, s=9).set(1, 4, ("raw", "<v>1234</v>"), t="s")
    b.write(path)


def empty_workbook(path):
    b = X.Book(minimal=True, title="Empty workbook")
    b.sheet("Sheet1")
    b.write(path)


def invalid_xml(path):
    b = X.Book(title="Invalid XML")
    b.sheet("Broken").row(1, ["a", "b", "c"])
    P = b.build()
    s = P["xl/worksheets/sheet1.xml"].decode()
    P["xl/worksheets/sheet1.xml"] = s.replace("</sheetData>", "</row></sheetData>", 1).replace("<row r=\"1\">", "<row r=\"1\"><c r=\"Z1\"><v>1</c>", 1).encode()
    X.write_zip(path, P)
