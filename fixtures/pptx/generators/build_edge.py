"""Edge-case packages: broken, odd and hostile containers (all tiny; the large one is build_runtime.py)."""
import io
import struct
import zipfile

import pptxbuilder as X
from dml import grp, para, run, sp, title


def _deck(title_text="Edge source", slides=3):
    d = X.Deck(title=title_text)
    for n in range(1, slides + 1):
        s = d.slide(1)
        s.add(title(2, f"Slide {n}"))
        s.add(sp(3, "Box", 914400, 2000000, 3000000, 1000000, paras=para(run(f"Content of slide {n}"))))
    return d


def truncated_zip(path):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in sorted(_deck("Truncation source", 8).build().items(), key=lambda kv: (kv[0] != "[Content_Types].xml", kv[0])):
            zi = zipfile.ZipInfo(n, date_time=X.STAMP)
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, data)
    data = buf.getvalue()
    open(path, "wb").write(data[: int(len(data) * 0.6)])


def no_presentation(path):
    parts = _deck("No presentation part").build()
    for n in ("ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"):
        del parts[n]
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(b'<Override PartName="/ppt/presentation.xml" ContentType="' + X.CT["presentation"].encode() + b'"/>', b"")
    X.write_zip(path, parts)


def slide_missing(path):
    parts = _deck("Missing slide part").build()
    for n in ("ppt/slides/slide2.xml", "ppt/slides/_rels/slide2.xml.rels"):
        del parts[n]
    X.write_zip(path, parts)


def bad_rels(path):
    d = _deck("Bad relationships", 2)
    s = d.slides[0]
    s.rels.add("image", "../media/missing.png")
    s.rels.add("slideLayout", "../slideLayouts/slideLayout99.xml")
    s.rels.add("oleObject", "../../outside/object.bin")
    d.slides[1].rels.add("hyperlink", "https://example.com/external-is-fine", external=True)
    d.write(path)


def invalid_xml(path):
    parts = _deck("Invalid XML", 3).build()
    s = parts["ppt/slides/slide2.xml"].decode()
    parts["ppt/slides/slide2.xml"] = s.replace("</p:spTree>", "</p:sp></p:spTree>", 1).replace("<p:cSld>", "<p:cSld><p:bg><p:bgPr><a:noFill></p:bgPr></p:bg>", 1).encode()
    X.write_zip(path, parts)


def deep_groups(path, depth=60):
    d = X.Deck(title="Deeply nested groups")
    s = d.slide(3)
    s.add(title(2, "Nested groups"))
    inner = sp(1000, "Innermost", 914400, 2000000, 1000000, 600000, paras=para(run("deep")))
    for n in range(depth):
        inner = grp(100 + n, f"Group {n}", 914400, 2000000, 1000000, 600000, inner)
    s.add(inner)
    d.write(path)


def relationship_cycles(path):
    d = _deck("Relationship cycles", 2)
    d.slides[0].rels.add("slide", "slide2.xml")
    d.slides[1].rels.add("slide", "slide1.xml")
    d.slides[1].rels.add("slide", "slide2.xml")
    d.write(path)


def _ole(path, streams):
    """An OLE compound-file header plus a directory sector naming the streams: a stub, not a usable compound file."""
    head = bytes([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]) + bytes(8) + struct.pack("<HHH", 0x3E, 3, 0xFFFE) + bytes(490)
    entries = b""
    for name in ["Root Entry"] + streams:
        raw = name.encode("utf-16-le")
        entries += raw.ljust(64, b"\0") + struct.pack("<H", len(raw) + 2) + bytes(62)
    open(path, "wb").write(head + entries.ljust(512, b"\0"))


def encrypted(path):
    _ole(path, ["EncryptionInfo", "EncryptedPackage"])


def legacy_ppt(path):
    _ole(path, ["PowerPoint Document", "Current User", "SummaryInformation"])
