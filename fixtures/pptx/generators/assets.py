"""Binary and package assets for the PPTX fixture generators: small original files, no third-party content.

The JPEG is a 16 x 16 picture drawn once with Pillow and stored below as base64 so that the generators stay stdlib-only
and byte-stable (a JPEG encoder's output is not guaranteed to be identical between versions).
"""
import base64
import struct
import zipfile
import io

import pptxbuilder as X

JPEG_B64 = (
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAA0JCgsKCA0LCgsODg0PEyAVExISEyccHhcgLikxMC4pLSwzOko+MzZGNywtQFdBRkxO"
    "UlNSMj5aYVpQYEpRUk//2wBDAQ4ODhMREyYVFSZPNS01T09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09P"
    "T09PT09PT0//wAARCAAQABADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUF"
    "BAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVW"
    "V1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi"
    "4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAEC"
    "AxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVm"
    "Z2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq"
    "8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDuv+PH/r0/9E//AGH/AKD9Pu26Kqf8eP8A16f+if8A7D/0H6fdWxfxep//2Q=="
)


def jpeg():
    return base64.b64decode(JPEG_B64)


def gif():
    """The classic 43-byte, 1 x 1 pixel GIF89a."""
    return bytes.fromhex("47494638396101000100800000000000ffffff21f90401000000002c00000000010001000002024401003b")


def svg(label="Spindle"):
    return ('<?xml version="1.0" encoding="UTF-8"?>\n<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64">'
            f'<rect width="64" height="64" fill="#e8eef8"/><circle cx="32" cy="32" r="22" fill="#4472c4"/><title>{label}</title></svg>').encode()


def emf():
    """An empty enhanced metafile: a header record and an end-of-file record (no drawing)."""
    header = struct.pack("<II", 1, 88) + struct.pack("<4i", 0, 0, 63, 63) + struct.pack("<4i", 0, 0, 1000, 1000) + struct.pack("<IIIIHHIII", 0x464D4520, 0x10000, 108, 2, 1, 0, 0, 0, 0) + struct.pack("<2i2i", 1, 1, 1, 1)
    eof = struct.pack("<IIIII", 14, 20, 0, 16, 20)
    return header + eof


def wmf():
    """An empty placeable Windows metafile: placeable header, standard header and an end record."""
    words = [0xCDD7, 0x9AC6, 0, 0, 0, 64, 64, 96, 0, 0]
    checksum = 0
    for w in words:
        checksum ^= w
    placeable = struct.pack("<11H", *words, checksum)
    standard = struct.pack("<HHHIHIH", 1, 9, 0x0300, 12, 0, 3, 0)
    return placeable + standard + struct.pack("<IH", 3, 0)


def wav():
    """A 100-sample, 8 kHz, 8-bit mono WAV of silence."""
    data = bytes([128]) * 100
    return b"RIFF" + struct.pack("<I", 36 + len(data)) + b"WAVEfmt " + struct.pack("<IHHIIHH", 16, 1, 1, 8000, 8000, 1, 8) + b"data" + struct.pack("<I", len(data)) + data


def mp4():
    """An inert MP4 stub: a file-type box and an empty media-data box (not a playable video)."""
    ftyp = struct.pack(">I4s4sI", 24, b"ftyp", b"isom", 512) + b"isomiso2"
    return ftyp + struct.pack(">I4s", 8, b"mdat")


def opaque_blob(label):
    """An inert placeholder for OLE and VBA payloads: readable text, never a real object or macro."""
    return (f"Spindle fidelity fixture: inert placeholder for {label}. This is not a real object or macro.\n").encode()


def mini_xlsx(rows):
    """A minimal workbook (one sheet, inline strings and numbers) for the workbook embedded in a chart."""
    cells = ""
    for r, row in enumerate(rows, 1):
        cs = ""
        for c, v in enumerate(row, 1):
            ref = f"{chr(64 + c)}{r}"
            cs += f'<c r="{ref}"><v>{v}</v></c>' if isinstance(v, (int, float)) else f'<c r="{ref}" t="inlineStr"><is><t>{X.esc(v)}</t></is></c>'
        cells += f'<row r="{r}">{cs}</row>'
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ss = "application/vnd.openxmlformats-officedocument.spreadsheetml."
    parts = {
        "[Content_Types].xml": X.content_types({"xl/workbook.xml": 0, "xl/worksheets/sheet1.xml": 0},
                                              {"xl/workbook.xml": ss + "sheet.main+xml", "xl/worksheets/sheet1.xml": ss + "worksheet+xml"}),
        "_rels/.rels": f'{X.HEAD}<Relationships xmlns="{X.PKG_REL}"><Relationship Id="rId1" Type="{X.REL}officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'{X.HEAD}<workbook xmlns="{main}" xmlns:r="{X.R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": f'{X.HEAD}<Relationships xmlns="{X.PKG_REL}"><Relationship Id="rId1" Type="{X.REL}worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": f'{X.HEAD}<worksheet xmlns="{main}"><sheetData>{cells}</sheetData></worksheet>',
    }
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for n in sorted(parts, key=lambda n: (n != "[Content_Types].xml", n)):
            zi = zipfile.ZipInfo(n, date_time=X.STAMP)
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, parts[n])
    return buf.getvalue()
