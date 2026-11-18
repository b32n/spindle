"""LibreOffice-produced fixtures.

A programmatic DOCX (built with the same generators as the other fixtures) is opened and re-saved
by LibreOffice running headless, so every part of the result (styles, numbering, settings, drawing
markup, comments, ...) is LibreOffice's own serialisation. The output is committed as a binary:
LibreOffice output is not byte-reproducible across versions, so rebuilding is only done on request.

    python3 fixtures/docx/generators/build_libreoffice.py            # needs `soffice` on PATH

The exact command (private profile directory so no lock file is shared with a running instance):

    soffice -env:UserInstallation=file:///tmp/spindle-lo-profile --headless --norestore --nolockcheck \
        --convert-to docx:"MS Word 2007 XML" --outdir <outdir> <input>.docx
"""
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import features as F  # noqa: E402
import features2 as F2  # noqa: E402
import features3 as F3  # noqa: E402
from docbuilder import PEOPLE, Doc, paragraph_text  # noqa: E402
from ooxml import fld, p, r, sect, tab, tbl, tc, tr  # noqa: E402


def soffice_version():
    return subprocess.run(["soffice", "--version"], capture_output=True, text=True).stdout.strip()


def resave(src, dst):
    tmp = tempfile.mkdtemp(prefix="spindle-lo-")
    try:
        subprocess.run(
            ["soffice", "-env:UserInstallation=file://" + os.path.join(tmp, "profile"), "--headless", "--norestore", "--nolockcheck",
             "--convert-to", "docx:MS Word 2007 XML", "--outdir", os.path.join(tmp, "out"), src],
            check=True, capture_output=True, timeout=240)
        shutil.copy(os.path.join(tmp, "out", os.path.basename(src)), dst)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def base_complex(path):
    doc = Doc("Harbour works programme")

    def s0():
        return (p(r("Harbour Works Programme", sz=48, b=True), style="Title") + p(r("Styles, lists, tables, images, headers, footers and footnotes, saved by LibreOffice"), style="Subtitle")
                + F.toc_fields(doc, [(1, "Programme overview"), (2, "Objectives"), (2, "Budget"), (1, "Schedule"), (1, "Risks")]) + F.notes(doc))

    def s1():
        out = F.H("Lists and styles", 1)
        for lvl, t in [(0, "Survey"), (1, "North quay"), (2, "Bollards"), (2, "Fenders"), (1, "East quay"), (0, "Design"), (1, "Drawings"), (0, "Build")]:
            out += p(r(t), style="ListParagraph", ppr=f'<w:numPr><w:ilvl w:val="{lvl}"/><w:numId w:val="1"/></w:numPr>')
        for t in ("Tender", "Award", "Mobilise"):
            out += p(r(t), style="ListParagraph", ppr='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr>')
        out += p(r(paragraph_text(700, 3)), style="Quote") + p(r(paragraph_text(701, 2)), style="IntenseQuote")
        out += p(r("Accent words, ") + r("strong", style="Strong") + r(", ") + r("emphasis", style="Emphasis") + r(", ") + r("accent", style="Accent") + r(" and a link ") + doc.link("https://example.com/works", "example.com/works") + r("."), style="BodyText")
        return out + F.body_p(702, 6) + F.body_p(703, 6)

    def s2():
        return F.tables(doc) + F2.images(doc)

    body, final = F.sections(doc, [s0, s1, s2, None])
    doc.finish(body, final, path)


def base_review(path):
    doc = Doc("Review of the harbour bye-laws", creator="Alice Archer")
    doc.settings_extra = "<w:trackRevisions/>"
    body = p(r("Review of the Harbour Bye-laws"), style="Title") + p(r("Tracked changes and comments recorded for a LibreOffice round trip"), style="Subtitle")
    body += F.H("Mooring", 1)
    c0, p0 = doc.comment(0, 8, "Is 48 hours long enough for visiting vessels?")
    doc.comment(1, 9, "Council guidance says 72 hours.", parent=p0)
    doc.comment(2, 10, "Agreed, changing it.", parent=p0, done=True)
    body += p(r("Visiting vessels may moor for up to ") + doc.dele(0, "48", 8) + doc.ins(1, "72", 9) + doc.commented([0, 1, 2], r(" hours")) + r(" without a permit. ") + doc.ins(2, "Permits are issued by the harbour office. ", 10) + r(paragraph_text(710, 2)), style="BodyText")
    c3, p3 = doc.comment(3, 11, "Cite the 2019 statute here.")
    body += p(doc.commented([3], r("Fees are set annually")) + doc.dele(1, " by the council", 9) + doc.ins(3, " by the harbour board", 9) + r(". ") + r(paragraph_text(711, 2)), style="BodyText")
    body += p(doc.dele(2, "This paragraph is struck out entirely by the third reviewer.", 10), style="BodyText")
    body += p(doc.ins(0, "This paragraph is new: night mooring requires a lit stern light.", 10), style="BodyText")
    body += F.H("Fishing", 1)
    body += p(r("Registered vessels ") + doc.ins(1, "and day-trip operators ", 11, b=True) + r("may land catch at the east quay. ") + doc.dele(0, "Auctions start at six. ", 11) + doc.ins(0, "Auctions start at seven. ", 11) + r(paragraph_text(712, 2)), style="BodyText")
    rows = [tr([tc(p(r("Vessel class", b=True)), w=3000, shade="D9E2F3"), tc(p(r("Fee", b=True)), w=3000, shade="D9E2F3")], header=True),
            tr([tc("Small", w=3000), tc(p(doc.dele(0, "10", 12) + doc.ins(0, "12", 12)), w=3000)]),
            tr([tc("Medium", w=3000), tc("25", w=3000)])]
    body += tbl(rows, [3000, 3000]) + p("")
    body += p(r("List of required documents:"), style="BodyText")
    for t, mark in (("Registration certificate", None), ("Insurance schedule", "ins"), ("Old permit form", "del"), ("Crew list", None)):
        inner = r(t) if mark is None else (doc.ins(3, t, 12) if mark == "ins" else doc.dele(3, t, 12))
        body += p(inner, style="ListParagraph", ppr='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>')
    body += F.body_p(713, 4)
    doc.finish(body, sect(), path)


def main():
    tmp = tempfile.mkdtemp(prefix="spindle-lo-base-")
    try:
        for name, base in (("libreoffice-complex", base_complex), ("libreoffice-review", base_review)):
            src = os.path.join(tmp, f"{name}.docx")
            base(src)
            dst = os.path.join(OUT, f"{name}.docx")
            resave(src, dst)
            print(f"wrote {dst} ({os.path.getsize(dst)} bytes) with {soffice_version()}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
