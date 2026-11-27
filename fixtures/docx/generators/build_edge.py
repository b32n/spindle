"""Edge-case packages. The harness must never crash on these; importFromDocx may throw and that is snapshotted.

edge-malformed      valid main document, but broken package plumbing and corrupt optional parts
edge-odd-valid      structurally valid but unusual markup that real files contain
edge-no-body        a package whose main document part is missing (the import is expected to throw)
edge-zip-bomb       a valid document plus one part that inflates 1000:1 to 24 MB (refused by the import limits)
edge-entity-bomb    a main document with a billion-laughs and an external-entity DOCTYPE (imports, entities stay inert)
"""
import features as F
import features2 as F2
from docbuilder import PEOPLE, Doc, paragraph_text
from ooxml import NS_DECL, fld, p, r, sect, tab, tbl, tc, tr, write_zip


def edge_malformed(out):
    doc = Doc("Malformed package")
    rid_ok = doc.link("https://example.com/ok", "a working link")
    body = p(r("Malformed package"), style="Title")
    body += p(r("Everything in this file is plausible except the package plumbing: "), style="BodyText")
    body += p(r("A link that works: ") + doc.link("https://example.org/works", "works") + r(", and one whose relationship id does not exist: ")
              + '<w:hyperlink r:id="rId999" w:history="1">' + r("dangling hyperlink", style="Hyperlink") + "</w:hyperlink>" + r("."), style="BodyText")
    # image with a relationship that has no target part, and one with an unknown rId
    rid_img = doc.rel("image", "media/missing.png")
    from ooxml import inline_image
    body += p(r("Image whose media part is missing: ") + inline_image(rid_img, 77, 100, 60, "missing media") + r(" and an image with an unknown relationship id: ") + inline_image("rId888", 78, 100, 60, "unknown rel"), style="BodyText")
    body += p(r("A style that does not exist."), style="NoSuchStyle")
    body += p(r("A list item pointing at a numbering definition that does not exist."), style="ListParagraph", ppr='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="77"/></w:numPr>')
    body += p(r("A footnote whose part is corrupt.") + '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>', style="BodyText")
    c, _ = doc.comment(0, 9, "This comment part is truncated on disk.")
    body += p(doc.commented([c], r("A comment on this phrase.")), style="BodyText")
    body += p(r("A table with a row that has more cells than the grid:"), style="BodyText")
    body += tbl([tr([tc("a", w=1500), tc("b", w=1500), tc("c-extra", w=1500)]), tr([tc("d", w=1500)])], [1500, 1500])
    body += p(r("A header reference to a part that is not in the package follows in the section properties."), style="BodyText")
    doc.fn = [(1, doc.note_para("footnote", "Footnote text"))]
    final = sect(refs='<w:headerReference w:type="default" r:id="rId777"/><w:footerReference w:type="default" r:id="rId778"/>')

    def mutate(d):
        # optional parts removed
        for n in ("word/styles.xml", "word/settings.xml", "word/numbering.xml", "word/fontTable.xml", "word/theme/theme1.xml"):
            d.parts.pop(n, None)
            d.overrides.pop("/" + n, None)
        d.doc_rels[:] = [x for x in d.doc_rels if not x[2].split("/")[-1] in ("styles.xml", "settings.xml", "numbering.xml", "fontTable.xml", "theme1.xml")]
        # corrupt optional parts: truncated mid-element, and an unterminated entity / tag soup
        d.parts["word/comments.xml"] = d.parts["word/comments.xml"][: len(d.parts["word/comments.xml"]) // 2]
        d.parts["word/footnotes.xml"] = b'<?xml version="1.0"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>Broken &bogus; <b></w:t></w:r></w:p></w:footnote>'
        d.parts["word/commentsExtended.xml"] = b"not xml at all \x00\x01\x02"
        # relationships: duplicate id, target-less relationship, dangling header/footer relationships
        d.doc_rels.append(("rId777", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header", "header_gone.xml", False))
        d.doc_rels.append(("rId778", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", "footer_gone.xml", False))
        d.doc_rels.append((d.doc_rels[0][0], "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "media/dup.png", False))
        # content types: override for a part that does not exist
        d.overrides["/word/ghost.xml"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        # stray entries
        d.parts["word/junk.bin"] = bytes(range(256)) * 4
        d.parts["word/media/zero.png"] = b""
        d.parts["customXml/item9.xml"] = b"<broken"

    doc.mutate = mutate
    doc.finish(body, final, out)


def edge_odd_valid(out):
    doc = Doc("Odd but valid")
    ind = 0

    def nested_sdt(depth, inner):
        for d in range(depth, 0, -1):
            inner = f'<w:sdt><w:sdtPr><w:alias w:val="level {d}"/><w:tag w:val="deep{d}"/><w:id w:val="{4000 + d}"/></w:sdtPr><w:sdtContent>{inner}</w:sdtContent></w:sdt>'
        return inner

    body = p("")  # leading empty paragraph
    body += p(r("   ", ) + r(""), style="BodyText")  # whitespace-only run, empty run
    body += p(r("Hidden text follows: ") + r("this sentence is hidden (w:vanish)", vanish=True) + r(" and this one is visible again."), style="BodyText")
    body += p(r("Special characters: ") + "<w:r><w:noBreakHyphen/></w:r>" + r(" non-breaking hyphen, ") + "<w:r><w:ptab w:relativeTo=\"margin\" w:alignment=\"right\" w:leader=\"dot\"/></w:r>" + r("position tab, ") + "<w:r><w:br w:type=\"textWrapping\" w:clear=\"all\"/></w:r>" + r("line break with clear, ") + r("NBSP here, zero-width​space, tab") + tab() + r("end."), style="BodyText")
    body += p(r("A 1,200-character unbroken token: " + "x" * 1200), style="BodyText")
    body += p("<w:permStart w:id=\"1\" w:edGrp=\"everyone\"/>" + r("Permission range (editable region)") + "<w:permEnd w:id=\"1\"/>", style="BodyText")
    body += p('<w:customXml w:uri="urn:spindle:odd" w:element="tag">' + r("Text wrapped in a custom XML element, ") + "</w:customXml>" + '<w:bdo w:val="rtl">' + r("bidi override run") + "</w:bdo>" + '<w:dir w:val="rtl">' + r(" directional run") + "</w:dir>", style="BodyText")
    # field in a hyperlink in a tracked insertion; hyperlink in a field
    ins = f'<w:ins w:id="{doc.nid()}" w:author="{PEOPLE[2]}" w:date="2026-01-06T10:00:00Z"><w:hyperlink r:id="{doc.rel("hyperlink", "https://example.com/odd", external=True)}">{fld("PAGE", "1", style="Hyperlink")}</w:hyperlink></w:ins>'
    body += p(r("Field in a hyperlink in a tracked insertion: ") + ins, style="BodyText")
    # deeply nested SDT wrapping a table whose cell has another SDT
    cell_inner = nested_sdt(3, p(r("Cell text inside nested controls"))) 
    tab_in_sdt = tbl([tr([tc(cell_inner, w=4000), tc("plain", w=4000)])], [4000, 4000])
    body += nested_sdt(10, p(r("Ten controls deep, wrapping a paragraph and a table:")) + tab_in_sdt)
    # revision inside a table inside a text box (anchored DrawingML text box)
    rev_tbl = tbl([tr([tc(p(doc.ins(1, "inserted inside a text-box table", 7)), w=2400), tc(p(doc.dele(2, "deleted inside a text-box table", 7)), w=2400)])], [2400, 2400])
    box = p(r("Text box containing a table with revisions:", b=True)) + rev_tbl + p("")
    wsp = F2._wsp("tb", 1, 0, 0, 340, 130, "rect", "F2F2F2", box)
    g = f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">{wsp}</a:graphicData></a:graphic>'
    body += p(r("Anchored text box: ") + f'<w:r><w:drawing>{F2._anchor_drawing(doc, g, 340, 130, "Odd Text Box", halign="center", vpos=0, wrap="<wp:wrapTopAndBottom/>")}</w:drawing></w:r>', style="BodyText")
    # AlternateContent with an unknown Requires and a fallback
    body += p(r("Alternate content requiring an unknown namespace: ") + '<w:r><mc:AlternateContent><mc:Choice Requires="zz"><zz:future xmlns:zz="urn:spindle:unknown" zz:attr="1"/></mc:Choice><mc:Fallback>' + r("fallback text for readers without zz") + "</mc:Fallback></mc:AlternateContent></w:r>", style="BodyText")
    # odd tables: gridSpan beyond grid, vMerge continue without restart, zero-width column, empty grid
    body += tbl([tr([tc("span 5 of 3", w=4000, span=5)]), tr([tc("", w=1000, vmerge="cont"), tc("orphan vMerge continue", w=1000), tc("", w=0)]), tr([tc("", w=1000), tc("", w=1000), tc("", w=1000)])], [1000, 1000, 1000])
    body += '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>table with an empty grid</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    body += p("")
    # altChunk (embedded HTML)
    doc.parts["word/afchunk.htm"] = b"<html><body><h1>Embedded HTML chunk</h1><p>Imported through w:altChunk.</p></body></html>"
    doc.defaults["htm"] = "text/html"
    rid = doc.rel("aFChunk", "afchunk.htm")
    body += f'<w:altChunk r:id="{rid}"/>'
    body += p(r("A paragraph with an extremely long run of mixed scripts: ") + r("Latin Ελληνικά Кириллица العربية 日本語 한국어 हिन्दी ไทย ") * 8, style="BodyText")
    # header/footer: content lives here too (tables, fields, text)
    h = doc.header_footer("header", p(r("Content in the header: ") + fld("TITLE", "Odd but valid") + tab() + r("Page ") + fld("PAGE", "1"), style="Header")
                          + tbl([tr([tc("header table L", w=3000), tc("header table R", w=3000)])], [3000, 3000]) + p(""), "header1")
    f = doc.header_footer("footer", tbl([tr([tc(p(r("Footer in a table; ") + fld("NUMPAGES", "3")), w=6000)])], [6000]) + p(""), "footer1")
    final = sect(refs=f'<w:headerReference w:type="default" r:id="{h}"/><w:footerReference w:type="default" r:id="{f}"/>')
    # custom XML datastore that nothing references
    doc.parts["customXml/item1.xml"] = '<?xml version="1.0" encoding="UTF-8"?><odd:root xmlns:odd="urn:spindle:odd"><odd:a>1</odd:a><odd:b><odd:c>nested</odd:c></odd:b></odd:root>'.encode()
    doc.parts["customXml/itemProps1.xml"] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><ds:datastoreItem ds:itemID="{AAAAAAAA-0000-4000-8000-000000000001}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"/>'
    doc.parts["customXml/_rels/item1.xml.rels"] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>'
    doc.overrides["/customXml/itemProps1.xml"] = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
    doc.rel("customXml", "../customXml/item1.xml")

    def mutate(d):
        d.parts["word/document.xml"] = d.parts["word/document.xml"].replace(b'mc:Ignorable="w14 w15"', b'xmlns:zz="urn:spindle:unknown" mc:Ignorable="w14 w15 zz" zz:note="ignorable attribute"', 1)

    doc.mutate = mutate
    doc.finish(body, final, out)


def edge_no_body(out):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    from ooxml import styles_xml
    write_zip(out, {"[Content_Types].xml": ct.encode(), "_rels/.rels": rels.encode(), "word/styles.xml": styles_xml().encode(),
                    "docProps/core.xml": b'<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>No body</dc:title></cp:coreProperties>'})


def edge_zip_bomb(out):
    """~32 KB on disk, 24 MB inflated: over the 200:1 ratio limit (and the 16 MB floor below which the ratio is not judged)."""
    doc = Doc("Zip bomb")
    body = p(r("Zip bomb"), style="Title")
    body += p(r("An ordinary document with one extra part that inflates a thousandfold. The import limits refuse the package before inflating it."), style="BodyText")
    doc.parts["word/media/bomb.bin"] = bytes(24 * 1024 * 1024)
    doc.finish(body, sect(), out)


# Nine levels of ten references each: 10**9 "lol" if anything expanded it. The external entity points at a local file.
ENTITY_DOCTYPE = (
    '<!DOCTYPE w:document [<!ENTITY lol "lol">'
    + "".join(f'<!ENTITY lol{i} "{("&lol%s;" % (i - 1 or "")) * 10}">' for i in range(1, 10))
    + '<!ENTITY xxe SYSTEM "file:///etc/passwd">]>'
)


def edge_entity_bomb(out):
    doc = Doc("XML entity expansion")
    body = p(r("XML entity expansion"), style="Title")
    body += p(r("The main part of this file declares a billion-laughs entity chain and an external entity. Neither may be expanded."), style="BodyText")
    # a raw run: r() would escape the references
    body += p('<w:r><w:t xml:space="preserve">Chain: &lol9; External: &xxe; Predefined: &amp; &lt;b&gt;.</w:t></w:r>', style="BodyText")
    body += p(r("A paragraph after the entities, to show the import still reads the document."), style="BodyText")

    def mutate(d):
        xml = d.parts["word/document.xml"]
        head = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        assert xml.startswith(head), xml[:80]
        d.parts["word/document.xml"] = head + ENTITY_DOCTYPE.encode() + xml[len(head):]

    doc.mutate = mutate
    doc.finish(body, sect(), out)
