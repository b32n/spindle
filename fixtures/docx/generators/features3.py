"""Feature blocks: i18n / bidi / CJK / symbols, typography extras, charts and embedded objects, fonts and themes."""
import hashlib
import uuid
from docbuilder import Doc, paragraph_text
from features import H, body_p
from ooxml import CT, NS_DECL, esc, fld, p, r, rpr, tab, tbl, tc, tr, write_zip


# ================================================================ i18n
AR = ["هذا مستند اختبار لدعم الكتابة من اليمين إلى اليسار.", "يقيس الميناء حركة السفن كل صباح ويسجل النتائج في دفتر خاص.", "تُنشر جداول المد والجزر عند البوابة الشمالية."]
HE = ["שלום עולם, זהו מסמך בדיקה לכתיבה מימין לשמאל.", "ועדת הנמל מתכנסת בכל חודש כדי לסקור את לוחות הגאות והשפל."]
ZH = ["这是一个用于测试中文排版的示例段落，包含标点符号、数字 2026 和英文单词 Spindle。", "港口委员会每月召开一次会议，审议潮汐表和船舶进出记录。"]
JA = ["これは日本語の組版を確認するためのサンプル文章です。ひらがな、カタカナ、漢字が混在しています。", "港湾委員会は毎月会議を開き、潮汐表を確認します。"]
KO = ["이것은 한국어 조판을 확인하기 위한 예시 문장입니다.", "항만 위원회는 매달 모여 조석표를 검토합니다."]


def i18n(doc: Doc):
    out = [H("Right-to-left, CJK, symbols and emoji", 1)]
    rtl = lambda text, lang="ar-SA", **kw: p(r(text, rtl=True, lang=lang, font="Liberation Sans", **kw), ppr='<w:bidi/><w:jc w:val="left"/>')  # jc left == visually right in a bidi paragraph
    out.append(p("Arabic (right-to-left paragraphs, bidi + rtl runs):", style="BodyText"))
    for t in AR:
        out.append(rtl(t))
    out.append(p("Hebrew:", style="BodyText"))
    for t in HE:
        out.append(rtl(t, "he-IL"))
    out.append(p("Mixed direction in one paragraph (English inside Arabic, numbers and punctuation):", style="BodyText"))
    out.append(p(r("ميناء ", rtl=True, lang="ar-SA") + r("Marguerite Harbour", lang="en-US") + r(" استقبل ", rtl=True, lang="ar-SA") + r("42", lang="en-US") + r(" سفينة في 2026/01/15.", rtl=True, lang="ar-SA"), ppr='<w:bidi/>'))
    out.append(p(r("English sentence with an embedded Hebrew word ") + r("שלום", rtl=True, lang="he-IL") + r(" and an Arabic word ") + r("مرحبا", rtl=True, lang="ar-SA") + r(", then more English."), style="BodyText"))
    out.append(p(r("Arabic-Indic digits: ١٢٣٤٥٦٧٨٩٠ and Hebrew with niqqud: שָׁלוֹם", rtl=True, lang="ar-SA"), ppr='<w:bidi/>'))
    # bidi table and list
    out.append(tbl([tr([tc(p(r(AR[0], rtl=True, lang="ar-SA"), ppr='<w:bidi/>'), w=4500), tc(p(r("Right-to-left table cell")), w=4500)])], [4500, 4500], extra="<w:bidiVisual/>"))
    out.append(p(r(HE[1], rtl=True, lang="he-IL"), style="ListParagraph", ppr='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:bidi/>'))
    out.append(p("Chinese, Japanese and Korean with East Asian fonts:", style="BodyText"))
    for t, lang in [(ZH[0], "zh-CN"), (ZH[1], "zh-CN"), (JA[0], "ja-JP"), (JA[1], "ja-JP"), (KO[0], "ko-KR"), (KO[1], "ko-KR")]:
        out.append(p(r(t, font="Liberation Sans", eastAsia="Noto Sans CJK SC", lang="en-US", ea=lang), style="BodyText"))
    out.append(p(f'<w:r><w:rPr><w:rFonts w:eastAsia="Noto Serif CJK JP"/><w:lang w:eastAsia="ja-JP"/><w:eastAsianLayout w:id="1" w:combine="1"/></w:rPr><w:t>12</w:t></w:r>' + r("  (combined digits), ruby: ") +
                 '<w:r><w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="10"/><w:hpsRaise w:val="18"/><w:hpsBaseText w:val="20"/><w:lid w:val="ja-JP"/></w:rubyPr><w:rt><w:r><w:t>かんじ</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>漢字</w:t></w:r></w:rubyBase></w:ruby></w:r>', style="BodyText"))
    out.append(tbl([tr([tc(p(r("縦書きのテキスト"), ), w=900, extra='<w:textDirection w:val="tbRl"/>'), tc(p(r("Vertical (tbRl) text direction in the neighbouring cell.")), w=6000)], height=1800)], [900, 6000]))
    out.append(p("Symbols: " , style="BodyText"))
    out.append(p(r("Unicode symbols: © ® ™ € £ ¥ § ¶ † ‡ • … ← → ↔ ⇒ ∀ ∃ ∑ ∏ ∞ ≈ ≠ ≤ ≥ ± × ÷ √ ∫ ♠ ♣ ♥ ♦ ☀ ☂ ☎ ✓ ✔ ✘ ★", font="DejaVu Sans"), style="BodyText"))
    out.append(p(r("Word symbol runs: ") + '<w:r><w:sym w:font="Symbol" w:char="F0B7"/></w:r>' + r(" bullet, ") + '<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>' + r(" tick, ") + '<w:r><w:sym w:font="Wingdings" w:char="F04A"/></w:r>' + r(" smiley."), style="BodyText"))
    out.append(p(r("Emoji: 😀 🚀 🌍 ⚓ 🐟 🌊 ❤️ 👍🏽 👩‍💻 🏳️‍🌈 🇳🇴 1️⃣ — including ZWJ sequences, skin tones, flags and keycaps."), style="BodyText"))
    out.append(p(r("Combining marks and ligature candidates: é (e + acute), ñ, ﬁ ﬂ, Å (angstrom U+212B), ǆ, and a surrogate-pair math letter 𝒜𝔹ℂ𝟘."), style="BodyText"))
    return "".join(out)


# ============================================================ typography
def typography(doc: Doc):
    from ooxml import sect
    out = [H("Typography extras", 1)]
    doc.settings_mid += '<w:autoHyphenation/><w:consecutiveHyphenLimit w:val="2"/><w:hyphenationZone w:val="360"/><w:doNotHyphenateCaps/>'
    soft = "Anti" + "<w:r><w:softHyphen/></w:r>" + "dis" + "<w:r><w:softHyphen/></w:r>" + "establishment" + "<w:r><w:softHyphen/></w:r>" + "arian"
    out.append(p(r("Automatic hyphenation is enabled for this document, so long justified lines break inside words. ") + r(paragraph_text(260, 5) + " ") + r("Pneumonoultramicroscopicsilicovolcanoconiosis and incomprehensibilities are long words. ")
                 + r("Soft hyphens: ") + r("Anti") + "<w:r><w:softHyphen/></w:r>" + r("dis") + "<w:r><w:softHyphen/></w:r>" + r("establishment") + "<w:r><w:softHyphen/></w:r>" + r("arianism. ") + r("Non-breaking hyphen: ") + "<w:r><w:noBreakHyphen/></w:r>" + r("."), style="BodyText", ppr='<w:jc w:val="both"/>'))
    out.append(p(r("This paragraph opts out of hyphenation. " + paragraph_text(261, 3)), style="BodyText", ppr="<w:suppressAutoHyphens/>"))
    # tab leaders
    tabs = lambda stops: "<w:tabs>" + "".join(f'<w:tab w:val="{v}" w:leader="{l}" w:pos="{pos}"/>' for v, l, pos in stops) + "</w:tabs>"
    out.append(p("Tab stops with leaders:", style="BodyText"))
    for label, lead, page in (("Preface", "dot", "iii"), ("Chapter one", "hyphen", "1"), ("Chapter two", "underscore", "17"), ("Appendix", "middleDot", "203"), ("Index", "heavy", "240")):
        out.append(p(r(label) + tab() + r(page), style="Normal", ppr=tabs([("right", lead, 8000)])))
    out.append(p(r("Item") + tab() + r("12.50") + tab() + r("1,204.75") + tab() + r("end"), style="Normal", ppr=tabs([("left", "none", 1800), ("decimal", "dot", 4000), ("center", "underscore", 6500), ("right", "dot", 9000)])))
    # drop caps
    out.append(p("Drop caps (framed first letter, three lines):", style="BodyText"))
    cap = lambda ch, lines=3, kind="drop": p(r(ch, sz=170, b=True, color="2F5496"), ppr=f'<w:framePr w:dropCap="{kind}" w:lines="{lines}" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/><w:spacing w:after="0" w:line="{240 * lines}" w:lineRule="exact"/>')
    out.append(cap("T") + p(r("he lighthouse had stood on the point for two hundred years, and every keeper had added a line to the logbook. " + paragraph_text(262, 4)), style="BodyText"))
    out.append(cap("W", 2, "margin") + p(r("hen the committee last met, the margin drop cap hung outside the text block. " + paragraph_text(263, 3)), style="BodyText"))
    # other character-level typography
    out.append(p(r("Character effects: ") + r("small caps ", smallCaps=True) + r("ALL CAPS ", caps=True) + r("highlighted ", highlight="yellow") + r("coloured ", color="C00000")
                 + '<w:r><w:rPr><w:spacing w:val="40"/></w:rPr><w:t xml:space="preserve">expanded spacing </w:t></w:r><w:r><w:rPr><w:w w:val="70"/></w:rPr><w:t xml:space="preserve">condensed </w:t></w:r><w:r><w:rPr><w:position w:val="6"/></w:rPr><w:t xml:space="preserve">raised </w:t></w:r>'
                 + '<w:r><w:rPr><w:emboss/></w:rPr><w:t xml:space="preserve">embossed </w:t></w:r><w:r><w:rPr><w:shadow/><w:outline/></w:rPr><w:t xml:space="preserve">shadow-outline </w:t></w:r><w:r><w:rPr><w:u w:val="wave" w:color="FF0000"/></w:rPr><w:t>wavy underline</w:t></w:r>', style="BodyText"))
    out.append(p(r("Paragraph borders, shading and keep-with-next. " + paragraph_text(264, 2)), style="BodyText", ppr='<w:keepNext/><w:keepLines/><w:pBdr><w:top w:val="single" w:sz="6" w:space="4" w:color="2F5496"/><w:left w:val="single" w:sz="18" w:space="6" w:color="2F5496"/><w:bottom w:val="single" w:sz="6" w:space="4" w:color="2F5496"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="EAF1FB"/>'))
    out.append(p(r("Line numbering and page borders apply to this whole section."), style="BodyText"))
    return "".join(out)


def watermark_header(doc: Doc, text="DRAFT"):
    """Header part with the standard VML WordArt watermark ('PowerPlusWaterMarkObject') plus header text; returns the rId."""
    wm = ('<w:r><w:rPr><w:noProof/></w:rPr><w:pict>'
          '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/><v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>'
          f'<v:shape id="PowerPlusWaterMarkObject357476642" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:412.4pt;height:247.45pt;rotation:315;z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Liberation Sans&quot;;font-size:1pt" string="{text}"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r>')
    return doc.header_footer("header", p(wm + r("Harbour records — confidential draft"), style="Header"), f"header{doc.nid()}")


# ============================================================ charts etc
def _xlsx_bytes(rows):
    """A tiny but valid .xlsx (one sheet, inline strings) used as the chart's embedded workbook."""
    def cell(ref, v):
        return f'<c r="{ref}" t="inlineStr"><is><t>{esc(v)}</t></is></c>' if isinstance(v, str) else f'<c r="{ref}"><v>{v}</v></c>'
    sheet = "".join(f'<row r="{i + 1}">' + "".join(cell(f"{chr(65 + j)}{i + 1}", v) for j, v in enumerate(row)) + "</row>" for i, row in enumerate(rows))
    parts = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{sheet}</sheetData></worksheet>',
    }
    import io, zipfile
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n, c in parts.items():
            zi = zipfile.ZipInfo(n, date_time=(2026, 1, 15, 9, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, c)
    return buf.getvalue()


def _chart_xml(kind, title, cats, series, ext_rid=None):
    def ser(i, name, vals):
        col = chr(66 + i)
        cat = '<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$%d</c:f><c:strCache><c:ptCount val="%d"/>%s</c:strCache></c:strRef></c:cat>' % (len(cats) + 1, len(cats), "".join(f'<c:pt idx="{k}"><c:v>{c}</c:v></c:pt>' for k, c in enumerate(cats)))
        val = '<c:val><c:numRef><c:f>Sheet1!$%s$2:$%s$%d</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="%d"/>%s</c:numCache></c:numRef></c:val>' % (col, col, len(cats) + 1, len(cats), "".join(f'<c:pt idx="{k}"><c:v>{v}</c:v></c:pt>' for k, v in enumerate(vals)))
        tx = f'<c:tx><c:strRef><c:f>Sheet1!${col}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{name}</c:v></c:pt></c:strCache></c:strRef></c:tx>'
        fill = ["4472C4", "ED7D31", "A5A5A5"][i % 3]
        sp = f'<c:spPr><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></c:spPr>'
        return f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>{tx}{sp}{"<c:invertIfNegative val=\"0\"/>" if kind == "bar" else ""}{cat}{val}</c:ser>'
    sers = "".join(ser(i, n, v) for i, (n, v) in enumerate(series))
    if kind == "bar":
        plot = f'<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>{sers}<c:gapWidth val="150"/><c:axId val="111"/><c:axId val="222"/></c:barChart>'
        axes = '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx><c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:crossAx val="111"/></c:valAx>'
    elif kind == "line":
        plot = f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{sers}<c:marker val="1"/><c:axId val="111"/><c:axId val="222"/></c:lineChart>'
        axes = '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx><c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/></c:valAx>'
    else:
        plot = f'<c:pieChart><c:varyColors val="1"/>{sers}<c:firstSliceAng val="0"/></c:pieChart>'
        axes = ""
    ext = f'<c:externalData r:id="{ext_rid}"><c:autoUpdate val="0"/></c:externalData>' if ext_rid else ""
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace {NS_DECL}><c:roundedCorners val="0"/><c:chart>'
            f'<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>{title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
            f'<c:plotArea><c:layout/>{plot}{axes}</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>{ext}</c:chartSpace>')


def charts(doc: Doc):
    out = [H("Charts and embedded objects", 1)]
    cats = ["Jan", "Feb", "Mar", "Apr", "May"]
    specs = [("bar", "Vessels per month", [("2025", [12, 15, 9, 20, 18]), ("2026", [14, 17, 12, 22, 25])]),
             ("line", "Tonnage trend", [("Grain", [30, 35, 33, 40, 44]), ("Timber", [20, 18, 25, 22, 28])]),
             ("pie", "Cargo share", [("Share", [45, 25, 20, 10, 5])])]
    for k, (kind, title, series) in enumerate(specs, 1):
        rows = [[""] + [n for n, _ in series]] + [[c] + [v[i] for _, v in series] for i, c in enumerate(cats)]
        xrid = f"rIdX{k}"
        doc.parts[f"word/embeddings/Microsoft_Excel_Sheet{k}.xlsx"] = _xlsx_bytes(rows)
        doc.defaults["xlsx"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        doc.parts[f"word/charts/_rels/chart{k}.xml.rels"] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="{xrid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Sheet{k}.xlsx"/></Relationships>').encode()
        rid = doc.part(f"word/charts/chart{k}.xml", _chart_xml(kind, title, cats, series, xrid), CT["chart"], "chart")
        pid = doc.nid()
        out.append(p(f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="4572000" cy="2743200"/><wp:docPr id="{pid}" name="Chart {pid}" descr="{title}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="{rid}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>', ppr='<w:jc w:val="center"/>'))
        out.append(p(r(f"Chart {k}: {title}."), style="Caption", ppr='<w:jc w:val="center"/>'))
    # embedded OLE object stub (workbook) with preview image
    doc.parts["word/embeddings/oleObject1.xlsx"] = _xlsx_bytes([["Berth", "Count"], ["North", 12], ["East", 7]])
    ole_rid = doc.rel("oleObject", "embeddings/oleObject1.xlsx").replace("rId", "rId")
    # relationship type for embedded packages is .../package
    doc.doc_rels[-1] = (doc.doc_rels[-1][0], "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package", "embeddings/oleObject1.xlsx", False)
    prev, _ = doc.image_rid("rect", 90, 60, (40, 140, 90))
    out.append(p(r("Embedded worksheet object (OLE package with a preview image): ") +
                 f'<w:r><w:object w:dxaOrig="1440" w:dyaOrig="900"><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
                 f'<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:72pt;height:45pt" o:ole=""><v:imagedata r:id="{prev}" o:title=""/></v:shape>'
                 f'<o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1234567890" r:id="{ole_rid}"/></w:object></w:r>', style="BodyText"))
    return "".join(out)


# ======================================================= fonts and themes
def _obfuscate(font_bytes, guid):
    key = bytes.fromhex(guid.replace("-", "").upper())[::-1]
    head = bytes(b ^ key[i % 16] for i, b in enumerate(font_bytes[:32]))
    return head + font_bytes[32:]


def embed_font(doc: Doc, family, path, guid_seed):
    """Embed a TrueType font (obfuscated .odttf, as Word stores them) and reference it from fontTable.xml."""
    data = open(path, "rb").read()
    guid = str(uuid.UUID(hashlib.md5(guid_seed.encode()).hexdigest())).upper()
    doc.parts["word/fonts/font1.odttf"] = _obfuscate(data, guid)
    doc.defaults["odttf"] = "application/vnd.openxmlformats-officedocument.obfuscatedFont"
    rid = "rIdFont1"
    doc.parts["word/_rels/fontTable.xml.rels"] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.odttf"/></Relationships>').encode()
    doc.font_extra = (f'<w:font w:name="{family}"><w:panose1 w:val="020B0604020202020204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/>'
                      f'<w:embedRegular r:id="{rid}" w:fontKey="{{{guid}}}"/></w:font>')
    doc.settings_extra += "<w:embedTrueTypeFonts/><w:saveSubsetFonts/>"
    return guid


def themed(doc: Doc, palette=("4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47")):
    """Runs/paragraphs/tables colored through theme references, plus theme fonts, tints and shades."""
    pal = {f"accent{i + 1}": c for i, c in enumerate(palette)}
    th = lambda name, extra="": f'<w:color w:val="{pal[name]}" w:themeColor="{name}"{extra}/>'
    runs = ""
    for n in ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6"):
        runs += f'<w:r><w:rPr>{th(n)}<w:b/></w:rPr><w:t xml:space="preserve">{n} </w:t></w:r>'
        runs += f'<w:r><w:rPr>{th(n, " w:themeShade=\"BF\"")}</w:rPr><w:t xml:space="preserve">{n}-shade </w:t></w:r>'
        runs += f'<w:r><w:rPr>{th(n, " w:themeTint=\"99\"")}</w:rPr><w:t xml:space="preserve">{n}-tint </w:t></w:r>'
    out = [H("Theme colours and fonts", 1)]
    out.append(p(runs, style="BodyText"))
    out.append(p('<w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/></w:rPr><w:t xml:space="preserve">Major theme font (headings). </w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/></w:rPr><w:t>Minor theme font (body).</w:t></w:r>', style="BodyText"))
    cell = lambda t, fill, tc_: f'<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="{tc_}" w:themeFill="{fill}"/></w:tcPr>{p(r(t, color="FFFFFF"))}</w:tc>'
    out.append(f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9600" w:type="dxa"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>' + '<w:gridCol w:w="2400"/>' * 4 + '</w:tblGrid><w:tr>' +
               cell("accent1 fill", "accent1", pal["accent1"]) + cell("accent2 fill", "accent2", pal["accent2"]) + cell("accent5 fill", "accent5", pal["accent5"]) + cell("accent6 fill", "accent6", pal["accent6"]) + "</w:tr></w:tbl>")
    out.append(p(""))
    return "".join(out)
