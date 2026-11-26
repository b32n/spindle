"""Theme and number-format styles, modern formulas and references."""
import xlsxbuilder as X
from build_core import sample

BORDER_STYLES = ["thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"]
PATTERNS = ["darkGray", "mediumGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightGrid"]
CODES60 = [
    "General", "0", "0.00", "#,##0", "#,##0.00", "0%", "0.00%", "0.00E+00", "##0.0E+0", "# ?/?", "# ??/??", "# ???/???", "mm-dd-yy", "d-mmm-yy", "d-mmm", "mmm-yy",
    "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss", "m/d/yy h:mm", "yyyy-mm-dd", "dd/mm/yyyy", "mmm d, yyyy", "dddd", "dddd, mmmm d, yyyy", "mmmm yyyy", "yyyy-mm-dd hh:mm:ss",
    "[h]:mm:ss", "[mm]:ss", "mm:ss.0", "$#,##0", '"$"#,##0.00', "$#,##0.00_);[Red]($#,##0.00)", "#,##0_);(#,##0)", "#,##0.00_);(#,##0.00)", "#,##0;[Red]-#,##0",
    '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', "0.0,,\"M\"", '#,##0.0,"K"', '0.0 "kg"', '"Total: "0', "000-00",
    "00000", "(000) 000-0000", "@", '@" (text)"', "0;-0;;@", '[>100]"big";"small"', '[<0][Red]0.0;[Blue]0.0', '[=0]"zero";0', "[$€-407] #,##0.00", "[$¥-411] #,##0", "[$-409]mmmm d, yyyy",
    "[$-F800]dddd, mmmm dd, yyyy", "0.00%;[Red]-0.00%", "#,##0.000", "* 0", "0_)",
]


def styles_theme(path):
    b = X.Book(title="Theme and fills")
    b.add("xl/theme/theme1.xml", X.theme_xml(), "theme")
    b.wbrels.add("theme", "theme/theme1.xml")
    st = b.styles
    st.colors = '<colors><indexedColors>' + "".join(f'<rgbColor rgb="FF{c}"/>' for c in ["000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF"] * 8) + "</indexedColors></colors>"
    st.cell_style_xfs.append('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>')
    st.cell_styles += ['<cellStyle name="Good" xfId="1" builtinId="26"/>', '<cellStyle name="Spindle Emphasis" xfId="1"/>']
    sh = b.sheet("Theme")
    sh.set(1, 1, "theme colour + tint", st.xf(font='<font><sz val="11"/><color theme="4" tint="-0.249977111117893"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>',
                                               fill='<fill><patternFill patternType="solid"><fgColor theme="5" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill>'))
    sh.set(2, 1, "indexed colour", st.xf(font='<font><sz val="11"/><color indexed="10"/><name val="Calibri"/></font>', fill='<fill><patternFill patternType="solid"><fgColor indexed="13"/><bgColor indexed="64"/></patternFill></fill>'))
    sh.set(3, 1, "named style (Good)", st.xf(fill=st.solid("C6EFCE"), xf_id=1))
    sh.set(4, 1, "gradient fill", st.xf(fill='<fill><gradientFill degree="90"><stop position="0"><color rgb="FFFFFFFF"/></stop><stop position="1"><color theme="4"/></stop></gradientFill></fill>'))
    for i, p in enumerate(PATTERNS):
        sh.set(5 + i, 1, p, st.xf(fill=f'<fill><patternFill patternType="{p}"><fgColor rgb="FF1F4E78"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>'))
    for i, p in enumerate(BORDER_STYLES):
        sh.set(5 + i, 3, p, st.xf(border=f'<border><left style="{p}"><color auto="1"/></left><right style="{p}"><color auto="1"/></right><top style="{p}"><color auto="1"/></top><bottom style="{p}"><color auto="1"/></bottom><diagonal/></border>'))
    sh.set(19, 1, "diagonal borders", st.xf(border='<border diagonalUp="1" diagonalDown="1"><left/><right/><top/><bottom/><diagonal style="thin"><color auto="1"/></diagonal></border>'))
    sh.x["cols"] = '<cols><col min="1" max="1" width="26" customWidth="1"/><col min="3" max="3" width="20" customWidth="1"/></cols>'
    b.write(path)


def number_formats(path):
    b = X.Book(title="60 number formats")
    st = b.styles
    sh = b.sheet("Formats")
    sh.row(1, ["Code", "Positive", "Negative", "Date-time", "Zero"])
    for i, code in enumerate(CODES60, 2):
        s = st.xf(code)
        sh.set(i, 1, ("i", code)).set(i, 2, 1234.5678, s).set(i, 3, -0.256, s).set(i, 4, 46024.75, s).set(i, 5, 0, s)
    sh.x["cols"] = '<cols><col min="1" max="1" width="52" customWidth="1"/><col min="2" max="5" width="24" customWidth="1"/></cols>'
    b.write(path)


def formulas_modern(path):
    b = X.Book(title="Modern formulas")
    X.dynamic_array_metadata(b)
    sh = sample(b.sheet("Modern"))
    sh.set(1, 6, "Formula").set(1, 7, "Result")
    rows = [
        ("IFS", "_xlfn.IFS(B2>5,\"big\",TRUE,\"small\")", "big"), ("XLOOKUP", "_xlfn.XLOOKUP(\"East\",A2:A6,B2:B6)", 7), ("CONCAT", "_xlfn.CONCAT(A2,\"-\",A3)", "North-South"),
        ("TEXTJOIN", "_xlfn.TEXTJOIN(\",\",TRUE,A2:A4)", "North,South,East"), ("STDEV.S", "_xlfn.STDEV.S(B2:B6)", 3.361547262794322), ("LET", "_xlfn.LET(_xlpm.x,B2,_xlpm.y,B3,_xlpm.x+_xlpm.y)", 15),
        ("LAMBDA", "_xlfn.LAMBDA(_xlpm.n,_xlpm.n*2)(B2)", 20), ("MAXIFS", "_xlfn.MAXIFS(B2:B6,A2:A6,\"?o*\")", 10),
    ]
    for i, (label, f, v) in enumerate(rows, 2):
        sh.set(i, 6, label).set(i, 7, ("f", f, v, ""))
    sh.set(11, 6, "SORT (dynamic array)").set(11, 7, ("f", "_xlfn._xlws.SORT(B2:B6)", 1, ' t="array" ref="G11:G15"'), cm="1")
    for i, v in enumerate([4, 5, 7, 10], 12):
        sh.set(i, 7, v)
    sh.set(17, 6, "legacy array (CSE)").set(17, 7, ("f", "SUM(B2:B6*C2:C6)", 357, ' t="array" ref="G17"'))
    sh.set(19, 6, "shared formula")
    sh.set(19, 7, ("f", "B2*2", 20, ' t="shared" ref="G19:G23" si="0"'))
    for i, v in enumerate([10, 14, 2, 8], 20):
        sh.set(i, 7, ("f", "", v, ' t="shared" si="0"'))
    b.write(path)


def references(path):
    b = X.Book(title="References")
    sample(b.sheet("Data"))
    sample(b.sheet("Other"), rows=[("a", 100, 200, 300)] * 3)
    b.sheet("Third").set(1, 1, 7)
    sh = b.sheets[0]
    sh.set(8, 1, "Whole column").set(8, 2, ("f", "SUM(B:B)", 27, ""))
    sh.set(9, 1, "Whole row").set(9, 2, ("f", "SUM(2:2)", 60, ""))
    sh.set(10, 1, "Cross-sheet").set(10, 2, ("f", "Other!B2*2", 200, ""))
    sh.set(11, 1, "3D").set(11, 2, ("f", "SUM(Data:Third!B2)", 110, ""))
    sh.set(12, 1, "Circular (self)").set(12, 2, ("f", "B12+1", 0, ""))
    sh.set(13, 1, "External workbook").set(13, 2, ("f", "[1]Remote!$A$1", 42, ""))
    b.add("xl/externalLinks/externalLink1.xml", f'{X.HEAD}<externalLink xmlns="{X.M}" xmlns:r="{X.R}"><externalBook r:id="rId1"><sheetNames><sheetName val="Remote"/></sheetNames>'
          '<sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>42</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>', X.SS + "externalLink+xml")
    rels = X.Rels()
    rels.add("externalLinkPath", "other-workbook.xlsx", external=True)
    b.add("xl/externalLinks/_rels/externalLink1.xml.rels", rels.xml())
    b.wb["externalReferences"] = f'<externalReferences><externalReference r:id="{b.wbrels.add("externalLink", "externalLinks/externalLink1.xml")}"/></externalReferences>'
    b.write(path)
