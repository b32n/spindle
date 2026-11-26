"""Manifest data for the XLSX corpus: what each file is, what it claims to contain and how it is made.

`constructs` are claims: the Jest suite checks each one against the package XML with the probes in
scripts/xlsx-fidelity/constructs.ts, and every construct in that registry must be claimed by some file.
"""
LICENSE = "Original content authored for this repository (MIT, same as the repo)"
OS = "Debian GNU/Linux 13 (trixie), Linux x86_64 container"


def f(id, title, constructs, notes, generator, tool="programmatic", ext="xlsx", **extra):
    return {"id": id, "filename": f"{id}.{ext}", "title": title, "tool": tool, "constructs": constructs.split(), "notes": notes, "generator": f"generators/{generator}", **extra}


CATALOG = [
    f("cell-types", "Cell types", "cells.numbers cells.sharedStrings cells.inlineStrings cells.booleans cells.errors cells.dates cells.unicode cells.rtl formulas.cachedValues workbook.calcSettings",
      "Integers, decimals, extreme magnitudes, text (spaces, line break, numeric-looking), booleans, all seven error values, 1900-system dates and times (including serial 60, the non-existent 1900-02-29), an inline string, Unicode (Greek, CJK, Korean, emoji, a combining accent) and Hebrew / Arabic text, and one formula with its cached value.", "build_core.py"),
    f("dates-1904", "1904 date system", "workbook.date1904 cells.dates", "A workbook with date1904 set: the same serials as in the 1900 system denote dates 1462 days earlier.", "build_core.py"),
    f("styles-basic", "Basic styles", "styles.basic styles.alignment styles.customNumFmts styles.indexedColors styles.borderStyles layout.columnStyles layout.rowHeights layout.colWidths",
      "Bold white-on-blue header, struck italic Times, thin / thick / double borders, centred wrapped text, 45 degree rotation with indent, shrink to fit, currency / percent / scientific / red-negative number formats, a column default style and a tall row.", "build_core.py"),
    f("number-format-probe", "Number format probe (33 codes x 3 values)", "styles.customNumFmts cells.dates", "Plan 2.2 item 5: 33 common Excel codes, each applied to 1234.5678, -0.256 and 46024.75, written by xlsxwriter. Expected display strings come from recalculating in LibreOffice (X3).", "build_probes.py", tool="xlsxwriter"),
    f("formulas-core", "Formulas core: the 159 probe formulas", "formulas.basic formulas.cachedValues formulas.xlfn formulas.array formulas.dynamicArray formulas.crossSheet",
      "Plan 2.2 item 3: 159 common formulas (math, lookup, text, date, statistics, financial, information, dynamic-array, operators) over a small data block, written by xlsxwriter with future-function prefixes. Cached values are xlsxwriter's default 0; expected values come from a LibreOffice recalculation (X3).", "build_probes.py", tool="xlsxwriter"),
    f("formula-syntax", "Formula syntax probe (23 forms)", "formulas.wholeColumn formulas.wholeRow formulas.threeD formulas.crossSheet formulas.array formulas.dynamicArray formulas.xlfn cells.booleans formulas.cachedValues",
      "Plan 2.2 item 4: whole-column and whole-row references, percent literal, 3D reference, array constants, range ending in INDEX, implicit intersection, spill operator, intersection and union operators, operator precedence, quoted sheet names. Each cached value is what Excel computes.", "build_probes.py"),
    f("layout", "Layout", "layout.colWidths layout.rowHeights layout.defaultSizes layout.hidden layout.outline layout.frozenPanes layout.merges layout.zoom layout.noGridlines layout.rtl sheets.multiple cells.rtl",
      "Custom column widths and row heights, default sizes, a hidden column and row, column and row outline groups, frozen first row and column, merged cells (across and down), zoom 130%, gridlines off, and a second, right-to-left sheet.", "build_core.py"),
    f("sheet-properties", "Sheet properties", "sheets.multiple sheets.hidden sheets.veryHidden sheets.tabColor sheets.activeTab", "Five sheets: two with tab colours, one hidden, one very hidden, and the third one active.", "build_core.py"),
    f("names", "Defined names", "names.defined names.sheetScoped names.printArea names.printTitles names.constants formulas.basic sheets.multiple", "Workbook-scoped and sheet-scoped names, print area and print titles, constant names (number, string, array) and a hidden name, used by two formulas.", "build_core.py"),
    f("cf-basic", "Conditional formatting", "cf.cellIs cf.expression cf.top10 cf.aboveAverage cf.duplicates cf.text cf.dates cf.colorScale cf.dataBar cf.iconSet cf.stopIfTrue", "Every classic rule type: cell value, formula (stop if true), top / bottom, above average, duplicate and unique values, text rules, dates occurring, 3-colour scale, data bar and icon set.", "build_core.py"),
    f("data-validation", "Data validation", "dv.list dv.whole dv.decimal dv.date dv.time dv.textLength dv.custom dv.messages", "One rule per type (list inline and from a range, whole, decimal, date, time, text length, custom formula) and a rule with a prompt, an error title and a warning style.", "build_core.py"),
    f("openpyxl-image", "openpyxl: workbook with an image", "images.png formulas.basic", "Written by openpyxl with one PNG. Absolute relationship targets; the importer at the time of writing throws a TypeError on it (plan 2.2 item 1).", "build_producers.py", tool="openpyxl"),
    f("openpyxl-chart", "openpyxl: workbook with charts", "charts.bar charts.line formulas.basic", "Written by openpyxl with a bar and a line chart. The importer at the time of writing throws a TypeError on it (plan 2.2 item 1).", "build_producers.py", tool="openpyxl"),
    f("openpyxl-comment", "openpyxl: workbook with a cell comment", "comments.notes formulas.basic", "Written by openpyxl with one cell comment (legacy note). The importer at the time of writing throws a TypeError on it (plan 2.2 item 1).", "build_producers.py", tool="openpyxl"),
    f("xlsxwriter-rich", "xlsxwriter: many features", "images.png charts.bar comments.notes cf.cellIs dv.list formulas.xlfn formulas.lambda formulas.dynamicArray cells.richText names.defined layout.frozenPanes layout.merges",
      "Written by xlsxwriter (an Excel-conformant writer): a note, an image, a column chart, conditional formatting, validation, future functions, LET, a dynamic-array SORT, a defined name, frozen panes, a merge and a rich string.", "build_producers.py", tool="xlsxwriter"),
    f("libreoffice-rich", "LibreOffice: many features", "cf.cellIs cf.colorScale dv.list comments.notes charts.bar tables.listObjects sheets.hidden sheets.tabColor names.defined layout.frozenPanes layout.merges print.pageSetup",
      "An openpyxl workbook with conditional formatting, validation, a note, a table, a chart, a hidden sheet, a name, a merge, a tab colour and landscape orientation, converted by LibreOffice headless so that LibreOffice wrote the package (plan 2.2 item 2).", "build_producers.py", tool="libreoffice"),
    f("edge-truncated-zip", "Edge: truncated zip", "edge.truncatedZip", "A workbook cut at 60% of its length (the central directory is missing). The importer must report a corrupt file, not crash.", "build_edge.py", expectImportError=True),
    f("edge-no-workbook", "Edge: no workbook part", "edge.noWorkbook edge.brokenRels", "A valid zip with content types and a root relationship to xl/workbook.xml, but no workbook part.", "build_edge.py", expectImportError=True),
    f("edge-bad-rels", "Edge: relationships that point nowhere", "edge.brokenRels sheets.multiple", "Workbook relationships to a missing worksheet, a missing theme and a part outside the package, and a sheet entry whose r:id has no relationship. One real sheet remains.", "build_edge.py"),
    f("edge-xml-entities", "Edge: XML entity declarations", "edge.xmlEntities cells.inlineStrings", "A worksheet with a DOCTYPE declaring nested internal entities (1000x expansion) and an external file entity, referenced from two inline strings. Importers must not expand or resolve them.", "build_edge.py"),
    f("edge-invalid-xml", "Edge: malformed worksheet XML", "edge.invalidXml", "A worksheet part with an unclosed element and a stray closing tag.", "build_edge.py", expectImportError=True),
    f("edge-encrypted", "Edge: encrypted workbook (OLE stub)", "edge.encrypted", "An OLE compound-file header and a directory naming EncryptionInfo and EncryptedPackage streams: the shape of a password-protected xlsx, with no content (a stub, not a usable file).", "build_edge.py", expectImportError=True),
    f("edge-legacy-xls", "Edge: legacy .xls (OLE stub)", "edge.legacyBinary", "An OLE compound-file header and a directory naming a Workbook stream: the shape of a legacy .xls, with no content (a stub, not a usable file).", "build_edge.py", ext="xls", expectImportError=True),
    f("edge-duplicate-sheet-names", "Edge: duplicate sheet names", "edge.duplicateSheetNames sheets.multiple", "Three sheets named Data, Data and DATA (sheet names are case-insensitive, so all three clash).", "build_edge.py"),
    f("edge-index-out-of-range", "Edge: indexes out of range", "edge.indexOutOfRange", "A shared-string cell pointing past the end of the string table (twice) and a cell using a style index that does not exist.", "build_edge.py"),
    f("edge-empty-workbook", "Edge: empty workbook", "edge.emptyWorkbook", "A minimal package (no styles, no shared strings, no views) with one sheet that holds no cells.", "build_edge.py"),
]

RUNTIME = [
    f("edge-zip-bomb", "Edge: zip bomb (generated at run time)", "edge.zipBomb", "About 300 KB on disk; the sheet part inflates to 300 MB of XML comment (ratio about 1000:1). Never committed: python3 generators/build_runtime.py --out DIR edge-zip-bomb.", "build_runtime.py", generatedAtRuntime=True, expectImportError=True),
    f("edge-huge-dimension", "Edge: huge dimension (generated at run time)", "edge.hugeDimension", "A tiny package whose sheet declares the dimension A1:XFD1048576 (17 billion cells). Never committed: python3 generators/build_runtime.py --out DIR edge-huge-dimension.", "build_runtime.py", generatedAtRuntime=True),
    f("perf-100k-x-10", "Performance: 100,000 rows x 10 columns (generated at run time)", "perf.largeSheet cells.numbers cells.inlineStrings", "One million cells (about 5 MB). Never committed: python3 generators/build_runtime.py --out DIR perf-100k-x-10.", "build_runtime.py", generatedAtRuntime=True),
]
