"""Core constructs: cell types, dates, basic styles, layout, sheet properties, names, conditional formatting, validation."""
import xlsxbuilder as X

DATA = [("North", 10, 20, 30), ("South", 5, 15, 25), ("East", 7, 8, 9), ("West", 1, 2, 3), ("Centre", 4, 6, 8)]


def sample(sh, rows=DATA, header=("Region", "Q1", "Q2", "Q3"), s=0):
    sh.row(1, header, s=s)
    for i, r in enumerate(rows, 2):
        sh.row(i, r)
    return sh


def cell_types(path):
    b = X.Book(title="Cell types")
    st = b.styles
    date, tm, dt = st.xf("yyyy-mm-dd"), st.xf("h:mm:ss"), st.xf("yyyy-mm-dd hh:mm")
    sh = b.sheet("Types")
    sh.row(1, ["Kind", "Value"])
    sh.row(2, ["integer", 42]).row(3, ["negative", -17]).row(4, ["decimal", 3.14159265358979]).row(5, ["large", 1.5e+300]).row(6, ["tiny", 2.5e-300])
    sh.row(7, ["text", "Hello, world"]).row(8, ["spaces", "  padded  "]).row(9, ["multi-line", "line one\nline two"]).row(10, ["numeric text", "00123"])
    sh.row(11, ["true", True]).row(12, ["false", False])
    for i, e in enumerate(["#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!"], 13):
        sh.set(i, 1, e).set(i, 2, ("e", e))
    sh.set(20, 1, "date 1900").set(20, 2, 45292, date).set(21, 1, "date-time").set(21, 2, 46024.75, dt).set(22, 1, "time").set(22, 2, 0.5, tm)
    sh.set(23, 1, "serial 60 (1900-02-29, never existed)").set(23, 2, 60, date).set(24, 1, "serial 1").set(24, 2, 1, date)
    sh.set(25, 1, "inline string").set(25, 2, ("i", "stored inline"))
    sh.set(26, 1, "unicode").set(26, 2, "Grüße, Ελληνικά, 日本語, 한국어, emoji \U0001F600 and é́")
    sh.set(27, 1, "RTL Hebrew").set(27, 2, "שלום עולם").set(28, 1, "RTL Arabic").set(28, 2, "مرحبا بالعالم")
    sh.set(29, 1, "formula with cache").set(29, 2, ("f", "B2+B3", 25, ""))
    sh.x["cols"] = '<cols><col min="1" max="1" width="36" customWidth="1"/><col min="2" max="2" width="40" customWidth="1"/></cols>'
    b.write(path)


def dates_1904(path):
    b = X.Book(title="1904 date system")
    b.wb["pr"] = '<workbookPr date1904="1"/>'
    d = b.styles.xf("yyyy-mm-dd")
    sh = b.sheet("Dates1904")
    sh.row(1, ["Serial", "Date (1904 system)"])
    for i, v in enumerate([0, 1, 366, 42004, 44562], 2):
        sh.set(i, 1, v).set(i, 2, v, d)
    sh.x["cols"] = '<cols><col min="1" max="2" width="20" customWidth="1"/></cols>'
    b.write(path)


def styles_basic(path):
    b = X.Book(title="Basic styles")
    st = b.styles
    bold = st.xf(font='<font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Arial"/><family val="2"/></font>', fill=st.solid("1F4E78"))
    ital = st.xf(font='<font><i/><u/><strike/><sz val="10"/><color indexed="10"/><name val="Times New Roman"/><family val="1"/></font>')
    box = st.xf(border='<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thick"><color rgb="FFFF0000"/></top><bottom style="double"><color auto="1"/></bottom><diagonal/></border>')
    centre = st.xf(align='horizontal="center" vertical="center" wrapText="1"')
    rot = st.xf(align='horizontal="left" textRotation="45" indent="1"')
    shrink = st.xf(align='shrinkToFit="1"')
    money, pct, sci = st.xf('"$"#,##0.00'), st.xf("0.0%"), st.xf("0.00E+00")
    neg = st.xf("#,##0;[Red]-#,##0")
    colstyle = st.xf(fill=st.solid("FFF2CC"))
    sh = b.sheet("Styles")
    sh.set(1, 1, "Header", bold).set(1, 2, "struck italic", ital).set(1, 3, "boxed", box).set(1, 4, "centred wrapped text that is long", centre).set(1, 5, "rotated", rot).set(1, 6, "shrink to fit this text", shrink)
    sh.set(2, 1, 1234.5, money).set(2, 2, 0.256, pct).set(2, 3, 12345.678, sci).set(2, 4, -1234, neg)
    sh.x["cols"] = f'<cols><col min="1" max="1" width="18" customWidth="1"/><col min="7" max="7" width="12" style="{colstyle}" customWidth="1"/></cols>'
    sh.rowattrs[1] = 'ht="45" customHeight="1"'
    b.write(path)


def layout(path):
    b = X.Book(title="Layout")
    sh = sample(b.sheet("Layout"), header=("Region", "Q1", "Q2", "Q3"))
    sh.x["sheetViews"] = '<sheetViews><sheetView tabSelected="1" zoomScale="130" zoomScaleNormal="130" showGridLines="0" workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="B2" sqref="B2"/></sheetView></sheetViews>'
    sh.x["sheetFormatPr"] = '<sheetFormatPr baseColWidth="10" defaultColWidth="14.5" defaultRowHeight="18" customHeight="1" outlineLevelRow="1" outlineLevelCol="1"/>'
    sh.x["cols"] = ('<cols><col min="1" max="1" width="24" customWidth="1"/><col min="2" max="2" width="9.5" customWidth="1"/><col min="3" max="3" width="14.5" hidden="1" customWidth="1"/>'
                    '<col min="4" max="5" width="14.5" outlineLevel="1" customWidth="1"/><col min="6" max="6" width="14.5" collapsed="1"/></cols>')
    sh.rowattrs.update({3: 'ht="40" customHeight="1"', 4: 'hidden="1"', 5: 'outlineLevel="1"', 6: 'outlineLevel="1"'})
    sh.set(8, 1, "merged across").set(10, 1, "merged down")
    sh.x["mergeCells"] = '<mergeCells count="2"><mergeCell ref="A8:C8"/><mergeCell ref="A10:A12"/></mergeCells>'
    rtl = b.sheet("RightToLeft")
    rtl.row(1, ["שלום", "عالم", "text"])
    rtl.x["sheetViews"] = '<sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews>'
    b.write(path)


def sheet_properties(path):
    b = X.Book(title="Sheet properties")
    b.wb["bookViews"] = '<bookViews><workbookView activeTab="2"/></bookViews>'
    for n, state, color in [("First", None, "FFFF8800"), ("Second", None, "FF00B050"), ("Third (active)", None, None), ("Hidden", "hidden", None), ("Very hidden", "veryHidden", None)]:
        sh = b.sheet(n, state)
        sh.set(1, 1, n)
        if color:
            sh.x["sheetPr"] = f'<sheetPr><tabColor rgb="{color}"/></sheetPr>'
        if n.startswith("Third"):
            sh.x["sheetViews"] = '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>'
    b.write(path)


def names(path):
    b = X.Book(title="Defined names")
    sample(b.sheet("Data"))
    sample(b.sheet("Other"))
    b.names = [
        '<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$D$6</definedName>',
        '<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName>',
        '<definedName name="Quarter1">Data!$B$2:$B$6</definedName>',
        '<definedName name="Local" localSheetId="1">Other!$C$2:$C$6</definedName>',
        '<definedName name="TaxRate">0.2</definedName>',
        '<definedName name="Greeting">"Hello"</definedName>',
        '<definedName name="Lookup">{1,2,3}</definedName>',
        '<definedName name="HiddenName" hidden="1">Data!$A$1</definedName>',
    ]
    sh = b.sheets[0]
    sh.set(8, 1, "Total Q1").set(8, 2, ("f", "SUM(Quarter1)", 27, "")).set(9, 1, "With tax").set(9, 2, ("f", "B8*(1+TaxRate)", 32.4, ""))
    b.write(path)


def cf_basic(path):
    b = X.Book(title="Conditional formatting")
    st = b.styles
    red = st.dxf('<font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill>')
    green = st.dxf('<font><b/><color rgb="FF006100"/></font><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill>')
    sh = b.sheet("CF")
    sh.row(1, ["Value", "Text", "Date", "Scale", "Bar", "Icons"])
    for i in range(2, 12):
        sh.set(i, 1, i * 3 % 11).set(i, 2, ["apple", "banana", "cherry", "apple", "", "date", "apple", "fig", "grape", "kiwi"][i - 2]).set(i, 3, 46000 + i * 9).set(i, 4, i * 7).set(i, 5, i * 5).set(i, 6, i * 4)
    sh.set(11, 2, ("e", "#N/A"))
    rules = [
        ("A2:A11", f'<cfRule type="cellIs" dxfId="{red}" priority="1" operator="greaterThan"><formula>7</formula></cfRule>'),
        ("A2:A11", f'<cfRule type="expression" dxfId="{green}" priority="2" stopIfTrue="1"><formula>MOD(A2,2)=0</formula></cfRule>'),
        ("A2:A11", f'<cfRule type="top10" dxfId="{red}" priority="3" rank="3"/><cfRule type="top10" dxfId="{green}" priority="4" percent="1" bottom="1" rank="20"/>'),
        ("A2:A11", f'<cfRule type="aboveAverage" dxfId="{green}" priority="5"/>'),
        ("A2:A11", f'<cfRule type="duplicateValues" dxfId="{red}" priority="6"/><cfRule type="uniqueValues" dxfId="{green}" priority="7"/>'),
        ("B2:B11", f'<cfRule type="containsText" dxfId="{red}" priority="8" operator="containsText" text="an"><formula>NOT(ISERROR(SEARCH("an",B2)))</formula></cfRule>'
                   f'<cfRule type="beginsWith" dxfId="{green}" priority="9" operator="beginsWith" text="a"><formula>LEFT(B2,1)="a"</formula></cfRule>'
                   f'<cfRule type="containsBlanks" dxfId="{red}" priority="10"><formula>LEN(TRIM(B2))=0</formula></cfRule>'
                   f'<cfRule type="containsErrors" dxfId="{red}" priority="11"><formula>ISERROR(B2)</formula></cfRule>'),
        ("C2:C11", f'<cfRule type="timePeriod" dxfId="{green}" priority="12" timePeriod="lastMonth"><formula>AND(MONTH(C2)=MONTH(EDATE(TODAY(),-1)),YEAR(C2)=YEAR(EDATE(TODAY(),-1)))</formula></cfRule>'),
        ("D2:D11", '<cfRule type="colorScale" priority="13"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule>'),
        ("E2:E11", '<cfRule type="dataBar" priority="14"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>'),
        ("F2:F11", '<cfRule type="iconSet" priority="15"><iconSet iconSet="3TrafficLights1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>'),
    ]
    sh.x["conditionalFormatting"] = "".join(f'<conditionalFormatting sqref="{r}">{x}</conditionalFormatting>' for r, x in rules)
    b.write(path)


def data_validation(path):
    b = X.Book(title="Data validation")
    sh = b.sheet("Validation")
    sh.row(1, ["List", "Whole", "Decimal", "Date", "Time", "Text length", "Custom", "Range list", "Messages"])
    sh.set(10, 1, "Yes").set(10, 2, "No")
    dvs = [
        ('type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A2:A5"', '<formula1>"Yes,No,Maybe"</formula1>'),
        ('type="whole" operator="between" allowBlank="1" showErrorMessage="1" sqref="B2:B5"', "<formula1>1</formula1><formula2>10</formula2>"),
        ('type="decimal" operator="greaterThanOrEqual" showErrorMessage="1" sqref="C2:C5"', "<formula1>0.5</formula1>"),
        ('type="date" operator="greaterThan" showErrorMessage="1" sqref="D2:D5"', "<formula1>45292</formula1>"),
        ('type="time" operator="lessThan" showErrorMessage="1" sqref="E2:E5"', "<formula1>0.5</formula1>"),
        ('type="textLength" operator="lessThanOrEqual" showErrorMessage="1" sqref="F2:F5"', "<formula1>8</formula1>"),
        ('type="custom" showErrorMessage="1" sqref="G2:G5"', "<formula1>ISNUMBER(G2)</formula1>"),
        ('type="list" showErrorMessage="1" sqref="H2:H5"', "<formula1>$A$10:$B$10</formula1>"),
        ('type="whole" operator="equal" errorStyle="warning" showInputMessage="1" showErrorMessage="1" errorTitle="Heads up" error="Enter 7." promptTitle="Seven" prompt="Type the number 7." sqref="I2:I5"', "<formula1>7</formula1>"),
    ]
    sh.x["dataValidations"] = f'<dataValidations count="{len(dvs)}">' + "".join(f"<dataValidation {a}>{f}</dataValidation>" for a, f in dvs) + "</dataValidations>"
    b.write(path)
