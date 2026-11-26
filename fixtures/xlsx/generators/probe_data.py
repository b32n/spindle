"""Data for the probe workbooks (plan 2.2 items 3 and 5): 159 common formulas and 33 number-format codes."""

FORMULAS = [
    'SUM(A1:A10)', 'AVERAGE(A1:A10)', 'COUNT(A1:C10)', 'MAX(A1:A10)', 'MIN(A1:A10)', 'ROUND(A1/3,2)', 'ABS(-3)', 'MOD(A2,3)',
    'POWER(A2,3)', 'SQRT(A3)', 'INT(A2/3)', 'CEILING(A2,2)', 'FLOOR(A2,2)', 'SIN(1)', 'COS(1)', 'PI()', 'RADIANS(180)',
    'FACT(5)', 'COMBIN(5,2)', 'SUMSQ(A1:A3)', 'EVEN(3)', 'ODD(2)', 'ISEVEN(4)', 'MAXIFS(A1:A10,B1:B10,"x")', 'AGGREGATE(9,6,A1:A10)',
    'VLOOKUP(3,A1:C10,2,FALSE)', 'HLOOKUP(1,A1:C3,2,FALSE)', 'INDEX(A1:C10,2,2)', 'MATCH(3,A1:A10,0)', 'CHOOSE(2,"a","b")',
    'IF(A1>1,"y","n")', 'IFS(A1>5,"a",TRUE,"b")', 'SWITCH(A1,1,"one","other")', 'IFERROR(1/0,"e")', 'AND(TRUE,FALSE)',
    'XOR(TRUE,FALSE)', 'SUMIF(B1:B10,"x",A1:A10)', 'SUMIFS(A1:A10,B1:B10,"x",A1:A10,">1")', 'COUNTIF(B1:B10,"x")', 'COUNTIFS(B1:B10,"x",A1:A10,">1")',
    'AVERAGEIF(B1:B10,"x",A1:A10)', 'SUMPRODUCT(A1:A3,A1:A3)', 'SUMPRODUCT((A1:A10>3)*1)', 'LEFT(B1,1)', 'RIGHT(B1,1)',
    'MID("abcdef",2,3)', 'LEN("abc")', 'UPPER("a")', 'LOWER("A")', 'PROPER("ab cd")', 'TRIM("  a  b ")', 'CONCAT("a","b")',
    'TEXTJOIN(",",TRUE,B1:B3)', 'SUBSTITUTE("aab","a","c")', 'FIND("b","abc")', 'SEARCH("B","abc")', 'REPT("a",3)', 'TEXT(1234.5,"#,##0.00")',
    'TEXT(DATE(2026,1,2),"yyyy-mm-dd")', 'TEXT(DATE(2026,1,2),"dddd")', 'TEXT(0.256,"0.0%")', 'VALUE("12")', 'CLEAN("a")',
    'EXACT("a","a")', 'NUMBERVALUE("1,5",",")', 'TEXTBEFORE("a-b","-")', 'TEXTSPLIT("a-b","-")', 'DATE(2026,2,3)', 'YEAR(DATE(2026,2,3))',
    'MONTH(DATE(2026,2,3))', 'DAY(DATE(2026,2,3))', 'WEEKDAY(DATE(2026,2,3))', 'EDATE(DATE(2026,1,31),1)', 'EOMONTH(DATE(2026,1,31),1)',
    'DATEDIF(DATE(2025,1,1),DATE(2026,3,1),"m")', 'NETWORKDAYS(DATE(2026,1,1),DATE(2026,1,31))', 'WORKDAY(DATE(2026,1,1),5)',
    'DAYS(DATE(2026,2,1),DATE(2026,1,1))', 'YEARFRAC(DATE(2026,1,1),DATE(2026,7,1))', 'ISOWEEKNUM(DATE(2026,1,5))', 'WEEKNUM(DATE(2026,1,5))',
    'DATEVALUE("2026-01-02")', 'MEDIAN(A1:A10)', 'STDEV(A1:A10)', 'STDEV.S(A1:A10)', 'VAR.S(A1:A10)', 'PERCENTILE(A1:A10,0.9)',
    'PERCENTILE.INC(A1:A10,0.9)', 'QUARTILE(A1:A10,1)', 'RANK(3,A1:A10)', 'RANK.EQ(3,A1:A10)', 'LARGE(A1:A10,2)', 'SMALL(A1:A10,2)',
    'MODE(A1:A10)', 'CORREL(A1:A10,A1:A10)', 'SLOPE(A1:A10,A1:A10)', 'FORECAST(2,A1:A10,A1:A10)', 'GEOMEAN(A1:A3)', 'PMT(0.05/12,60,-10000)',
    'FV(0.05,10,-100)', 'PV(0.05,10,-100)', 'NPV(0.1,A1:A3)', 'IRR(D1:D4)', 'XNPV(0.1,D1:D4,E1:E4)', 'XIRR(D1:D4,E1:E4)',
    'SLN(1000,100,5)', 'NPER(0.05,-100,1000)', 'RATE(10,-100,800)', 'ISBLANK(Z1)', 'ISNUMBER(A1)', 'ISTEXT(B1)', 'ISERROR(1/0)',
    'ISNA(NA())', 'TYPE(A1)', 'N(5)', 'ROW(A5)', 'COLUMN(C1)', 'ROWS(A1:A10)', 'COLUMNS(A1:C1)', 'ADDRESS(2,3)', 'OFFSET(A1,2,1)',
    'INDIRECT("A2")', 'XLOOKUP(3,A1:A10,A1:A10)', 'XMATCH(3,A1:A10)', 'LOOKUP(3,A1:A10)', 'TRANSPOSE(A1:B2)', 'FILTER(A1:A10,A1:A10>8)',
    'SORT(A1:A3,1,-1)', 'UNIQUE(B1:B10)', 'SEQUENCE(3)', 'SORTBY(A1:A3,A1:A3,-1)', 'LET(x,2,x*3)', 'HSTACK(A1:A2,B1:B2)',
    'VSTACK(A1:A2,B1:B2)', 'TAKE(A1:A10,2)', 'COUNTA(A1:C10)', 'COUNTBLANK(A1:C10)', 'SUBTOTAL(9,A1:A10)', 'PRODUCT(A1:A3)',
    'MROUND(7,5)', 'TRUNC(2.7)', 'SIGN(-2)', 'LOG(100,10)', 'LN(1)', 'EXP(1)', 'GCD(12,18)', 'LCM(4,6)', 'ROUNDUP(2.1,0)',
    'ROUNDDOWN(2.9,0)', 'A1^2', '-2^2', 'A1:A3*2', '"a"&1', '(1+2)*3', '1=1', '"a"<"b"', '50%', 'SUM(A1:A3,10)', 'SUM(Data!A1:A3)',
]

NUMBER_FORMATS = [
    '0', '0.00', '#,##0', '#,##0.00', '$#,##0.00', '"$"#,##0.00', '0%', '0.00%', '0.00E+00', '# ?/?', '# ??/??', 'yyyy-mm-dd',
    'm/d/yyyy', 'd-mmm-yy', 'dd/mm/yyyy', 'mmm d, yyyy', 'dddd', 'h:mm AM/PM', 'hh:mm:ss', 'yyyy-mm-dd hh:mm', '[h]:mm:ss',
    '#,##0;[Red]-#,##0', '#,##0;(#,##0)', '0.0,,"M"', '#,##0.00_);(#,##0.00)', '[>100]"big";"small"', '000-00', '@',
    '0.0 "kg"', '#,##0.0,"K"', '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)', '[$€-407] #,##0.00', 'mmmm yyyy',
]

# Functions xlsxwriter must be told to write as dynamic arrays (they spill).
DYNAMIC = {"FILTER", "SORT", "UNIQUE", "SEQUENCE", "SORTBY", "LET", "HSTACK", "VSTACK", "TAKE", "TRANSPOSE", "TEXTSPLIT"}

# Formula-syntax probe (plan 2.2 item 4): (formula as stored in the file, cached value Excel computes, note).
# The first form must stay in row 2 (implicit intersection reads A2).
# Data: A1:A5 = 1..5, B1 = 10, C1 = 20, Sheet2!A1 = 100, 'Q1 Data'!A1 = 7; formulas sit in column G (rows 2..), none in row 1.
SYNTAX = [
    ("_xlfn.SINGLE(A1:A3)", 2, "implicit intersection (@): the formula is in row 2"),
    ("SUM(A:A)", 15, "whole column"), ("SUM(1:1)", 32, "whole row (includes J1)"), ("A1+50%", 1.5, "percent literal"),
    ("SUM(Sheet1:Sheet2!A1)", 101, "3D reference"), ("SUM({1,2,3})", 6, "array constant"), ("SUM({1,2;3,4})", 10, "2D array constant"),
    ("SUM(A1:INDEX(A1:A5,3))", 6, "range ending in INDEX"), ("ROW(1:1)", 1, "row range argument"),
    ("SUM(_xlfn.ANCHORARRAY(J1))", 6, "spill operator (J1#)"), ("SUMPRODUCT(A1:A3*2)", 12, "array arithmetic"), ("-2^2", 4, "unary minus binds first"),
    ("2^3^2", 64, "left-associative power"), ('"a"&1', "a1", "concatenation"), ("1=1", True, "comparison"), ('"a"<"b"', True, "text comparison"),
    ("SUM(A1:A3 A2:A4)", 5, "range intersection"), ("SUM(A1:A3,A5)", 11, "argument list"), ("SUM((A1:A2,A4))", 7, "union operator"),
    ("'Q1 Data'!A1", 7, "quoted sheet name"), ("$A$1+A$1+$A1", 3, "mixed anchors"), ("SUM(A1:A5)/COUNT(A1:A5)", 3, "division of aggregates"),
    ('IF(A1=1,"y","n")', "y", "text result"),
]
