"""Cases for the number-format conformance table (build_numfmt_oracle.py): which codes to render with which values.

The oracle is LibreOffice. Where Excel is known to differ (negative dates, the 1900 calendar before 1900-03-01,
time rounding, General's 11-character width, the system date tokens) build_numfmt_oracle.py replaces LibreOffice's
string by the Excel one and records why; those rows are marked `excel: true`.
"""
from probe_data import NUMBER_FORMATS
from build_styles_formulas import CODES60

PROBE_VALUES = [1234.5678, -0.256, 46024.75]
CORPUS_VALUES = [1234.5678, -0.256, 46024.75, 0]

# A format with a decimal point always shows it (LibreOffice drops a point with no digit after it).
POINT = {("0.#", 5): "5.", ("#.", 1234.5678): "1235.", ("#.", -7.5): "-8.", ("#.", 0): ".", ("0.", 1234.5678): "1235.", ("0.", -7.5): "-8.", ("0.", 0): "0.",
         ("#.##", 0): "."}

# Codes the corpus file writes as built-in ids 14 and 22; the importer maps the ids to these en-US codes (ECMA-376 18.8.30).
BUILTIN_SPELLING = {"mm-dd-yy": "m/d/yyyy", "m/d/yy h:mm": "m/d/yyyy h:mm"}

ECMA = [  # examples in the style of ECMA-376 18.8.31 and Microsoft's number-format help (modelled on them, not copied)
    ("#.000", [8.9, 0.631, -2]), ("0.#", [0.631, 1.3, 5]), ("#.0#", [1234.568, 1.5, 0.049]), ("???.???", [44.398, 102.65, 2.8]),
    ("#,##0", [0, 999.5, 1234567.891]), ("#,###", [0, 5, 1234567]), ("0.0,", [12200000, 999999]), ("0,", [1500, 499]),
    ("# ?/?", [1.75, 0.5, 3]), ("# ??/??", [3.14159, 0.333]), ("?/?", [1.5, 0.25]), ("# ?/8", [2.3, 0.5, 4]), ("# ?/16", [0.3, 1.5]), ("0/16", [0.3, 2]),
    ("# ??/100", [0.123, 2.5]), ("#,##0.00_);(#,##0.00)", [-1234.5, 0, 12]), ("0.00E+00", [12345.6789, 0.000123, 0]), ("##0.0E+0", [12345.6789, 0.000123]),
    ("0.0E+00", [-1234.5, 99999]), ("0.00E-00", [12345.6789, 0.000123]), ("0.0e+0", [1234.5, 0.00123]), ("0.00%", [0.256, -0.001, 1.5]),
    ("0%", [0.5, 0.005, 12]), ('0.0%;-0.0%;"zero"', [0.123, -0.456, 0]), ("$#,##0.00", [1234.5, -1234.5, 0]), ("00000", [42, 123456]),
    ("#,##0.0;[Red]-#,##0.0", [1234.55, -1234.55]), ("0.00;-0.00;;@", [1, -1, 0, "text"]), ("m/d/yyyy h:mm", [46024.75, 36526.5]),
    ("hh:mm AM/PM", [0.25, 0.75, 0.5]), ("d-mmm-yyyy", [46024, 36526]),
]

CONDITIONS = [  # section selection; LibreOffice decides the cases the specification leaves open
    ('[>100]"big";"small"', [150, 50, -5, 0]), ('[>100]0"big";0"small"', [150, 50, -5]), ('[>100]"a";[<0]"b";"c"', [150, 50, -5, 0]),
    ('[>100]"a";"b";"c"', [150, 50, -5, 0]), ("[<=5]0.0;0.00", [3, 5, 6, -5]), ("[<0][Red]0.0;[Blue]0.0", [-5, 5, 0]), ('[=0]"zero";0', [0, 5, -5]),
    ("[<10]0.0;[>=100]0.00;0", [6, 50, 150, -5]), ('[<=0]"n"0;0', [-5, 0, 2]), ("[Red]0;[Blue]-0", [5, -5]), ('0;[>3]0"x";0"y"', [5, 2, -2]),
    ("0;-0;;@", [5, -5, 0]), ('"p"0;"n"0;"z"0;"t"@', [5, -5, 0, "abc"]), ('0;0;0;"t"@', [5, "abc"]),
    ("0.00;(0.00)", [5, -5, 0]), ('0.00;(0.00);"-"', [5, -5, 0]), ("[Blue]0.0;[Red]-0.0;[Green]0", [1.25, -1.25, 0]),
]

TEXT = [("@", ["abc", ""]), ('@" (text)"', ["abc"]), ('0;0;0;"t"@', ["abc"]), ("0.00", ["abc"]), ('"a"@;"b"', ["abc"]), ('0;"t"@', ["abc"]),
        ('"[" @ "]"', ["x"]), ("General", ["abc"])]

LITERALS = [(c, [1234.5678, -7.5, 0]) for c in [
    '0.0 "kg"', '"Total: "0', "000-00", "(000) 000-0000", '0 "a" 0', "0\\x", "\\$0.00", "0_)", "_(0_)", "0.0\\%", '#,##0.00 "USD"', '0 "%"',
    '"x"0.00"y"', "0.00-", "($0.00)", "[$USD] #,##0.00", "[$€-407] #,##0.00", "[$-409]#,##0.00", "[$£-809]#,##0.00", '#,##0.00[$ kr-41D]',
    "0.0,,,", "0.00,,", '#,##0.0,"K"', "#.##", "#.", ".00", "0.", "###0.00", "?", "??0", "#?#",
]]

NUMBERS = [(c, [0, 1, -1, 0.5, 1234.5678, -1234.5678, 1e-7, 123456789012]) for c in
           ["0", "0.00", "#,##0.00", "0.0000", "0.00E+00", "#,##0", "0%", "0.0#"]]

# General: Excel fits the value into 11 characters (a default-width column); LibreOffice keeps 15 digits. Expected = Excel.
GENERAL_EXCEL = {
    1 / 3: "0.333333333", -1 / 3: "-0.33333333", 3.141592653589793: "3.141592654", 123456789012: "1.23457E+11", 1e-5: "1E-05", 1.5e-5: "1.5E-05",
    0.000123456789: "0.000123457", -12345678901: "-1.2346E+10", 1e20: "1E+20", 123456.7890123: "123456.789", 99999999999.9: "1E+11", 100000000000: "1E+11",
    1234567890123456: "1.23457E+15",
}
GENERAL_VALUES = list(GENERAL_EXCEL) + [0.1 + 0.2, 12345678901, 0.0001, 0, 100, -0.256, 1234.5678]

DATES = [(c, [0, 1, 59, 60, 61, 366, 36526, 45000.5, 46024.75, 2958465, -1]) for c in
         ["yyyy-mm-dd", "dddd", "mmm d yyyy", "dddd, mmmm d, yyyy", "m/d/yy", "mmmmm", "yy", "d", "m", "[$-409]d mmm yyyy", "mmm-yy"]]

TIMES = [(c, [0.5, 0.75, 1.5, 100.25, 0.123456789, 0.00001157407, 0.999994213, 0.99999999, 1.000000001, 0.0001, -0.5]) for c in
         ["h:mm:ss", "hh:mm", "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm am/pm", "h:mm A/P", "[h]:mm:ss", "[hh]:mm", "[mm]:ss", "[ss]", "h:mm:ss.000", "mm:ss.0",
          "mm:ss.00", "m:ss", "h:m", "s", 'h "h" mm "m"', "hh:mm:ss.00 AM/PM", "d h:mm"]]

# The system long date and time tokens: LibreOffice leaves the number unformatted, Excel expands them (en-US).
SYSTEM = {("[$-F800]", 46024.75): "Friday, January 2, 2026", ("[$-F800]", 1): "Sunday, January 1, 1900", ("[$-F400]", 46024.75): "6:00:00 PM",
          ("[$-F400]", 0.5): "12:00:00 PM", ("[$-F800]dddd, mmmm dd, yyyy", 46024.75): "Friday, January 02, 2026", ("[$-F400]h:mm:ss AM/PM", 0.75): "6:00:00 PM"}

DATE1904 = [(c, [0, 1, 1462, 44562, 46024.75]) for c in ["yyyy-mm-dd", "dddd", "mmm d, yyyy", "yyyy-mm-dd hh:mm"]]
