"""Regenerate the number-format conformance table that the sheets-core engine tests check against.

    python3 -I fixtures/xlsx/generators/build_numfmt_oracle.py

Every (code, value) of numfmt_cases.py is written to a workbook with the code as a custom number format (no built-in
ids, which LibreOffice maps to its locale), LibreOffice headless converts it to CSV with "save cell content as shown",
and the display strings go into packages/sheets-core/src/number-format/number-format-oracle.json together with the
LibreOffice version.

`expected` is LibreOffice's string except where Excel is known to differ (negative dates and times, the 1900 calendar
before 1900-03-01, time rounding to the displayed second, General's 11-character width, the system date tokens). Those
rows carry `excel: true` and a `note`; the Excel string is computed by the small renderer below (self-authored, same author as the engine: not independent evidence) (dates and
times) or taken from numfmt_cases.py. None of it was checked in Excel (plan decision D8). LibreOffice pads `?` with a
figure space (U+2007); that is replaced by a plain space, which is what Excel writes.

LibreOffice and the xlsx builder are authoring tools only: the output is committed, nothing here runs in CI.
"""
import csv
import datetime
import io
import json
import os
import re
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
sys.path.insert(0, HERE)

import numfmt_cases as C  # noqa: E402
import xlsxbuilder as X  # noqa: E402
from build_styles_formulas import CODES60  # noqa: E402
from probe_data import NUMBER_FORMATS  # noqa: E402

OUT = os.path.join(ROOT, "packages", "sheets-core", "src", "number-format", "number-format-oracle.json")
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]

# Always a custom number format id: the xlsx builder maps well-known codes to built-in ids, which LibreOffice shows in its locale.
X.Styles.fmt = lambda self, code: self.fmts.setdefault(code, 164 + len(self.fmts))


def libreoffice(cases, date1904=False):
    book = X.Book(title="number format oracle")
    if date1904:
        book.wb["pr"] = '<workbookPr date1904="1"/>'
    sheet = book.sheet("Formats")
    for i, (code, value) in enumerate(cases, 1):
        sheet.set(i, 1, value, book.styles.xf(code) if code != "General" else 0)
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "oracle.xlsx")
        book.write(path)
        out = os.path.join(tmp, "out")
        subprocess.run(
            ["soffice", f"-env:UserInstallation=file://{os.path.join(tmp, 'profile')}", "--headless", "--convert-to",
             "csv:Text - txt - csv (StarCalc):44,34,76,1,,1033,false,false,true,false,false,-1", "--outdir", out, path],
            check=True, capture_output=True, timeout=600)
        with open(os.path.join(out, os.listdir(out)[0]), encoding="utf-8", newline="") as f:
            rows = list(csv.reader(io.StringIO(f.read())))
    shown = [(r[0] if r else "").replace(" ", " ") for r in rows]
    return shown + [""] * (len(cases) - len(shown))


# ---------------------------------------------------------------------------
# Independent mini renderer for date and time codes with Excel's semantics
# ---------------------------------------------------------------------------
TOKEN = re.compile(r'"[^"]*"|\\.|\[\$[^\]]*\]|\[(?:h+|m+|s+)\]|AM/PM|am/pm|A/P|a/p|y+|m+|d+|h+|s+(?:\.0+)?|[^"\\\[ymdhsAa]+|.', re.I)


def is_datetime(code):
    stripped = re.sub(r'"[^"]*"|\\.|\[\$[^\]]*\]', "", code)
    return ";" not in code and re.search(r"[ymdhs]|\[(h+|m+|s+)\]", stripped, re.I) is not None and "General" not in code


def has_date_part(code):
    """True when the code shows a year, month or day (a bare `m` or `mm` without hours or seconds is a month)."""
    stripped = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', "", code)
    return re.search(r"[yd]", stripped, re.I) is not None or (re.search("m", stripped, re.I) is not None and not re.search("[hs]", stripped, re.I))


def excel_date(serial):
    """(year, month, day, weekday 0=Sunday) in Excel's 1900 system; None for #### (negative or past 9999)."""
    day = int(serial // 1)
    if day < 0 or day > 2958465:
        return None
    weekday = (day - 1) % 7
    if day == 0:
        return 1900, 1, 0, weekday
    if day == 60:
        return 1900, 2, 29, weekday
    d = datetime.date(1899, 12, 31 if day < 60 else 30) + datetime.timedelta(days=day)
    return d.year, d.month, d.day, weekday


def excel_datetime(code, value, system=1900):
    """The Excel display of `value` for a date/time code, or None when this renderer does not cover the code."""
    if value < 0:
        return "####"
    tokens = [t for t in TOKEN.findall(code)]
    digits = max([len(t.split(".")[1]) for t in tokens if re.fullmatch(r"s+\.0+", t, re.I)] + [0])
    has_seconds = any(re.fullmatch(r"s+(\.0+)?|\[s+\]", t, re.I) for t in tokens)
    # rounded to the displayed second (or fraction), otherwise cut to the displayed field after float noise at a millisecond
    unit = 10 ** digits if has_seconds else 1000
    day, frac = int(value // 1), value - int(value // 1)
    units = int(float(f"{frac * 86400 * unit:.15g}") + 0.5)
    if units >= 86400 * unit:
        day, units = day + 1, units - 86400 * unit
    if system == 1904:
        base = datetime.date(1904, 1, 1) + datetime.timedelta(days=day)
        date = (base.year, base.month, base.day, (base.weekday() + 1) % 7)
    else:
        date = excel_date(day)
    twelve = any(re.fullmatch(r"am/pm|a/p", t, re.I) for t in tokens)
    total = day * 86400 * unit + units
    hour = units // (3600 * unit)
    kinds = [re.match(r"[ymdhsAa\[]", t) and (t[1] if t[0] == "[" else t[0]).lower() for t in tokens]
    out = ""
    for i, t in enumerate(tokens):
        k = kinds[i]
        low = t.lower()
        if t.startswith('"'):
            out += t[1:-1]
        elif t.startswith("\\"):
            out += t[1]
        elif t.startswith("[$"):
            continue
        elif t.startswith("["):
            n, per = len(t) - 2, {"h": 3600, "m": 60, "s": 1}[t[1].lower()] * unit
            out += str(total // per).zfill(n)
        elif low in ("am/pm", "a/p"):
            half = t.split("/")
            out += half[0] if hour < 12 else half[1]
        elif k == "y":
            if date is None:
                return "####"
            out += str(date[0]).zfill(4) if len(t) > 2 else str(date[0] % 100).zfill(2)
        elif k == "d":
            if date is None:
                return "####"
            out += str(date[2]).zfill(len(t)) if len(t) <= 2 else (WEEKDAYS[date[3]][:3] if len(t) == 3 else WEEKDAYS[date[3]])
        elif k == "h":
            h = (hour % 12 or 12) if twelve else hour
            out += str(h).zfill(len(t))
        elif k == "s":
            sec = (units // unit) % 60
            out += str(sec).zfill(min(len(t.split(".")[0]), 2))
            if "." in t:
                out += "." + str(units % unit).zfill(digits)[: len(t) - len(t.split(".")[0]) - 1]
        elif k == "m":
            # minutes after an hour token or before a seconds token, otherwise the month
            prev = next((kinds[j] for j in range(i - 1, -1, -1) if kinds[j] in ("y", "m", "d", "h", "s")), None)
            nxt = next((kinds[j] for j in range(i + 1, len(tokens)) if kinds[j] in ("y", "m", "d", "h", "s")), None)
            if prev == "h" or nxt == "s":
                out += str((units // (60 * unit)) % 60).zfill(len(t))
            else:
                if date is None:
                    return "####"
                n = len(t)
                out += str(date[1]).zfill(n) if n <= 2 else (MONTHS[date[1] - 1][0] if n == 5 else MONTHS[date[1] - 1][:3] if n == 3 else MONTHS[date[1] - 1])
        else:
            out += t
    return out


def excel_expected(code, value, lo, system=1900):
    """(expected, note) where Excel is known to differ from LibreOffice, else None."""
    if (code, value) in C.SYSTEM:
        return C.SYSTEM[(code, value)], "Excel expands the system long date/time token (en-US); LibreOffice leaves the number unformatted", False
    if isinstance(value, (int, float)) and value < 0 and ";" not in code and not is_datetime(code) and "[" not in code.replace("[$", "") and not re.search(r"[@]|General", code):
        digits_shown = re.sub(r"[^0-9]", "", lo)
        if digits_shown and not re.search("[1-9]", digits_shown) and not lo.startswith("-"):
            return "-" + lo, "Excel keeps the minus sign of a negative number that rounds to zero (-0, -0.00); LibreOffice drops it", False
    if (code, value) in C.POINT:
        return C.POINT[(code, value)], "Excel always writes the decimal point of a format that has one, even when no digit follows; LibreOffice drops it", False
    if code == "General" and value in C.GENERAL_EXCEL:
        return C.GENERAL_EXCEL[value], "Excel's General fits the value into 11 characters (a default-width column); LibreOffice keeps 15 significant digits", False
    if isinstance(value, (int, float)) and is_datetime(code):
        mine = excel_datetime(code, value, system)
        if mine is not None and mine != lo:
            if value < 0:
                return mine, "Excel shows #### for a negative date or time; LibreOffice shows a date before 1899-12-30 or an absolute time", False
            if mine.lower() == lo.lower():
                return mine, "Excel writes AM/PM in the case the format code uses (am/pm, A/P); LibreOffice picks its own", True
            if system == 1900 and value < 61 and has_date_part(code):
                return mine, "Excel's 1900 date system: serial 0 is 1900-01-00, serial 60 is 1900-02-29 and serials 1-59 are a day later than LibreOffice's", False
            if re.search(r"[hms]", code, re.I) and value % 1:
                return mine, "Excel is taken to round a time to the nearest displayed second (or fraction of a second) and to cut it to the displayed field when seconds are not shown; LibreOffice truncates", True
            raise SystemExit(f"{code!r} {value!r}: LibreOffice says {lo!r}, the Excel renderer says {mine!r} and no known reason applies")
    return None


def main():
    groups = []

    def add(group, pairs, values=None):
        for code, vals in pairs:
            for v in (vals if values is None else values):
                groups.append((group, code, v))

    add("probe", [(c, None) for c in NUMBER_FORMATS], C.PROBE_VALUES)
    add("corpus", [(c, None) for c in CODES60], C.CORPUS_VALUES)
    add("builtin", [(c, None) for c in C.BUILTIN_SPELLING.values()], C.CORPUS_VALUES)
    for name, pairs in (("ecma", C.ECMA), ("conditions", C.CONDITIONS), ("text", C.TEXT), ("literals", C.LITERALS), ("numbers", C.NUMBERS), ("dates", C.DATES), ("times", C.TIMES)):
        add(name, pairs)
    add("system", [(c, [v]) for (c, v) in C.SYSTEM])
    add("general", [("General", C.GENERAL_VALUES)])
    # one row per distinct (code, value): the probe codes also sit in the corpus list
    seen, rows_in = set(), []
    for g, code, v in groups:
        key = (g, code, repr(v))
        if key not in seen:
            seen.add(key)
            rows_in.append((g, code, v))
    shown = libreoffice([(c, v) for _, c, v in rows_in])
    rows = []
    for (group, code, value), lo in zip(rows_in, shown):
        row = {"group": group, "code": code, "value": value, "libreoffice": lo, "expected": lo}
        fix = excel_expected(code, value, lo)
        if fix:
            row.update(expected=fix[0], excel=True, note=fix[1])
            if fix[2]:
                row["unverified"] = True
        elif group == "conditions" and isinstance(value, (int, float)) and value < 0 and re.search(r"\[(<>|>=|>|=|<=?[^0\]])", code):
            row.update(unverified=True, note="the sign rule for an explicit condition follows LibreOffice (the minus sign is dropped in the second and later sections); Excel is not known to do the same")
        rows.append(row)
    dates1904 = [(c, v) for c, vs in C.DATE1904 for v in vs]
    for (code, value), lo in zip(dates1904, libreoffice(dates1904, date1904=True)):
        row = {"group": "date1904", "code": code, "value": value, "libreoffice": lo, "expected": lo, "dateSystem": 1904}
        fix = excel_expected(code, value, lo, 1904)
        if fix and fix[0] != lo:
            raise SystemExit(f"1904: {code!r} {value!r}: LibreOffice {lo!r}, Excel renderer {fix[0]!r}")
        rows.append(row)
    lo_version = subprocess.run(["soffice", "--version"], capture_output=True, text=True, check=True).stdout.split()[1]
    table = {"generator": "fixtures/xlsx/generators/build_numfmt_oracle.py", "libreoffice": lo_version, "builtinSpelling": C.BUILTIN_SPELLING, "rows": rows}
    with open(OUT, "w") as f:
        json.dump(table, f, indent=1, ensure_ascii=False)
        f.write("\n")
    print(f"wrote {len(rows)} rows ({sum(1 for r in rows if r.get('excel'))} Excel overrides) to {os.path.relpath(OUT, ROOT)}")


if __name__ == "__main__":
    main()
