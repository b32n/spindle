# Number formats

`packages/sheets-core/src/number-format/` renders a cell value from an Excel number-format code (ECMA-376 Part 1,
18.8.31). It is used by the canvas (`formatNumber`), by `TEXT()` and by the structured format types (date, time,
datetime, scientific, fraction, duration, custom). It has no dependencies and never throws.

```ts
import { formatValue, formatDisplay, compileFormat, isValidFormatCode } from '@b32nio/spindle-sheets-core';

formatDisplay(46024.75, 'dddd, mmmm d, yyyy');        // "Friday, January 2, 2026"
formatValue(-1234.5, '#,##0.00;[Red](#,##0.00)');     // { text: "(1,234.50)", color: "Red" }
formatValue(5, '_($* #,##0.00_)');                    // { text: " $5.00 ", fill: { index: 2, char: " " } }
formatValue(1, '0.0 abc');                            // { text: "1", invalid: true }  (shown as General)
formatDisplay(0, 'yyyy-mm-dd', { dateSystem: 1904 }); // "1904-01-01"
```

Codes are compiled once and cached (500 entries). `compileFormat(code)` returns `{ ok, error?, sections }`.

## Supported grammar

| Feature | Notes |
|---|---|
| Sections | up to four, `positive;negative;zero;text`; one section applies to all numbers (a negative gets `-`), two = positive and zero; negative, three adds zero; the negative section shows the magnitude, its literal text carries the sign |
| Conditions | `[>100]`, `[<=5]`, `[=0]`, `[<>3]`: the first section whose condition holds; a missing condition in the first section always holds, in the second it means "negative" when there is a third section and "anything else" otherwise; a lone condition that is not met shows General |
| Text section | the 4th section when there are four, otherwise the last section if it contains `@`; any other format leaves text unchanged |
| Colours | `[Red]`, `[Blue]`, `[Green]`, `[Cyan]`, `[Magenta]`, `[Yellow]`, `[White]`, `[Black]`, `[Color1]` to `[Color56]`; recorded in `result.color`, the canvas ignores it |
| Locale and currency | `[$€-407]`, `[$USD]`, `[$-409]`: the symbol is inserted, the language id is recorded on the section (`section.locale`) and does not change month or day names (en-US only) |
| Digits | `0` `#` `?`, `.`, thousands separators, scaling commas (each trailing comma divides by 1000), `%` (each one multiplies by 100) |
| Exponent | `E+00`, `E-00`, `e+0`, engineering codes such as `##0.0E+0` |
| Fractions | `# ?/?`, `# ??/??`, `?/?`, fixed denominators `# ?/8`, `0/16`, `# ??/100`; the closest fraction with the smallest denominator |
| Literals | `"quoted"`, `\x`, and unquoted `$ - + / ( ) : ! ^ & ' ~ { } < > = space`, digits 1-9 and any non-ASCII character |
| Padding | `_x` renders a space; `*x` (repeat x to fill the cell) renders nothing and is recorded in `result.fill` (`index` into `text`, `char`), like LibreOffice, so the text is the same on every device |
| `@` | text placeholder |
| `General` | Excel's General in an 11-character column: whole numbers up to 11 digits, otherwise as many decimals as fit, `E+nn` from 1e11 and below 1e-4, trailing zeros dropped |
| Dates and times | `yy yyyy m mm mmm mmmm mmmmm d dd ddd dddd h hh s ss AM/PM am/pm A/P a/p`, `[h] [mm] [ss]` elapsed, fractional seconds `.0 .00 .000`, `m` is minutes after an hour token or before a seconds token (otherwise a month), `[$-F800]` and `[$-F400]` alone are the en-US long date and long time |
| Date systems | 1900 (with Excel's 29 February 1900 = serial 60 and serial 0 = "1900-01-00") and 1904, `options.dateSystem` |
| Values | numbers; text (text section); booleans show `TRUE`/`FALSE`; error strings (`#N/A`...) pass through; `null` is empty; `NaN` and infinities are `#NUM!` |

Rounding is half away from zero on the 15 significant digits of the number (so 1.005 with `0.00` is `1.01`).
A negative number that rounds to zero keeps its sign, as in Excel: `-0.4` with `0` is `-0`, `-0.001` with `0.00` is `-0.00`
(a negative section such as `0.00;(0.00)` shows `(0.00)`). A format with seconds rounds the time of day to the displayed second (or
fraction of a second); a format without seconds cuts it to the displayed field (float noise is first rounded off at a millisecond), so
23:59:59.5 is `23:59` with `h:mm` and `0:00:00` with `h:mm:ss`; a date-only code never rolls the date forward. These time rules are
an assumption (unverified, see below). Dates before 1900-01-01 (a negative serial) and after 9999-12-31, and numbers too large for a
time, date or fraction field, show `####` (never an exponent), unless a section for negative numbers handles them (`[h]:mm;-[h]:mm`).

`General` fits 11 characters by default (the canvas). `formatValue(x, code, { fullGeneral: true })`, which `TEXT()` uses, has no column:
15 significant digits, `E+nn` from 1e21 and below 1e-6 (`TEXT(1/3,"General")` is `0.333333333333333`); `@` and text-only sections follow.
`TEXT()` reads text as Excel does: a decimal, `$5`, `1,234`, `50%`, `(5)` (as -5) or a date (`"2024-01-15"`); other text passes through unchanged. A value that overflows after many `%` signs shows `####`.

## Limits

Excel's own: a code is at most 255 characters, a fraction denominator at most five digits (variable `?????` or fixed `99999`), fractional
seconds at most three digits; anything beyond is an invalid code. The nearest fraction comes from the continued fraction of the value
(O(digits)), the comma handling is one pass, so a format costs the same however it is written; `limits.test.ts` formats adversarial
255-character codes within a time budget.

## Invalid codes

A code Excel would reject is invalid: an unterminated quote or bracket, a dangling `\`, `_` or `*`, more than four
sections, a letter that is not a token, two colours or two conditions in one section, `[Color0]`/`[Color57]`, an unsupported bracket
item (`[DBNum1]`), digit placeholders or an exponent in a date format, `E+` without digits, `General` mixed with digits.
`formatValue` shows the value in General and sets `invalid: true`; `TEXT()` returns `#VALUE!`. An empty code is General for
display and `""` for `TEXT()`.

## Where it differs from Excel or LibreOffice on purpose

The conformance table `packages/sheets-core/src/number-format/number-format-oracle.json` (generated by
`fixtures/xlsx/generators/build_numfmt_oracle.py` from LibreOffice, see the script header) holds every row with
LibreOffice's string and the string the engine must produce. Rows where Excel is known to differ from LibreOffice carry
`excel: true` and the reason; those Excel strings are self-authored: written from what Excel is believed to do, partly by a small Python renderer in
the generator (same author as the engine, so not independent evidence), never from running Excel (real Excel files come later, plan decision D8).

- **How to read the table**: `libreoffice` is the oracle. Rows with `excel: true` replace LibreOffice's string by the one Excel is believed to
  show; the Excel string is computed by the generator's own small renderer (same author, so it is not independent evidence). Of those rows,
  the ones tagged `unverified` are assumptions pending real Excel files (plan D8): the time rounding rules, the case of `am/pm`, and the sign of
  explicit-condition sections (LibreOffice drops the minus sign in any section after the first; Excel may not). The other corrections are documented
  Excel behaviour: negative dates and times `####`, serials 0 to 60 of the 1900 system, the system date tokens, `0.#` keeping its decimal point,
  General's 11-character width, and the sign of a negative that rounds to zero.
- The headline match rate is the strict LibreOffice-only one: 802 of 1,023 rows (the rest are the corrections above, listed separately).
- `[$-409]`-style language ids only record the locale; month and day names are English. Separators are `.` and `,`.
  Calendars other than Gregorian (`[$-...]` calendar bits, `B1`/`B2`, `e`, `g`) are invalid.
- A format with only literal text and an `@` shows a number in General (LibreOffice does the same).
