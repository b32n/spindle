# Function list

The sheets engine ships **229 built-in functions**. This is the published list, by category, with the evidence for
each function's Excel parity. It is checked against the function catalog (`formula-parser/functions/catalog.ts`) by
`functions-doc.test.ts`, so a function cannot be added without being listed here.

How to read the **Excel parity** column:

- *N oracle rows, LibreOffice-verified*: the formula conformance suite (`packages/sheets-core/src/formula-conformance-oracle.json`,
  see [Conformance suite](formulas.md#conformance-suite)) runs N formulas for the function through the engine and compares them with the
  values LibreOffice 25.2 computes. LibreOffice is the available oracle, not Excel itself.
- *Excel-documented (unverified)*: LibreOffice 25.2 lacks the function or differs from Excel, so the row carries the behaviour Excel's
  documentation gives (tagged `excel: true, unverified: true` in the table) until a real Excel file confirms it (plan decision D8).
  This covers the rows where LibreOffice's error code or calendar differs from Excel's.
- *unit tests*: older functions covered by their module's unit tests (a few also appear in the 159-formula probe).

Notes that apply to the whole list:

- Dotted names (`STDEV.S`, `WORKDAY.INTL`) are registered as written; the `_xlfn.` spellings Excel stores in files are mapped on import (plan S4).
- An omitted argument in the middle of a call (`INDEX(A1:C3,,2)`, `NUMBERVALUE(A1,,",")`) keeps its position for the functions that
  take one (`INDEX`, `NUMBERVALUE`, `FIXED`, `WORKDAY.INTL`, `NETWORKDAYS.INTL`); every other function still skips it.
- Functions that return arrays spill into the cells below and to the right (see [Dynamic arrays](formulas.md#dynamic-arrays-spill)).
- The regular-expression functions (`REGEXMATCH`, `REGEXEXTRACT`, `REGEXREPLACE`) are Google Sheets only: they are not Excel functions and are
  outside the Excel parity set.
- Deferred until Excel-computed files exist (plan S12, D8; LibreOffice has no oracle): `TEXTBEFORE`, `TEXTAFTER`, `TEXTSPLIT`.
- Not implemented: `LET`, `LAMBDA` and its helpers (`MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`), which need names in scope in the evaluator (post-launch).

## Math and aggregation (72)

| Function | Description | Excel parity |
|---|---|---|
| `ABS(number)` | Returns the absolute value of a number. | 1 oracle rows, LibreOffice-verified |
| `ACOS(number)` | Returns the arccosine of a number, in radians (-1 to 1). | 4 oracle rows, LibreOffice-verified |
| `ACOSH(number)` | Returns the inverse hyperbolic cosine of a number (at least 1). | 3 oracle rows, 1 Excel-documented (unverified) |
| `AGGREGATE(function_num, options, ref1, [ref2 or k, ...])` | Applies one of 19 functions to a list, optionally ignoring errors. | 45 oracle rows, 6 Excel-documented (unverified) |
| `ASIN(number)` | Returns the arcsine of a number, in radians (-1 to 1). | 4 oracle rows, LibreOffice-verified |
| `ASINH(number)` | Returns the inverse hyperbolic sine of a number. | 3 oracle rows, LibreOffice-verified |
| `ATAN(number)` | Returns the arctangent of a number, in radians. | 3 oracle rows, LibreOffice-verified |
| `ATAN2(x_num, y_num)` | Returns the arctangent of the point (x, y), in radians. | 5 oracle rows, 1 Excel-documented (unverified) |
| `ATANH(number)` | Returns the inverse hyperbolic tangent of a number (between -1 and 1). | 4 oracle rows, 2 Excel-documented (unverified) |
| `AVERAGE(number1, [number2, ...])` | Returns the arithmetic mean of the arguments. | 1 oracle rows, LibreOffice-verified |
| `AVERAGEIF(range, criterion, [average_range])` | Averages the cells that meet a criterion. | 1 oracle rows, LibreOffice-verified |
| `AVERAGEIFS(average_range, criteria_range1, criterion1, [criteria_range, criterion, ...])` | Averages the cells that meet multiple criteria. | unit tests |
| `CEILING(number, [significance])` | Rounds a number up to the nearest multiple of significance. | 1 oracle rows, LibreOffice-verified |
| `CEILING.MATH(number, [significance], [mode])` | Rounds a number up to the nearest multiple of significance. | 10 oracle rows, LibreOffice-verified |
| `COMBIN(number, number_chosen)` | Returns the number of combinations for a number of items. | 7 oracle rows, 2 Excel-documented (unverified) |
| `COS(number)` | Returns the cosine of an angle in radians. | 6 oracle rows, LibreOffice-verified |
| `COSH(number)` | Returns the hyperbolic cosine of a number. | 3 oracle rows, LibreOffice-verified |
| `COUNT(value1, [value2, ...])` | Counts how many of the arguments are numbers. | 1 oracle rows, LibreOffice-verified |
| `COUNTA(value1, [value2, ...])` | Counts how many of the arguments are non-empty. | 1 oracle rows, LibreOffice-verified |
| `COUNTBLANK(range)` | Counts the empty cells in a range. | 1 oracle rows, LibreOffice-verified |
| `COUNTIF(range, criterion)` | Counts the cells that meet a criterion. | 1 oracle rows, LibreOffice-verified |
| `COUNTIFS(criteria_range1, criterion1, [criteria_range, criterion, ...])` | Counts the cells that meet multiple criteria. | 1 oracle rows, LibreOffice-verified |
| `DEGREES(angle)` | Converts radians to degrees. | 3 oracle rows, LibreOffice-verified |
| `EVEN(number)` | Rounds a number away from zero to the nearest even integer. | 7 oracle rows, LibreOffice-verified |
| `EXP(exponent)` | Returns e raised to the given power. | 1 oracle rows, LibreOffice-verified |
| `FACT(number)` | Returns the factorial of a number. | 8 oracle rows, 2 Excel-documented (unverified) |
| `FACTDOUBLE(number)` | Returns the double factorial of a number. | 6 oracle rows, 3 Excel-documented (unverified) |
| `FLOOR(number, [significance])` | Rounds a number down to the nearest multiple of significance. | 1 oracle rows, LibreOffice-verified |
| `FLOOR.MATH(number, [significance], [mode])` | Rounds a number down to the nearest multiple of significance. | 6 oracle rows, LibreOffice-verified |
| `GCD(number1, [number2, ...])` | Returns the greatest common divisor. | 1 oracle rows, LibreOffice-verified |
| `INT(number)` | Rounds a number down to the nearest integer. | 1 oracle rows, LibreOffice-verified |
| `ISEVEN(number)` | Returns TRUE when a number is even. | 6 oracle rows, LibreOffice-verified |
| `ISODD(number)` | Returns TRUE when a number is odd. | 4 oracle rows, LibreOffice-verified |
| `LCM(number1, [number2, ...])` | Returns the least common multiple. | 1 oracle rows, LibreOffice-verified |
| `LN(number)` | Returns the natural logarithm of a number. | 1 oracle rows, LibreOffice-verified |
| `LOG(number, [base])` | Returns the logarithm of a number to a given base. | 1 oracle rows, LibreOffice-verified |
| `LOG10(number)` | Returns the base-10 logarithm of a number. | unit tests |
| `MAX(number1, [number2, ...])` | Returns the largest of the arguments. | 1 oracle rows, LibreOffice-verified |
| `MAXIFS(max_range, criteria_range1, criterion1, [criteria_range, criterion, ...])` | Returns the largest value among the cells that meet multiple criteria. | 11 oracle rows, 1 Excel-documented (unverified) |
| `MIN(number1, [number2, ...])` | Returns the smallest of the arguments. | 1 oracle rows, LibreOffice-verified |
| `MINIFS(min_range, criteria_range1, criterion1, [criteria_range, criterion, ...])` | Returns the smallest value among the cells that meet multiple criteria. | 5 oracle rows, LibreOffice-verified |
| `MOD(number, divisor)` | Returns the remainder after division. | 1 oracle rows, LibreOffice-verified |
| `MROUND(number, multiple)` | Rounds a number to the nearest multiple. | 1 oracle rows, LibreOffice-verified |
| `ODD(number)` | Rounds a number away from zero to the nearest odd integer. | 6 oracle rows, LibreOffice-verified |
| `PERMUT(number, number_chosen)` | Returns the number of permutations for a number of items. | 4 oracle rows, 1 Excel-documented (unverified) |
| `PI()` | Returns the value of pi. | 4 oracle rows, LibreOffice-verified |
| `POWER(base, exponent)` | Raises a number to a power. | 1 oracle rows, LibreOffice-verified |
| `PRODUCT(number1, [number2, ...])` | Multiplies all the numbers given as arguments. | 1 oracle rows, LibreOffice-verified |
| `QUOTIENT(numerator, denominator)` | Returns the integer part of a division. | 7 oracle rows, 1 Excel-documented (unverified) |
| `RADIANS(angle)` | Converts degrees to radians. | 4 oracle rows, LibreOffice-verified |
| `RAND()` | Returns a random number between 0 and 1. | unit tests |
| `RANDBETWEEN(low, high)` | Returns a random integer between two values. | unit tests |
| `ROUND(number, [digits])` | Rounds a number to a number of digits. | 1 oracle rows, LibreOffice-verified |
| `ROUNDDOWN(number, [digits])` | Rounds a number toward zero. | 1 oracle rows, LibreOffice-verified |
| `ROUNDUP(number, [digits])` | Rounds a number away from zero. | 1 oracle rows, LibreOffice-verified |
| `SIGN(number)` | Returns the sign of a number: -1, 0 or 1. | 1 oracle rows, LibreOffice-verified |
| `SIN(number)` | Returns the sine of an angle in radians. | 8 oracle rows, LibreOffice-verified |
| `SINH(number)` | Returns the hyperbolic sine of a number. | 3 oracle rows, LibreOffice-verified |
| `SQRT(number)` | Returns the positive square root of a number. | 1 oracle rows, LibreOffice-verified |
| `SQRTPI(number)` | Returns the square root of a number times pi. | 5 oracle rows, 1 Excel-documented (unverified) |
| `SUBTOTAL(function_num, ref1, [ref2, ...])` | Returns a subtotal (sum, average, count, …) chosen by a function number. | 1 oracle rows, LibreOffice-verified |
| `SUM(number1, [number2, ...])` | Adds all the numbers given as arguments. | 3 oracle rows, LibreOffice-verified |
| `SUMIF(range, criterion, [sum_range])` | Adds the cells that meet a criterion. | 1 oracle rows, LibreOffice-verified |
| `SUMIFS(sum_range, criteria_range1, criterion1, [criteria_range, criterion, ...])` | Adds the cells that meet multiple criteria. | 1 oracle rows, LibreOffice-verified |
| `SUMPRODUCT(array1, [array2, ...])` | Sums the products of corresponding array entries. | 2 oracle rows, LibreOffice-verified |
| `SUMSQ(number1, [number2, ...])` | Returns the sum of the squares of the arguments. | 6 oracle rows, LibreOffice-verified |
| `SUMX2MY2(array_x, array_y)` | Returns the sum of the differences of the squares of two arrays. | 5 oracle rows, 1 Excel-documented (unverified) |
| `SUMX2PY2(array_x, array_y)` | Returns the sum of the sums of the squares of two arrays. | 4 oracle rows, 1 Excel-documented (unverified) |
| `SUMXMY2(array_x, array_y)` | Returns the sum of the squares of the differences of two arrays. | 4 oracle rows, 2 Excel-documented (unverified) |
| `TAN(number)` | Returns the tangent of an angle in radians. | 4 oracle rows, LibreOffice-verified |
| `TANH(number)` | Returns the hyperbolic tangent of a number. | 3 oracle rows, LibreOffice-verified |
| `TRUNC(number, [digits])` | Truncates a number to a number of digits. | 1 oracle rows, LibreOffice-verified |

## Logical (11)

| Function | Description | Excel parity |
|---|---|---|
| `AND(logical1, [logical2, ...])` | Returns TRUE if all arguments are TRUE. | 1 oracle rows, LibreOffice-verified |
| `FALSE()` | Returns the logical value FALSE. | unit tests |
| `IF(condition, value_if_true, [value_if_false])` | Returns one value if a condition is true and another if it is false. | 1 oracle rows, LibreOffice-verified |
| `IFERROR(value, value_if_error)` | Returns a fallback value if the first argument is an error. | 1 oracle rows, LibreOffice-verified |
| `IFNA(value, value_if_na)` | Returns a fallback value if the first argument is #N/A. | unit tests |
| `IFS(condition1, value1, [condition, value, ...])` | Returns the value for the first condition that is true. | 1 oracle rows, LibreOffice-verified |
| `NOT(logical)` | Reverses the logic of its argument. | unit tests |
| `OR(logical1, [logical2, ...])` | Returns TRUE if any argument is TRUE. | unit tests |
| `SWITCH(expression, case1, value1, [case, value / default, ...])` | Compares an expression against cases and returns the first match. | 1 oracle rows, LibreOffice-verified |
| `TRUE()` | Returns the logical value TRUE. | unit tests |
| `XOR(logical1, [logical2, ...])` | Returns TRUE if an odd number of arguments are TRUE. | 1 oracle rows, LibreOffice-verified |

## Text (32)

| Function | Description | Excel parity |
|---|---|---|
| `ARABIC(text)` | Converts a Roman numeral to a number. | 7 oracle rows, 1 Excel-documented (unverified) |
| `BASE(number, radix, [min_length])` | Writes a number as text in another base (2 to 36). | 11 oracle rows, 4 Excel-documented (unverified) |
| `CHAR(number)` | Returns the character for a character code. | unit tests |
| `CLEAN(text)` | Removes the non-printing characters (codes 0-31) from text. | 8 oracle rows, 1 Excel-documented (unverified) |
| `CODE(text)` | Returns the character code of the first character. | unit tests |
| `CONCAT(text1, [text2, ...])` | Joins several text values into one. | 1 oracle rows, LibreOffice-verified |
| `CONCATENATE(text1, [text2, ...])` | Joins several text values into one. | unit tests |
| `DECIMAL(text, radix)` | Converts text in another base (2 to 36) to a number. | 8 oracle rows, 3 Excel-documented (unverified) |
| `DOLLAR(number, [decimals])` | Rounds a number and returns it as dollar-formatted text. | 6 oracle rows, 2 Excel-documented (unverified) |
| `EXACT(text1, text2)` | Returns TRUE if two text values are exactly equal. | 1 oracle rows, LibreOffice-verified |
| `FIND(find_text, within_text, [start])` | Finds one text value within another (case-sensitive). | 1 oracle rows, LibreOffice-verified |
| `FIXED(number, [decimals], [no_commas])` | Rounds a number and returns it as text, with thousands commas. | 12 oracle rows, 1 Excel-documented (unverified) |
| `LEFT(text, [num_chars])` | Returns the first characters of a text value. | 1 oracle rows, LibreOffice-verified |
| `LEN(text)` | Returns the number of characters in a text value. | 1 oracle rows, LibreOffice-verified |
| `LOWER(text)` | Converts text to lowercase. | 1 oracle rows, LibreOffice-verified |
| `MID(text, start, num_chars)` | Returns characters from the middle of a text value. | 1 oracle rows, LibreOffice-verified |
| `NUMBERVALUE(text, [decimal_separator], [group_separator])` | Converts text to a number with the separators you give, whatever the locale. | 18 oracle rows, 4 Excel-documented (unverified) |
| `PROPER(text)` | Capitalizes the first letter of each word. | 1 oracle rows, LibreOffice-verified |
| `REPLACE(old_text, start, num_chars, new_text)` | Replaces part of a text value by position. | unit tests |
| `REPT(text, count)` | Repeats text a given number of times. | 1 oracle rows, LibreOffice-verified |
| `RIGHT(text, [num_chars])` | Returns the last characters of a text value. | 1 oracle rows, LibreOffice-verified |
| `ROMAN(number, [form])` | Writes a number as a Roman numeral (1 to 3999). | 14 oracle rows, LibreOffice-verified |
| `SEARCH(find_text, within_text, [start])` | Finds one text value within another (case-insensitive, wildcards). | 1 oracle rows, LibreOffice-verified |
| `SUBSTITUTE(text, old_text, new_text, [instance])` | Replaces occurrences of old text with new text. | 1 oracle rows, LibreOffice-verified |
| `T(value)` | Returns the text of a value, or empty text for anything else. | 7 oracle rows, LibreOffice-verified |
| `TEXT(value, format)` | Formats a number as text using a format string. | 4 oracle rows, LibreOffice-verified |
| `TEXTJOIN(delimiter, ignore_empty, text1, [text2, ...])` | Joins text values together with a delimiter. | 1 oracle rows, LibreOffice-verified |
| `TRIM(text)` | Removes extra spaces from text. | 1 oracle rows, LibreOffice-verified |
| `UNICHAR(number)` | Returns the character for a Unicode code point. | 7 oracle rows, 2 Excel-documented (unverified) |
| `UNICODE(text)` | Returns the Unicode code point of the first character. | 6 oracle rows, LibreOffice-verified |
| `UPPER(text)` | Converts text to uppercase. | 1 oracle rows, LibreOffice-verified |
| `VALUE(text)` | Converts a text value to a number. | 1 oracle rows, LibreOffice-verified |

## Lookup (12)

| Function | Description | Excel parity |
|---|---|---|
| `ADDRESS(row, column, [abs_num], [a1], [sheet])` | Builds a cell-reference string from row and column numbers. | 1 oracle rows, LibreOffice-verified |
| `CHOOSE(index, value1, [value2, ...])` | Returns one of a list of values, chosen by index. | 1 oracle rows, LibreOffice-verified |
| `COLUMNS(range)` | Returns the number of columns in a range. | 1 oracle rows, LibreOffice-verified |
| `HLOOKUP(lookup_value, table, row_index, [is_sorted])` | Searches a range's first row and returns a value from a row. | 1 oracle rows, LibreOffice-verified |
| `HYPERLINK(url, [label])` | Creates a hyperlink with optional display text. | unit tests |
| `INDEX(array, [row], [column])` | Returns the value at a position in a range, or a whole row or column when an index is 0. | 17 oracle rows, LibreOffice-verified |
| `LOOKUP(lookup_value, lookup_vector, [result_vector])` | Looks up a value in a vector and returns a corresponding result. | 1 oracle rows, LibreOffice-verified |
| `MATCH(lookup_value, lookup_array, [match_type])` | Returns the position of a value within a range. | 1 oracle rows, LibreOffice-verified |
| `ROWS(range)` | Returns the number of rows in a range. | 1 oracle rows, LibreOffice-verified |
| `VLOOKUP(lookup_value, table, col_index, [is_sorted])` | Searches a range's first column and returns a value from a column. | 1 oracle rows, LibreOffice-verified |
| `XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])` | Searches a range and returns the matching item from another range. | 1 oracle rows, LibreOffice-verified |
| `XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])` | Returns the position of a value within a range. | 1 oracle rows, LibreOffice-verified |

## Statistical (30)

| Function | Description | Excel parity |
|---|---|---|
| `CORREL(array1, array2)` | Returns the correlation coefficient of two ranges. | 1 oracle rows, LibreOffice-verified |
| `COVAR(array1, array2)` | Returns the population covariance of two ranges. | unit tests |
| `FORECAST(x, known_ys, known_xs)` | Predicts a value along a linear trend. | 7 oracle rows, 1 Excel-documented (unverified) |
| `FORECAST.LINEAR(x, known_ys, known_xs)` | Predicts a value along a linear trend. | 3 oracle rows, 1 Excel-documented (unverified) |
| `FREQUENCY(data_array, bins_array)` | Counts values per interval, as a column (spills). | 7 oracle rows, LibreOffice-verified |
| `GEOMEAN(number1, [number2, ...])` | Returns the geometric mean of positive numbers. | 8 oracle rows, 2 Excel-documented (unverified) |
| `INTERCEPT(known_ys, known_xs)` | Returns the intercept of the linear regression line. | 5 oracle rows, 1 Excel-documented (unverified) |
| `LARGE(array, k)` | Returns the k-th largest value in a data set. | 1 oracle rows, LibreOffice-verified |
| `MEDIAN(number1, [number2, ...])` | Returns the median of the arguments. | 1 oracle rows, LibreOffice-verified |
| `MODE(number1, [number2, ...])` | Returns the most frequently occurring value. | 1 oracle rows, Excel-documented (unverified) |
| `MODE.SNGL(number1, [number2, ...])` | Returns the most frequently occurring value. | 6 oracle rows, 2 Excel-documented (unverified) |
| `PERCENTILE(array, k)` | Returns the k-th percentile of values (k between 0 and 1). | 1 oracle rows, LibreOffice-verified |
| `PERCENTILE.EXC(array, k)` | Returns the k-th percentile of values (k strictly between 0 and 1). | 5 oracle rows, 2 Excel-documented (unverified) |
| `PERCENTILE.INC(array, k)` | Returns the k-th percentile of values (k from 0 to 1, inclusive). | 9 oracle rows, 3 Excel-documented (unverified) |
| `QUARTILE(array, quart)` | Returns a quartile (0–4) of a data set. | 1 oracle rows, LibreOffice-verified |
| `QUARTILE.EXC(array, quart)` | Returns a quartile (1-3) of a data set, exclusive. | 4 oracle rows, 2 Excel-documented (unverified) |
| `QUARTILE.INC(array, quart)` | Returns a quartile (0-4) of a data set, inclusive. | 6 oracle rows, 2 Excel-documented (unverified) |
| `RANK(number, ref, [order])` | Returns the rank of a number within a list. | 1 oracle rows, LibreOffice-verified |
| `RANK.EQ(number, ref, [order])` | Returns the rank of a number within a list; ties share the top rank. | 8 oracle rows, LibreOffice-verified |
| `RSQ(known_ys, known_xs)` | Returns the square of the Pearson correlation coefficient. | 5 oracle rows, 1 Excel-documented (unverified) |
| `SLOPE(known_ys, known_xs)` | Returns the slope of the linear regression line. | 8 oracle rows, 1 Excel-documented (unverified) |
| `SMALL(array, k)` | Returns the k-th smallest value in a data set. | 1 oracle rows, LibreOffice-verified |
| `STDEV(number1, [number2, ...])` | Estimates standard deviation from a sample. | 1 oracle rows, LibreOffice-verified |
| `STDEV.P(number1, [number2, ...])` | Calculates standard deviation of an entire population. | 4 oracle rows, LibreOffice-verified |
| `STDEV.S(number1, [number2, ...])` | Estimates standard deviation from a sample. | 7 oracle rows, 1 Excel-documented (unverified) |
| `STDEVP(number1, [number2, ...])` | Calculates standard deviation of an entire population. | unit tests |
| `VAR(number1, [number2, ...])` | Estimates variance from a sample. | unit tests |
| `VAR.P(number1, [number2, ...])` | Calculates variance of an entire population. | 6 oracle rows, 1 Excel-documented (unverified) |
| `VAR.S(number1, [number2, ...])` | Estimates variance from a sample. | 5 oracle rows, LibreOffice-verified |
| `VARP(number1, [number2, ...])` | Calculates variance of an entire population. | unit tests |

## Date and time (25)

| Function | Description | Excel parity |
|---|---|---|
| `DATE(year, month, day)` | Builds a date from year, month and day. | 1 oracle rows, LibreOffice-verified |
| `DATEDIF(start_date, end_date, unit)` | Returns the difference between two dates in a chosen unit. | 1 oracle rows, LibreOffice-verified |
| `DATEVALUE(date_text)` | Converts a date written as text to a date. | 1 oracle rows, LibreOffice-verified |
| `DAY(date)` | Returns the day of the month of a date. | 1 oracle rows, LibreOffice-verified |
| `DAYS(end_date, start_date)` | Returns the number of days between two dates. | 15 oracle rows, 6 Excel-documented (unverified) |
| `DAYS360(start_date, end_date, [method])` | Returns the days between two dates on a 360-day year (US or European method). | 14 oracle rows, 2 Excel-documented (unverified) |
| `EDATE(start_date, months)` | Returns the date a number of months from a start date. | 6 oracle rows, 1 Excel-documented (unverified) |
| `EOMONTH(start_date, months)` | Returns the last day of the month a number of months away. | 1 oracle rows, LibreOffice-verified |
| `HOUR(time)` | Returns the hour of a time. | unit tests |
| `ISOWEEKNUM(date)` | Returns the ISO 8601 week number of a date. | 11 oracle rows, 1 Excel-documented (unverified) |
| `MINUTE(time)` | Returns the minute of a time. | unit tests |
| `MONTH(date)` | Returns the month of a date. | 1 oracle rows, LibreOffice-verified |
| `NETWORKDAYS(start_date, end_date, [holidays])` | Counts the working days between two dates. | 1 oracle rows, LibreOffice-verified |
| `NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])` | Counts the working days between two dates, with a custom weekend. | 16 oracle rows, 1 Excel-documented (unverified) |
| `NOW()` | Returns the current date and time. | unit tests |
| `SECOND(time)` | Returns the second of a time. | unit tests |
| `TIME(hour, minute, second)` | Builds a time from hour, minute and second. | unit tests |
| `TIMEVALUE(time_text)` | Converts a time written as text to a time. | unit tests |
| `TODAY()` | Returns the current date. | unit tests |
| `WEEKDAY(date, [type])` | Returns the day of the week of a date. | 1 oracle rows, LibreOffice-verified |
| `WEEKNUM(date, [type])` | Returns the week number of a date. | 1 oracle rows, LibreOffice-verified |
| `WORKDAY(start_date, days, [holidays])` | Returns a date a number of working days away. | 1 oracle rows, LibreOffice-verified |
| `WORKDAY.INTL(start_date, days, [weekend], [holidays])` | Returns a date a number of working days away, with a custom weekend. | 20 oracle rows, 2 Excel-documented (unverified) |
| `YEAR(date)` | Returns the year of a date. | 1 oracle rows, LibreOffice-verified |
| `YEARFRAC(start_date, end_date, [basis])` | Returns the fraction of a year between two dates. | 17 oracle rows, 3 Excel-documented (unverified) |

## Information (12)

| Function | Description | Excel parity |
|---|---|---|
| `ERROR.TYPE(value)` | Returns a number identifying an error value. | unit tests |
| `ISBLANK(value)` | Returns TRUE if the value is empty. | 1 oracle rows, LibreOffice-verified |
| `ISERR(value)` | Returns TRUE if the value is an error other than #N/A. | unit tests |
| `ISERROR(value)` | Returns TRUE if the value is any error. | 1 oracle rows, LibreOffice-verified |
| `ISLOGICAL(value)` | Returns TRUE if the value is a boolean. | unit tests |
| `ISNA(value)` | Returns TRUE if the value is the #N/A error. | 1 oracle rows, LibreOffice-verified |
| `ISNONTEXT(value)` | Returns TRUE if the value is not text. | 7 oracle rows, LibreOffice-verified |
| `ISNUMBER(value)` | Returns TRUE if the value is a number. | 1 oracle rows, LibreOffice-verified |
| `ISTEXT(value)` | Returns TRUE if the value is text. | 1 oracle rows, LibreOffice-verified |
| `N(value)` | Converts a value to a number. | 1 oracle rows, LibreOffice-verified |
| `NA()` | Returns the #N/A error value. | unit tests |
| `TYPE(value)` | Returns a number indicating the type of a value. | 1 oracle rows, LibreOffice-verified |

## Financial (21)

| Function | Description | Excel parity |
|---|---|---|
| `DDB(cost, salvage, life, period, [factor])` | Returns the double-declining-balance depreciation for a period. | 11 oracle rows, 4 Excel-documented (unverified) |
| `EFFECT(nominal_rate, npery)` | Returns the effective annual interest rate. | 7 oracle rows, 3 Excel-documented (unverified) |
| `FV(rate, nper, pmt, [pv], [type])` | Returns the future value of an investment. | 1 oracle rows, LibreOffice-verified |
| `FVSCHEDULE(principal, schedule)` | Returns the future value of a principal after a series of interest rates. | 4 oracle rows, 1 Excel-documented (unverified) |
| `IPMT(rate, per, nper, pv, [fv], [type])` | Returns the interest portion of a given payment. | unit tests |
| `IRR(values, [guess])` | Returns the internal rate of return for a series of cash flows. | 1 oracle rows, LibreOffice-verified |
| `ISPMT(rate, per, nper, pv)` | Returns the interest paid in a period of a loan with even principal payments. | 5 oracle rows, 1 Excel-documented (unverified) |
| `MIRR(values, finance_rate, reinvest_rate)` | Returns the modified internal rate of return. | 7 oracle rows, 2 Excel-documented (unverified) |
| `NOMINAL(effect_rate, npery)` | Returns the nominal annual interest rate. | 5 oracle rows, 2 Excel-documented (unverified) |
| `NPER(rate, pmt, pv, [fv], [type])` | Returns the number of periods for an investment. | 1 oracle rows, LibreOffice-verified |
| `NPV(rate, value1, [value2, ...])` | Returns the net present value of a series of cash flows. | 1 oracle rows, LibreOffice-verified |
| `PDURATION(rate, pv, fv)` | Returns the periods an investment needs to reach a value. | 5 oracle rows, 2 Excel-documented (unverified) |
| `PMT(rate, nper, pv, [fv], [type])` | Returns the periodic payment for a loan or annuity. | 1 oracle rows, LibreOffice-verified |
| `PPMT(rate, per, nper, pv, [fv], [type])` | Returns the principal portion of a given payment. | unit tests |
| `PV(rate, nper, pmt, [fv], [type])` | Returns the present value of an investment. | 1 oracle rows, LibreOffice-verified |
| `RATE(nper, pmt, pv, [fv], [type], [guess])` | Returns the interest rate per period of an annuity. | 1 oracle rows, LibreOffice-verified |
| `RRI(nper, pv, fv)` | Returns the equivalent interest rate for the growth of an investment. | 5 oracle rows, 1 Excel-documented (unverified) |
| `SLN(cost, salvage, life)` | Returns the straight-line depreciation for one period. | 7 oracle rows, LibreOffice-verified |
| `SYD(cost, salvage, life, period)` | Returns the sum-of-years-digits depreciation for a period. | 7 oracle rows, 2 Excel-documented (unverified) |
| `XIRR(values, dates, [guess])` | Returns the internal rate of return of cash flows on given dates. | 14 oracle rows, 5 Excel-documented (unverified) |
| `XNPV(rate, values, dates)` | Returns the net present value of cash flows on given dates. | 11 oracle rows, 5 Excel-documented (unverified) |

## Reference (7)

| Function | Description | Excel parity |
|---|---|---|
| `CELL(info_type, [reference])` | Returns information about a cell. | unit tests |
| `COLUMN([reference])` | Returns the column number of a reference. | 1 oracle rows, LibreOffice-verified |
| `INDIRECT(ref_text, [a1])` | Returns the reference specified by a text string. | 1 oracle rows, LibreOffice-verified |
| `ISFORMULA(reference)` | Returns TRUE if the referenced cell contains a formula. | unit tests |
| `ISREF(value)` | Returns TRUE if the value is a reference. | unit tests |
| `OFFSET(reference, rows, cols, [height], [width])` | Returns a reference shifted from a starting reference. | 1 oracle rows, LibreOffice-verified |
| `ROW([reference])` | Returns the row number of a reference. | 1 oracle rows, LibreOffice-verified |

## Array (dynamic arrays) (7)

| Function | Description | Excel parity |
|---|---|---|
| `FILTER(array, include, [if_empty])` | Returns the rows of a range that meet a condition. | 1 oracle rows, LibreOffice-verified |
| `SEQUENCE(rows, [columns], [start], [step])` | Generates a sequence of numbers in an array. | 1 oracle rows, LibreOffice-verified |
| `SORT(array, [sort_index], [sort_order], [by_column])` | Sorts the rows of a range. | 1 oracle rows, LibreOffice-verified |
| `SORTBY(array, by_array1, [sort_order1], [by_array, sort_order, ...])` | Sorts a range by the values in one or more other ranges. | 1 oracle rows, LibreOffice-verified |
| `SPLIT(text, delimiter, [split_by_each], [remove_empty])` | Splits text into an array around a delimiter. | unit tests |
| `TRANSPOSE(array)` | Swaps the rows and columns of an array. | 9 oracle rows, 1 Excel-documented (unverified) |
| `UNIQUE(array, [by_column], [exactly_once])` | Returns the unique rows of a range. | 1 oracle rows, LibreOffice-verified |
