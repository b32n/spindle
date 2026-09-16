# Formula System

The formula system gives `@b32nio/spindle-sheets-core` spreadsheet calculation:
a parser, a library of **146 built-in functions**, a dependency graph with
**topological recalculation**, **dynamic arrays (spill)**, and the metadata
that powers **formula autocomplete**. Everything lives under
`packages/sheets-core/src/formula-parser/` (plus `formula-graph.ts` and
`spill.ts` alongside `workbook.ts`).

## Function library

Functions are organised by category, one module per category under
`formula-parser/functions/`:

| Module | Category | Count |
|---|---|---|
| `math.ts` | math & aggregation | 37 |
| `logical.ts` | logical | 11 |
| `text.ts` | text | 21 |
| `lookup.ts` | lookup & reference | 12 |
| `stats.ts` | statistical | 13 |
| `date.ts` | date & time | 19 |
| `information.ts` | information | 11 |
| `financial.ts` | financial | 9 |
| `reference.ts` | reference-aware | 7 |
| `array.ts` | dynamic arrays | 6 |

`functions/index.ts` is the registry. It exposes three maps, by how the
evaluator must call the function:

- **`eagerFunctions`** — arguments are fully evaluated before the function
  runs. The default; most functions are eager.
- **`lazyFunctions`** — the function receives argument *thunks* so it can
  short-circuit and catch errors from a branch it does not take
  (`IF`, `IFS`, `SWITCH`, `IFERROR`, `IFNA`, the `IS*` predicates).
- **`refFunctions`** — the function receives the argument's *AST* rather than
  its value, because it needs the reference itself (`ROW`, `COLUMN`,
  `OFFSET`, `INDIRECT`, `ISREF`, `ISFORMULA`, `CELL`).

`volatileFunctions` is the set of functions whose result can change with no
input change (`RAND`, `RANDBETWEEN`, `NOW`, `TODAY`, `OFFSET`, `INDIRECT`).

### Function-metadata catalog

`functions/catalog.ts` holds a `FunctionDoc` for every function — its
signature, argument names (with optional/variadic flags), category and a
one-line description:

```typescript
interface FunctionDoc {
  name: string;
  category: FunctionCategory;
  args: { name: string; optional?: boolean; variadic?: boolean }[];
  description: string;
}
```

A test enforces two-way completeness: every registered function has a
catalog entry and vice-versa. The catalog is exported from the package's
public API and drives formula autocomplete in the React editors.

## Parser

`FormulaParser` (`parser.ts`) turns a formula string into a
`ParsedFormulaNode` AST and collects the cells it depends on. The workbook
then converts that numeric AST into a **stable AST** (`stable-ast.ts`) in
which every reference is a `(rowId, colId)` stable-ID pair instead of a
numeric coordinate. Because inserts, deletes and sorts only renumber the
sheet's order maps, stable-ID references keep pointing at the same logical
cells; the displayed `A1` text is regenerated from the AST on each evaluation.

Evaluation walks the AST and dispatches each function call down one of the
three paths above (reference → lazy → eager).

## Dependency graph and recalculation

`FormulaGraphImpl` (`formula-graph.ts`) tracks which formulas depend on
which cells:

- **`cellDependents`** — a reverse index from *any* referenced cell key
  (a value cell or a formula cell) to the formulas that read it. Because it
  covers value cells too, changing a plain value correctly recalculates the
  formulas above it.
- **Range dependencies** are stored as **rectangles** of two stable corner
  keys, not expanded cell-by-cell. Containment is tested at recalc time, so
  a cell that is empty when the formula is entered — or a row/column
  inserted into the range later — is still tracked.

On a change, `collectDirty` walks the transitive dependents (plus every
volatile cell) and records the dependency edges among them.
`topologicalOrder` (Kahn's algorithm) then orders that set so each cell is
recomputed **exactly once**, dependencies before dependents; any cell in, or
downstream of, a cycle is reported as `#CIRCULAR!`.

## Dynamic arrays (spill)

A formula that returns a 2D array — `SEQUENCE`, `FILTER`, `SORT`, `SORTBY`,
`UNIQUE`, `SPLIT` — **spills** into a rectangular block of cells. The anchor
cell holds the formula and the array's top-left value; the rest of the array
is a **derived overlay** held in a per-sheet `SpillIndex` (`spill.ts`). The
overlay is never written to `sheet.cells` and never synced over the
collaboration CRDT — each client recomputes it from the anchor formula. A
target cell that already holds content yields `#SPILL!`.

Recalculation runs to a fixed point: when a recomputed formula spills (or
stops spilling), its footprint cells are fed back as recalc seeds, so a
formula that reads a spilled cell stays in sync.

## Autocomplete and parameter help

`analyzeFormula` (`formula-context.ts`) is a lightweight backward scanner —
not the full parser — that inspects the in-progress formula text and the
caret position to report the partial function-name token and/or the
enclosing function call and argument index. Combined with the catalog, it
drives the autocomplete dropdown and the signature tooltip in the React
in-cell editor and formula bar.

## Reference adjustment

When formulas are copied or filled, `formula-adjust.ts` rewrites their
references — relative references shift with the copy offset, absolute
(`$`-marked) references stay put, mixed references shift only their relative
part. Structural edits (insert/delete/sort) need no formula rewriting: the
stable AST already references cells by ID.

## Function Reference

All 146 built-in functions, grouped by category:

### Math (37 functions)

| Function | Description |
|---|---|
| `SUM(number1, [number2, ...])` | Adds all the numbers given as arguments. |
| `AVERAGE(number1, [number2, ...])` | Returns the arithmetic mean of the arguments. |
| `COUNT(value1, [value2, ...])` | Counts how many of the arguments are numbers. |
| `COUNTA(value1, [value2, ...])` | Counts how many of the arguments are non-empty. |
| `COUNTBLANK(range)` | Counts the empty cells in a range. |
| `MAX(number1, [number2, ...])` | Returns the largest of the arguments. |
| `MIN(number1, [number2, ...])` | Returns the smallest of the arguments. |
| `PRODUCT(number1, [number2, ...])` | Multiplies all the numbers given as arguments. |
| `SUMPRODUCT(array1, [array2, ...])` | Sums the products of corresponding array entries. |
| `SUMIF(range, criterion, [sum_range])` | Adds the cells that meet a criterion. |
| `SUMIFS(sum_range, criteria_range1, criterion1, [criteria_range, criterion, ...])` | Adds the cells that meet multiple criteria. |
| `COUNTIF(range, criterion)` | Counts the cells that meet a criterion. |
| `COUNTIFS(criteria_range1, criterion1, [criteria_range, criterion, ...])` | Counts the cells that meet multiple criteria. |
| `AVERAGEIF(range, criterion, [average_range])` | Averages the cells that meet a criterion. |
| `AVERAGEIFS(average_range, criteria_range1, criterion1, [criteria_range, criterion, ...])` | Averages the cells that meet multiple criteria. |
| `SUBTOTAL(function_num, ref1, [ref2, ...])` | Returns a subtotal (sum, average, count, …) chosen by a function number. |
| `ABS(number)` | Returns the absolute value of a number. |
| `SIGN(number)` | Returns the sign of a number: -1, 0 or 1. |
| `INT(number)` | Rounds a number down to the nearest integer. |
| `TRUNC(number, [digits])` | Truncates a number to a number of digits. |
| `ROUND(number, [digits])` | Rounds a number to a number of digits. |
| `ROUNDUP(number, [digits])` | Rounds a number away from zero. |
| `ROUNDDOWN(number, [digits])` | Rounds a number toward zero. |
| `MROUND(number, multiple)` | Rounds a number to the nearest multiple. |
| `CEILING(number, [significance])` | Rounds a number up to the nearest multiple of significance. |
| `FLOOR(number, [significance])` | Rounds a number down to the nearest multiple of significance. |
| `MOD(number, divisor)` | Returns the remainder after division. |
| `POWER(base, exponent)` | Raises a number to a power. |
| `SQRT(number)` | Returns the positive square root of a number. |
| `EXP(exponent)` | Returns e raised to the given power. |
| `LN(number)` | Returns the natural logarithm of a number. |
| `LOG(number, [base])` | Returns the logarithm of a number to a given base. |
| `LOG10(number)` | Returns the base-10 logarithm of a number. |
| `GCD(number1, [number2, ...])` | Returns the greatest common divisor. |
| `LCM(number1, [number2, ...])` | Returns the least common multiple. |
| `RAND()` | Returns a random number between 0 and 1. |
| `RANDBETWEEN(low, high)` | Returns a random integer between two values. |

### Logical (11 functions)

| Function | Description |
|---|---|
| `AND(logical1, [logical2, ...])` | Returns TRUE if all arguments are TRUE. |
| `OR(logical1, [logical2, ...])` | Returns TRUE if any argument is TRUE. |
| `XOR(logical1, [logical2, ...])` | Returns TRUE if an odd number of arguments are TRUE. |
| `NOT(logical)` | Reverses the logic of its argument. |
| `TRUE()` | Returns the logical value TRUE. |
| `FALSE()` | Returns the logical value FALSE. |
| `IF(condition, value_if_true, [value_if_false])` | Returns one value if a condition is true and another if it is false. |
| `IFERROR(value, value_if_error)` | Returns a fallback value if the first argument is an error. |
| `IFNA(value, value_if_na)` | Returns a fallback value if the first argument is #N/A. |
| `IFS(condition1, value1, [condition, value, ...])` | Returns the value for the first condition that is true. |
| `SWITCH(expression, case1, value1, [case, value / default, ...])` | Compares an expression against cases and returns the first match. |

### Text (21 functions)

| Function | Description |
|---|---|
| `CONCAT(text1, [text2, ...])` | Joins several text values into one. |
| `CONCATENATE(text1, [text2, ...])` | Joins several text values into one. |
| `TEXTJOIN(delimiter, ignore_empty, text1, [text2, ...])` | Joins text values together with a delimiter. |
| `LEFT(text, [num_chars])` | Returns the first characters of a text value. |
| `RIGHT(text, [num_chars])` | Returns the last characters of a text value. |
| `MID(text, start, num_chars)` | Returns characters from the middle of a text value. |
| `LEN(text)` | Returns the number of characters in a text value. |
| `FIND(find_text, within_text, [start])` | Finds one text value within another (case-sensitive). |
| `SEARCH(find_text, within_text, [start])` | Finds one text value within another (case-insensitive, wildcards). |
| `SUBSTITUTE(text, old_text, new_text, [instance])` | Replaces occurrences of old text with new text. |
| `REPLACE(old_text, start, num_chars, new_text)` | Replaces part of a text value by position. |
| `UPPER(text)` | Converts text to uppercase. |
| `LOWER(text)` | Converts text to lowercase. |
| `PROPER(text)` | Capitalizes the first letter of each word. |
| `TRIM(text)` | Removes extra spaces from text. |
| `TEXT(value, format)` | Formats a number as text using a format string. |
| `VALUE(text)` | Converts a text value to a number. |
| `REPT(text, count)` | Repeats text a given number of times. |
| `CHAR(number)` | Returns the character for a character code. |
| `CODE(text)` | Returns the character code of the first character. |
| `EXACT(text1, text2)` | Returns TRUE if two text values are exactly equal. |

### Lookup (12 functions)

| Function | Description |
|---|---|
| `VLOOKUP(lookup_value, table, col_index, [is_sorted])` | Searches a range's first column and returns a value from a column. |
| `HLOOKUP(lookup_value, table, row_index, [is_sorted])` | Searches a range's first row and returns a value from a row. |
| `XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])` | Searches a range and returns the matching item from another range. |
| `MATCH(lookup_value, lookup_array, [match_type])` | Returns the position of a value within a range. |
| `XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])` | Returns the position of a value within a range. |
| `INDEX(array, [row], [column])` | Returns the value at a position in a range. |
| `LOOKUP(lookup_value, lookup_vector, [result_vector])` | Looks up a value in a vector and returns a corresponding result. |
| `CHOOSE(index, value1, [value2, ...])` | Returns one of a list of values, chosen by index. |
| `ROWS(range)` | Returns the number of rows in a range. |
| `COLUMNS(range)` | Returns the number of columns in a range. |
| `ADDRESS(row, column, [abs_num], [a1], [sheet])` | Builds a cell-reference string from row and column numbers. |
| `HYPERLINK(url, [label])` | Creates a hyperlink with optional display text. |

### Statistical (13 functions)

| Function | Description |
|---|---|
| `MEDIAN(number1, [number2, ...])` | Returns the median of the arguments. |
| `MODE(number1, [number2, ...])` | Returns the most frequently occurring value. |
| `VAR(number1, [number2, ...])` | Estimates variance from a sample. |
| `VARP(number1, [number2, ...])` | Calculates variance of an entire population. |
| `STDEV(number1, [number2, ...])` | Estimates standard deviation from a sample. |
| `STDEVP(number1, [number2, ...])` | Calculates standard deviation of an entire population. |
| `PERCENTILE(array, k)` | Returns the k-th percentile of values (k between 0 and 1). |
| `QUARTILE(array, quart)` | Returns a quartile (0–4) of a data set. |
| `RANK(number, ref, [order])` | Returns the rank of a number within a list. |
| `LARGE(array, k)` | Returns the k-th largest value in a data set. |
| `SMALL(array, k)` | Returns the k-th smallest value in a data set. |
| `CORREL(array1, array2)` | Returns the correlation coefficient of two ranges. |
| `COVAR(array1, array2)` | Returns the population covariance of two ranges. |

### Date & Time (19 functions)

| Function | Description |
|---|---|
| `TODAY()` | Returns the current date. |
| `NOW()` | Returns the current date and time. |
| `DATE(year, month, day)` | Builds a date from year, month and day. |
| `TIME(hour, minute, second)` | Builds a time from hour, minute and second. |
| `YEAR(date)` | Returns the year of a date. |
| `MONTH(date)` | Returns the month of a date. |
| `DAY(date)` | Returns the day of the month of a date. |
| `HOUR(time)` | Returns the hour of a time. |
| `MINUTE(time)` | Returns the minute of a time. |
| `SECOND(time)` | Returns the second of a time. |
| `WEEKDAY(date, [type])` | Returns the day of the week of a date. |
| `WEEKNUM(date, [type])` | Returns the week number of a date. |
| `EDATE(start_date, months)` | Returns the date a number of months from a start date. |
| `EOMONTH(start_date, months)` | Returns the last day of the month a number of months away. |
| `DATEDIF(start_date, end_date, unit)` | Returns the difference between two dates in a chosen unit. |
| `NETWORKDAYS(start_date, end_date, [holidays])` | Counts the working days between two dates. |
| `WORKDAY(start_date, days, [holidays])` | Returns a date a number of working days away. |
| `DATEVALUE(date_text)` | Converts a date written as text to a date. |
| `TIMEVALUE(time_text)` | Converts a time written as text to a time. |

### Information (11 functions)

| Function | Description |
|---|---|
| `ISBLANK(value)` | Returns TRUE if the value is empty. |
| `ISNUMBER(value)` | Returns TRUE if the value is a number. |
| `ISTEXT(value)` | Returns TRUE if the value is text. |
| `ISLOGICAL(value)` | Returns TRUE if the value is a boolean. |
| `ISERROR(value)` | Returns TRUE if the value is any error. |
| `ISERR(value)` | Returns TRUE if the value is an error other than #N/A. |
| `ISNA(value)` | Returns TRUE if the value is the #N/A error. |
| `TYPE(value)` | Returns a number indicating the type of a value. |
| `ERROR.TYPE(value)` | Returns a number identifying an error value. |
| `NA()` | Returns the #N/A error value. |
| `N(value)` | Converts a value to a number. |

### Financial (9 functions)

| Function | Description |
|---|---|
| `PMT(rate, nper, pv, [fv], [type])` | Returns the periodic payment for a loan or annuity. |
| `FV(rate, nper, pmt, [pv], [type])` | Returns the future value of an investment. |
| `PV(rate, nper, pmt, [fv], [type])` | Returns the present value of an investment. |
| `NPER(rate, pmt, pv, [fv], [type])` | Returns the number of periods for an investment. |
| `RATE(nper, pmt, pv, [fv], [type], [guess])` | Returns the interest rate per period of an annuity. |
| `NPV(rate, value1, [value2, ...])` | Returns the net present value of a series of cash flows. |
| `IRR(values, [guess])` | Returns the internal rate of return for a series of cash flows. |
| `IPMT(rate, per, nper, pv, [fv], [type])` | Returns the interest portion of a given payment. |
| `PPMT(rate, per, nper, pv, [fv], [type])` | Returns the principal portion of a given payment. |

### Reference (7 functions)

| Function | Description |
|---|---|
| `ROW([reference])` | Returns the row number of a reference. |
| `COLUMN([reference])` | Returns the column number of a reference. |
| `OFFSET(reference, rows, cols, [height], [width])` | Returns a reference shifted from a starting reference. |
| `INDIRECT(ref_text, [a1])` | Returns the reference specified by a text string. |
| `ISREF(value)` | Returns TRUE if the value is a reference. |
| `ISFORMULA(reference)` | Returns TRUE if the referenced cell contains a formula. |
| `CELL(info_type, [reference])` | Returns information about a cell. |

### Array (6 functions)

| Function | Description |
|---|---|
| `SEQUENCE(rows, [columns], [start], [step])` | Generates a sequence of numbers in an array. |
| `UNIQUE(array, [by_column], [exactly_once])` | Returns the unique rows of a range. |
| `SORT(array, [sort_index], [sort_order], [by_column])` | Sorts the rows of a range. |
| `SORTBY(array, by_array1, [sort_order1], [by_array, sort_order, ...])` | Sorts a range by the values in one or more other ranges. |
| `FILTER(array, include, [if_empty])` | Returns the rows of a range that meet a condition. |
| `SPLIT(text, delimiter, [split_by_each], [remove_empty])` | Splits text into an array around a delimiter. |
