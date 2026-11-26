# Formula System

The formula system gives `@b32nio/spindle-sheets-core` spreadsheet calculation:
a parser, a library of **193 built-in functions**, a dependency graph with
**topological recalculation**, **dynamic arrays (spill)**, and the metadata
that powers **formula autocomplete**. Everything lives under
`packages/sheets-core/src/formula-parser/` (plus `formula-graph.ts` and
`spill.ts` alongside `workbook.ts`). The syntax it understands, and what each
form does under structural edits, is described in [Formula syntax](#formula-syntax).

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
  its value, because it needs the reference itself (`ROW`, `COLUMN`, `ROWS`,
  `COLUMNS`, `COUNTBLANK`, `OFFSET`, `INDIRECT`, `INDEX`, `ISREF`, `ISFORMULA`,
  `CELL`). `INDEX` is registered here (over a computed array it falls back to the
  ordinary lookup `INDEX`) so that `INDEX(...)` can end a range.

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
`ParsedFormulaNode` AST and collects a best-effort set of the cells it reads.
Syntax is handled by two modules: `tokenizer.ts` (references are recognised
as single tokens, because their meaning depends on their neighbours: `A:A`,
`1:1`, `Sheet1:Sheet3!`, `A1#`) and `syntax.ts` (a precedence-climbing parser
with Excel's operator order, see [Formula syntax](#formula-syntax)). A formula
that does not parse is reported with an `#ERROR! <message>` value that names the
offending token and its position; the engine no longer reads unparseable text
as a string.

The workbook then converts the numeric AST into a **stable AST**
(`stable-ast.ts`) in which every reference is a `(rowId, colId)` stable-ID pair
instead of a numeric coordinate. Because inserts, deletes and sorts only
renumber the sheet's order maps, stable-ID references keep pointing at the same
logical cells; the displayed `A1` text is regenerated from the AST on each
evaluation (so the text you typed is normalised: spacing, case of function
names and column letters, quoting of sheet names).

Evaluation walks the AST and dispatches each function call down one of the
three paths above (reference → lazy → eager).

## Formula syntax

Operators, loosest to tightest binding:

| Level | Operators | Notes |
|---|---|---|
| comparison | `=` `<>` `<` `>` `<=` `>=` | |
| concatenation | `&` | |
| additive | `+` `-` | |
| multiplicative | `*` `/` | |
| power | `^` | left-associative: `2^3^2` is 64 |
| percent | postfix `%` | `50%`, `10%^2` is 0.01, `A1+50%`; applies to a number, a range or an array element-wise |
| sign | prefix `-` `+` | binds tighter than `^`: `-2^2` is 4 |
| intersection | a space between two references | `A1:C3 B2:D4` |
| range | `:` | `A1:B2`, `A1:INDEX(A1:A5,3)` |

Reference and literal forms:

| Form | Example | Meaning |
|---|---|---|
| Whole column / row | `A:A`, `$A:$C`, `1:1`, `$2:$4`, `Sheet2!A:A` | See [whole-column and whole-row references](#whole-column-and-whole-row-references) |
| 3D reference | `Sheet1:Sheet3!A1`, `Sheet1:Sheet3!A1:B2`, `'Sheet 1:Sheet 3'!A:A` | The same cells on every sheet from the first to the last, in workbook order; valid as a function argument only (`SUM`, `AVERAGE`, `COUNT`, `MAX`, …): anywhere else it is `#VALUE!` |
| Range between computed references | `A1:INDEX(A1:A5,3)`, `INDEX(…):INDEX(…)`, `OFFSET(A1,1,0):C5`, `INDIRECT("A1"):INDIRECT("B3")` | The bounding box of both ends; both must be on the same sheet (else `#VALUE!`) |
| Union | `(A1:A2,C1)` | A `,` inside parentheses; valid as a function argument (`SUM((A1:A2,C1))`), otherwise `#VALUE!` |
| Intersection | `A1:A3 A2:A4` | The overlap; a single cell collapses to a value, no overlap is `#NULL!` |
| Implicit intersection | `@A1:A3` | The cell of the range in the formula's own row (vertical range) or column (horizontal range), both for a 2-D range; `#VALUE!` when the formula lies outside the range; the top-left value of a computed array |
| Spill reference | `A1#`, `Sheet2!B2#` | The whole spill range anchored at that cell (`SUM(A1#)`, `ROWS(A1#)`); `#REF!` when the cell anchors no spill |
| Array constant | `{1,2,3}`, `{1,2;3,4}`, `{"a",TRUE,#N/A,-1.5}` | `,` separates columns and `;` rows; numbers (optionally signed), text, `TRUE`/`FALSE` and errors only; every row must have the same length, otherwise a parse error (`#ERROR! Array constant at position 0 is ragged: row 2 has 1 elements, row 1 has 2`) |
| Error literal | `#NULL!` `#DIV/0!` `#VALUE!` `#REF!` `#NAME?` `#NUM!` `#N/A` `#SPILL!` `#CALC!` | Evaluates to that error; `ISERROR`, `IFERROR`, `ISNA`, `ERROR.TYPE` see it as one |
| Defined name | `Quarter1` | Not resolved yet: evaluates to `#NAME?` (the text is kept as typed, so a save does not corrupt it) |

An unquoted sheet name is read as words of letters, digits, `_` and `.` joined by
single spaces (`Q1 Data!A1`, `My Sheet!A1`, `2024!A1`), as earlier versions did, but
only where an operand starts (not after an operand, `:` or a sheet prefix) and never
inside a span; Excel itself always quotes these names, and the engine prints them
quoted (`'Q1 Data'!A1`). `A1 Sheet2!B2` at the start of an operand is therefore the
sheet "A1 Sheet2", and a span of such sheets needs quotes (`'Q1 Data:Q2 Data'!A1`).
Array-constant errors are tagged in the AST, so `{#N/A}` (an error) and `{"#N/A"}`
(text) stay different through printing.

A formula the engine cannot parse (an imported `Sales[Q1]`) keeps the value the file
saved and its text through recalculation and edits elsewhere; on export that value is
written back, and the `#ERROR! ...` of a formula typed in that does not parse is never
written as a result. A formula whose tree is deeper than 500 levels (a flat chain of
thousands of `+`, `%` or spaces) is rejected as `#ERROR! Formula is too complex`.

Malformed input is rejected with a message, never read as text: `A:` (a
half-open range), `1:A`, `Sheet1:!A1`, `{1,2;3}` (ragged array), `{}`,
`{A1}` (an array constant holds no references), `SUM(1`, `1+`, `(A1,2)` (a union
needs references), `2 A1` (an intersection needs references), `Sales[Q1]`
(table references are not supported yet), and nesting deeper than 256 levels.

### Whole-column and whole-row references

`A:A`, `A:C`, `1:1` and `1:3` are stored **unbounded** in the AST: a
whole-column range holds the ids of its columns and the sentinel id `*` for the
row axis, a whole-row range the ids of its rows and `*` for the column axis. For
evaluation the unbounded axis is **bounded to the sheet's used range**: the last
row (or column) anywhere on the sheet that holds a value, a formula or a spilled
value. Every whole column on a sheet therefore has the same height, so
`SUMPRODUCT(A:A*B:B)` and `SUM(A:A+B:B)` line up whatever the columns hold, and an
empty column gives blanks, not an error. The bounded axis keeps its nominal size
(`A:C` is three columns), except that a span of more than 1,024 columns or rows
(`A:XFD`, `1:1048576`) is trimmed to the used range as well, so such a range costs
what the data costs. One known difference from Excel: `COUNTIF(A:A,"")` counts the
blanks inside the used range only (Excel counts the whole column). Shape functions use the
nominal size: `ROWS(A:A)` is 1,048,576, `COLUMNS(1:1)` is 16,384, `COUNTBLANK`
counts blanks out of the nominal size, and `INDEX(A:A,500)` indexes the full
column (a blank, not `#REF!`, past the data).

**Dependencies.** A whole-column range is one rectangle bounded on one axis
(`RangeDependency.axis`); it is never expanded into cells. An edit anywhere in
the referenced columns, and only there, dirties the formula. `SUM(A:A)` over a
column of a million rows evaluates in about two seconds on a laptop (one pass
over the stored cells), and an edit recomputes only the formulas whose columns
it touches.

### What a structural edit does to each form

| Edit | `A:A` / `A:C` | `1:1` / `1:3` | `A1:B2`, `A1`, `@`, `A1#`, `:` over refs | 3D (`Sheet1:Sheet3!A1`) |
|---|---|---|---|---|
| Insert rows | unchanged (text and value stay valid; new rows count) | follows its rows (`1:1` becomes `3:3` after two rows are inserted above); grows when rows are inserted inside `1:3` | follow their ids as before | not moved: it names positions |
| Delete rows | unchanged | `#REF!` when a row that is a corner of the range is deleted (printed `#REF!`) | as before: a deleted corner row gives `#REF!` | not moved |
| Insert columns | follows its columns (`A:A` becomes `C:C`); grows when inserted inside `A:C` | unchanged | follow their ids | not moved |
| Delete columns | `#REF!` when a corner column is deleted | unchanged | as before | not moved |
| Copy / fill | relative columns move (`A:A` one column right is `B:B`), `$A:$A` stays, rows never change | relative rows move (`1:1` one row down is `2:2`), `$1:$1` stays | as before (`@` and `#` operands move with the cell) | relative positions move with the copy; sheets stay |
| Rename a sheet | the prefix follows | the prefix follows | the prefix follows | both end names follow; the span is quoted when needed (`'First one:Last'!A1`) |
| Delete a sheet | `#REF!` if it is the referenced sheet | `#REF!` | `#REF!` | an end sheet moves **inward** (Excel's rule); deleting a sheet inside the span just drops it from the total; a span left with only the deleted sheet is `#REF!` |
| Reorder sheets | | | | the span is read from the live order |
| Add a sheet | | | | included only if it lands between the end sheets (a new sheet is appended after the last) |

Deleting a *corner* of a multi-row or multi-column range gives `#REF!` for every
range form, plain or whole (Excel shrinks the range instead): this is the
engine's existing rule for ranges and it is applied unchanged to `A:C` and
`1:3`. The whole-column and whole-row forms are valid again as soon as the
formula is re-entered.

A 3D reference is **positional**: it stores the cell positions (not ids) and the
two end sheet names, so inserting or deleting rows on one sheet of the span does
not move it, and its span follows the workbook sheet order. Its formula is
**volatile** (recomputed on every pass), as is any formula with a `:` between
computed references (`A1:INDEX(…)`), because the cells either depend on are only
known when the formula is evaluated. A formula with a deleted reference prints
it as the single token `#REF!` (`=SUM(#REF!)`), which parses back.

## Dependency graph and recalculation

`FormulaGraphImpl` (`formula-graph.ts`) tracks which formulas depend on
which cells:

- **`cellDependents`** — a reverse index from *any* referenced cell key
  (a value cell or a formula cell) to the formulas that read it. Because it
  covers value cells too, changing a plain value correctly recalculates the
  formulas above it.
- **Range dependencies** are stored as **rectangles** of two stable corner
  keys, not expanded cell-by-cell (a whole-column or whole-row range is a
  rectangle bounded on one axis). Containment is tested at recalc time, so
  a cell that is empty when the formula is entered — or a row/column
  inserted into the range later — is still tracked.

On a change, `collectDirty` walks the transitive dependents (plus every
volatile cell) and records the dependency edges among them.
`topologicalOrder` (Kahn's algorithm) then orders that set so each cell is
recomputed **exactly once**, dependencies before dependents; any cell in, or
downstream of, a cycle: the cells on the cycle evaluate to `0`, as in Excel, and the workbook emits a `circularReference` event listing the cells on the cycle. `workbook.setCircularReferenceMode('ref')` makes them `#REF!` instead; `'error'` shows the diagnostic `#CIRCULAR!` token. `#CIRCULAR!` is never written to XLSX.

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

## Conformance suite

`packages/sheets-core/src/formula-conformance-oracle.json` holds one row per
formula (the 159-formula probe of `fixtures/xlsx/formulas-core.xlsx` plus the
batch-A cases of `fixtures/xlsx/generators/conformance_cases.py`) with the value
LibreOffice 25.2 computes when it recalculates the authored workbook.
`python3 -I fixtures/xlsx/generators/build_conformance_oracle.py` regenerates it
(run locally; CI never runs LibreOffice). Where LibreOffice differs from Excel
(for example it reports an invalid argument as `#VALUE!` where Excel gives
`#NUM!`) the row carries the documented Excel value, tagged `excel: true,
unverified: true`, with LibreOffice's value and the reason in `note`.
`formula-conformance-oracle.test.ts` runs every row through the engine (numbers
to a relative 1e-9, text exactly, errors by code); rows the engine cannot pass yet
sit in an explicit `KNOWN_FAILING` list with a reason, and the suite asserts that
they still fail, so the list shrinks as functions land.

## Function Reference

All 193 built-in functions, grouped by category:

### Math (67 functions)

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
| `SIN(number)` | Returns the sine of an angle in radians. |
| `COS(number)` | Returns the cosine of an angle in radians. |
| `TAN(number)` | Returns the tangent of an angle in radians. |
| `ASIN(number)` | Returns the arcsine of a number, in radians (-1 to 1). |
| `ACOS(number)` | Returns the arccosine of a number, in radians (-1 to 1). |
| `ATAN(number)` | Returns the arctangent of a number, in radians. |
| `ATAN2(x_num, y_num)` | Returns the arctangent of the point (x, y), in radians. |
| `SINH(number)` | Returns the hyperbolic sine of a number. |
| `COSH(number)` | Returns the hyperbolic cosine of a number. |
| `TANH(number)` | Returns the hyperbolic tangent of a number. |
| `ASINH(number)` | Returns the inverse hyperbolic sine of a number. |
| `ACOSH(number)` | Returns the inverse hyperbolic cosine of a number (at least 1). |
| `ATANH(number)` | Returns the inverse hyperbolic tangent of a number (between -1 and 1). |
| `PI()` | Returns the value of pi. |
| `RADIANS(angle)` | Converts degrees to radians. |
| `DEGREES(angle)` | Converts radians to degrees. |
| `FACT(number)` | Returns the factorial of a number. |
| `FACTDOUBLE(number)` | Returns the double factorial of a number. |
| `COMBIN(number, number_chosen)` | Returns the number of combinations for a number of items. |
| `PERMUT(number, number_chosen)` | Returns the number of permutations for a number of items. |
| `SUMSQ(number1, [number2, ...])` | Returns the sum of the squares of the arguments. |
| `EVEN(number)` | Rounds a number away from zero to the nearest even integer. |
| `ODD(number)` | Rounds a number away from zero to the nearest odd integer. |
| `ISEVEN(number)` | Returns TRUE when a number is even. |
| `ISODD(number)` | Returns TRUE when a number is odd. |
| `CEILING.MATH(number, [significance], [mode])` | Rounds a number up to the nearest multiple of significance. |
| `FLOOR.MATH(number, [significance], [mode])` | Rounds a number down to the nearest multiple of significance. |
| `MAXIFS(max_range, criteria_range1, criterion1, [criteria_range, criterion, ...])` | Returns the largest value among the cells that meet multiple criteria. |
| `MINIFS(min_range, criteria_range1, criterion1, [criteria_range, criterion, ...])` | Returns the smallest value among the cells that meet multiple criteria. |
| `AGGREGATE(function_num, options, ref1, [ref2 or k, ...])` | Applies one of 19 functions to a list, optionally ignoring errors. |

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

### Statistical (30 functions)

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
| `STDEV.S(number1, [number2, ...])` | Estimates standard deviation from a sample. |
| `STDEV.P(number1, [number2, ...])` | Calculates standard deviation of an entire population. |
| `VAR.S(number1, [number2, ...])` | Estimates variance from a sample. |
| `VAR.P(number1, [number2, ...])` | Calculates variance of an entire population. |
| `PERCENTILE.INC(array, k)` | Returns the k-th percentile of values (k from 0 to 1, inclusive). |
| `PERCENTILE.EXC(array, k)` | Returns the k-th percentile of values (k strictly between 0 and 1). |
| `QUARTILE.INC(array, quart)` | Returns a quartile (0-4) of a data set, inclusive. |
| `QUARTILE.EXC(array, quart)` | Returns a quartile (1-3) of a data set, exclusive. |
| `RANK.EQ(number, ref, [order])` | Returns the rank of a number within a list; ties share the top rank. |
| `MODE.SNGL(number1, [number2, ...])` | Returns the most frequently occurring value. |
| `SLOPE(known_ys, known_xs)` | Returns the slope of the linear regression line. |
| `INTERCEPT(known_ys, known_xs)` | Returns the intercept of the linear regression line. |
| `RSQ(known_ys, known_xs)` | Returns the square of the Pearson correlation coefficient. |
| `FORECAST(x, known_ys, known_xs)` | Predicts a value along a linear trend. |
| `FORECAST.LINEAR(x, known_ys, known_xs)` | Predicts a value along a linear trend. |
| `GEOMEAN(number1, [number2, ...])` | Returns the geometric mean of positive numbers. |
| `FREQUENCY(data_array, bins_array)` | Counts values per interval, as a column (spills). |

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
