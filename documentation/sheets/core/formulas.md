# Formula System

The formula system gives `@b32nio/spindle-sheets-core` spreadsheet calculation:
a parser, a library of **242 built-in functions** (the [published list](functions.md)), a dependency graph with
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
| `math-ext.ts` | trigonometry, factorials, parity, `CEILING.MATH` / `FLOOR.MATH` (batch A), `QUOTIENT`, `SQRTPI`, `SUMX2MY2` family (batch B) | 32 |
| `aggregate.ts` | `MAXIFS`, `MINIFS`, `AGGREGATE` (batch A) | 3 |
| `logical.ts` | logical | 11 |
| `text.ts` | text | 21 |
| `text-ext.ts` | `CLEAN`, `NUMBERVALUE`, `FIXED`, `DOLLAR`, `BASE`, `ROMAN`, `ARRAYTOTEXT`, `VALUETOTEXT`, ... (batch B) | 13 |
| `lookup.ts` | lookup & reference | 12 |
| `stats.ts` | statistical | 13 |
| `stats-ext.ts` | dotted statistics names, regression, `GEOMEAN`, `FREQUENCY` (batch A) | 17 |
| `date.ts` | date & time | 19 |
| `date-ext.ts` | `DAYS`, `DAYS360`, `YEARFRAC`, `ISOWEEKNUM`, `WORKDAY.INTL`, `NETWORKDAYS.INTL` (batch B) | 6 |
| `information.ts` | information | 12 |
| `financial.ts` | financial | 9 |
| `financial-ext.ts` | `XNPV`, `XIRR`, `MIRR`, `SLN`, `SYD`, `DDB`, `EFFECT`, `NOMINAL`, `RRI`, `PDURATION`, `ISPMT`, `FVSCHEDULE` (batch B) | 12 |
| `reference.ts` | reference-aware | 11 |
| `array.ts` | dynamic arrays | 6 |
| `array-ext.ts` | `TRANSPOSE`, `CHOOSECOLS` / `CHOOSEROWS`, `TAKE` / `DROP`, `HSTACK` / `VSTACK`, `EXPAND`, `TOCOL` / `TOROW`, `WRAPROWS` / `WRAPCOLS` (batch B) | 12 |

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
batch-A cases of `fixtures/xlsx/generators/conformance_cases.py` and the batch-B cases of `conformance_cases_b.py`) with the value
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

LibreOffice 25.2 has no `ARRAYTOTEXT`, `VALUETOTEXT` and no array shaper but `TRANSPOSE` (`TAKE`, `DROP`, `HSTACK`, `VSTACK`, `CHOOSECOLS`, `CHOOSEROWS`, `EXPAND`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`; it answers `#NAME?`), so their rows carry
the behaviour Excel's documentation gives, tagged `excel: true, unverified: true`. Where LibreOffice differs from Excel the row carries Excel's documented value, tagged `excel: true, unverified: true`, until real Excel files (plan D8) confirm it.
`TEXTBEFORE`, `TEXTAFTER` and `TEXTSPLIT` are deferred until such files exist (plan S12 does them only with a trustworthy oracle: LibreOffice 25.2 has none); their
probe rows are in `KNOWN_FAILING`. A row tagged `unverified` without `excel` keeps LibreOffice's value but flags that Excel's behaviour is not known to agree.

### Omitted arguments

The parser skips an empty argument (`F(1,,2)`) for most functions. For the ones where Excel gives an omitted argument its default and
the position matters (`INDEX(A1:C3,,2)`, `TAKE(A1:C5,,2)`, `NUMBERVALUE(a,,",")`, `WORKDAY.INTL(a,b,,holidays)`, `XIRR(v,d,)`), the names in
`functions/omittable.ts` keep a placeholder node (a `number` node whose value is `undefined`) so that the later arguments keep their
places; the function sees `undefined`, as for an argument left off the end, and the stable AST writes it back as nothing. A leading `_xlfn.` is
ignored when matching the name.

### `INDEX` with a row or column of 0

`INDEX(range, 0, col)` and `INDEX(range, row, 0)` (or with the argument left blank, `INDEX(range,,col)`) select a whole column or row, as
an array that spills and feeds aggregates (`SUM(INDEX(A1:C3,0,2))`); a single index into a table (`INDEX(A1:C3,2)`) selects the whole row.
`functions/lookup.ts` `indexWindow` does the arithmetic for both the value form and the reference form, so these are also range
endpoints (`INDEX(A1:C3,0,2):INDEX(A1:C3,0,3)`). A negative index is `#VALUE!`, one past the range `#REF!`; an omitted `area_num` is 1.

### Array and finance, batch B

`array-ext.ts` returns 2D arrays for the spill engine. An empty cell that passes through (`TRANSPOSE(A1:A3)` over a blank `A2`, `HSTACK`, `TOCOL`
without the blank-skipping mode) becomes 0, as in Excel; `HSTACK`, `VSTACK`, `EXPAND` and `WRAPROWS`/`WRAPCOLS` pad with `#N/A` unless told otherwise;
an empty result (`TAKE(a,0)`, `DROP` of everything, `TOCOL` of nothing) is `#CALC!`. Results are limited to a million cells (`#NUM!`). Only `TRANSPOSE` has a
LibreOffice oracle. In the
`formulas-core` fixture `=TRANSPOSE(A1:B2)` evaluates, but its spill range is blocked by the next probe row's formula (`#SPILL!`), so that cell stays a
mismatch.

`XIRR` runs Newton's method from the guess (default 10%); when that fails to converge, leaves the domain or stops on a flat tail
(a root a hair above -100% is `#NUM!`), it scans rates `e^s - 1` from -100% to 10^10 for a sign change (compared as signs) and bisects the
bracket nearest the guess. This is a deliberate superset of Excel, whose Newton iteration gives up with `#NUM!` (or a near-zero rate) in some
cases where a root exists; with several roots the answer depends on the guess, as in Excel. `XNPV` and `XIRR` need dates that are not negative
and do not precede the first one and, for `XIRR`, a positive and a negative cash flow (`#NUM!` otherwise). `DDB` uses the closed form
`max(cost*(1-factor/life)^k, salvage)` per period, which matches Excel's period loop.

### The 1900 calendar

Date functions follow Excel's 1900 date system, in which 29 February 1900 exists (serial 60), 1900-01-01 is serial 1 and serial 0 is
"1900-01-00": `DATE` (including overflow, `DATE(1900,2,30)` is 61 and `DATE(1900,3,0)` is 60), `EDATE`, `EOMONTH`, `DATEDIF`, `YEAR`, `MONTH`, `DAY` (`DAY(60)` is 29,
`DAY(0)` is 0) and the batch-B date functions read and write serials through `number-format/dates.ts`, so
`DAYS(DATE(1900,3,1),DATE(1900,2,28))` is 2 as in Excel (LibreOffice has no such day and says 1; those rows are tagged in the oracle).
`YEARFRAC` basis 1 counts 1900 as a leap year for consistency with that calendar; Excel's own behaviour there is unverified (tagged row).

## Function Reference

The full list, by category, with each function's signature and the evidence for its Excel parity, is [functions.md](functions.md).
