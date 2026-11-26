# Spindle Sheets — remaining work

Gaps between the current spreadsheet and the feature set users expect from
Excel / Google Sheets. **Formula coverage is essentially complete** — the bulk
of the remaining work is **feature parity** in the non-formula feature set.

## Formula engine — shipped

The engine ships **193 functions** across math, logical, text, lookup,
statistical, date/time, information, financial, reference and array categories
(`packages/sheets-core/src/formula-parser/functions/`).

- [x] **Standard function library** — all categories above (193 in `catalog.ts`)
- [x] **Excel-grade recalculation** — topological dirty-propagation (each cell
      computed once per pass), volatile functions, complete range-rectangle
      dependency tracking (`formula-graph.ts`, `workbook.ts`)
- [x] **Dynamic arrays / spill** — `SEQUENCE`, `UNIQUE`, `SORT`, `SORTBY`,
      `FILTER`, `SPLIT`; a spilling formula fills a block, with `#SPILL!` on a
      blocked target (`spill.ts`)
- [x] **Formula autocomplete + parameter help** while editing a formula
- [x] **Reference syntax** — whole-column and whole-row references (`A:A`, `1:3`,
      bounded to the used range, one rectangle in the dependency graph), postfix
      `%`, array constants (`{1,2;3,4}`), 3D references (`Sheet1:Sheet3!A1`),
      ranges ending in `INDEX` / `OFFSET` / `INDIRECT`, union `(A1,B1)`,
      intersection (space), `@`, the spill operator `A1#` and the `#NULL!` literal
      (`formula-parser/tokenizer.ts`, `syntax.ts`; semantics in
      [`core/formulas.md`](core/formulas.md#formula-syntax)). A formula that does
      not parse now shows `#ERROR! <message>` instead of being read as text.

### Remaining formula items

- [ ] **Defined names and table references** — `Quarter1` evaluates to `#NAME?` and
      `Sales[Q1]` does not parse (the text of an imported formula is kept; plan S8).
- [ ] **Shrinking ranges on a corner delete** — deleting the first or last row or
      column of a range (`A1:A5`, and now `A:C` / `1:3` too) gives `#REF!`; Excel
      shrinks the range. 3D references are positional and do not follow row or
      column edits on their sheets.

- [ ] **Hidden rows and nested subtotals in `AGGREGATE` / `SUBTOTAL`** — `AGGREGATE`
      (all 19 functions, both forms) skips errors for options 2, 3, 6 and 7, but
      options 1, 3, 5 and 7 (ignore hidden rows, nested `SUBTOTAL`/`AGGREGATE`) and
      `SUBTOTAL` 101-111 act like their visible-everything counterparts: the
      evaluation context does not expose row visibility or which cells hold
      subtotals (`functions/aggregate.ts`).
- [ ] **Range versus direct arguments in the statistics functions** — eager
      functions receive values, not references, so `strictNumbers` treats a range
      and a typed argument alike: numeric text inside a range counts as a number
      (Excel ignores it), and a typed `"a"` or `TRUE` is dropped (Excel gives
      `#VALUE!` and 1). Affects the whole SUM / STDEV / SUMSQ / GEOMEAN /
      FREQUENCY family, not just batch A; fixing it needs argument provenance in
      the evaluator.
- [ ] **`ARRAYFORMULA`** — not implemented at all (no catalog entry, zero
      references); needs full elementwise scalar broadcast.
- [ ] **Custom date codes in the formula bar and editor** — an imported date whose code has no structured type (`d-mmm-yy`, `dddd, mmmm d, yyyy`) is a
      `custom` format; the formula bar and cell editor show its serial number and no date picker opens (the `date` type shows a date and a picker).
      Autofill already treats it as a date.
- [ ] **Number formats: locale and the reader's limits** — month and day names are English whatever `[$-LCID]` says, separators are `.` and `,`,
      `*x` fill is not repeated to the cell width, non-Gregorian calendars and `[DBNum1]` are invalid. The XLSX reader library drops the backslashes
      of a code and returns no code for some built-in ids (see `documentation/sheets/core/number-formats.md`); both go away with plan S2.
- [x] **`TEXT()` date format codes** — `TEXT()` uses the Excel number-format
      engine (`number-format/`): date and time codes, sections, conditions,
      scaling commas, fractions. Remaining gaps are listed in
      `documentation/sheets/core/number-formats.md` (en-US names only, no
      locale separators, no non-Gregorian calendars, `*x` fill not repeated).
- [ ] **`#SPILL!` auto-recovery** — a blocked anchor does not re-spill when the
      blocking cell is cleared; the anchor isn't a dependency of the blocker, so
      it's never marked dirty. The formula must be re-entered (`workbook.ts`).
- [ ] **Missing functions** — `TRANSPOSE` (a `transpose` *helper* exists but isn't
      exported as a function), `YEARFRAC`, `TEXTSPLIT`, `TEXTBEFORE` / `TEXTAFTER`,
      `CLEAN`, `NUMBERVALUE`, `DAYS`, `ISOWEEKNUM`, `XNPV`, `XIRR`, `SLN` (batch B,
      plan S12), and `LET` / `LAMBDA` / `HSTACK` / `VSTACK` / `TAKE` (post-launch).
      Also `INDEX(range, 0, col)` / `INDEX(range, row, 0)` (whole-column /
      whole-row array form) returns `#REF!` instead of the column/row
      (`functions/lookup.ts`). The unprefixed names work; the `_xlfn.` spellings
      Excel writes into files are mapped on import (plan S4).
- [x] **Math, trigonometry and statistics, batch A** — `SIN` to `ATANH`, `PI`,
      `RADIANS`, `DEGREES`, `FACT`, `FACTDOUBLE`, `COMBIN`, `PERMUT`, `SUMSQ`,
      `EVEN`, `ODD`, `ISEVEN`, `ISODD`, `CEILING.MATH`, `FLOOR.MATH`, `MAXIFS`,
      `MINIFS`, `AGGREGATE`, the dotted statistics names (`STDEV.S`, `VAR.P`,
      `PERCENTILE.INC` / `.EXC`, `QUARTILE.INC` / `.EXC`, `RANK.EQ`, `MODE.SNGL`),
      `SLOPE`, `INTERCEPT`, `RSQ`, `FORECAST` / `FORECAST.LINEAR`, `GEOMEAN` and
      `FREQUENCY` (plan S11). The conformance suite
      (`formula-conformance-oracle.json`) checks them against LibreOffice values.

## Feature parity

Gaps against Excel / Google Sheets in the non-formula feature set.

### Fully absent (no engine, no UI)

- [ ] **Conditional formatting** — value- and formula-driven cell styling
- [ ] **Data validation** — dropdown lists, input restrictions
- [ ] **Named ranges** — and their use in formulas
- [ ] **Charts** — bar / line / pie, etc.
- [ ] **Pivot tables**
- [ ] **Sparklines**
- [ ] **Cell & sheet protection** — locked cells, protected ranges
- [ ] **Row / column grouping & outline**
- [ ] **Paste special** — values-only, formats-only, transpose (the copy/paste
      path is plain TSV only; `TRANSPOSE` exists as a function, not a paste mode)
- [ ] **XLSX fidelity gaps** — XLSX import/export shipped in 0.4.0 (implementation now lives in
      `@b32nio/spindle-io`, where fidelity work lands)
      (`exportToXlsx` / `importFromXlsx`, plus Open/Download in the toolbar), but
      charts, conditional formatting and data validation are not carried over. A corpus of test workbooks
      with a construct manifest (`fixtures/xlsx/`) now defines what "carried over" has to cover, and the full
      loss report (`inspectXlsx(bytes, { detail: 'full' })`, ids in `documentation/import-loss.md`) lists what each
      file holds that is still not carried over (tab colour, zoom, print setup, theme colours, array formulas...).
- [ ] **Print / page setup** for sheets (PDF output)

### Engine ready, UI missing

- [ ] **Multi-column sort UI** — the engine accepts a `SortOrder[]` and has the
      multi-column toggle-cycle logic, but the React layer only ever builds a
      single-key sort (`WorkbookCanvas.tsx`); the multi-column path is never
      exposed.

## Already shipped

Real-time collaboration, comments and comment threads, freeze panes,
single-column sort, per-column filters, cell styles (fonts, colors, borders,
alignment, wrap, rotation), number formats, hyperlinks, date-entry
auto-detection, hide/show and insert/delete rows and columns, merged cells,
multiple sheets, undo/redo, find & replace, autofill (numbers, dates,
month/weekday name lists and text patterns), **CSV and XLSX import/export** (with Open/Download in the toolbar), the
193-function formula engine with dynamic arrays, and formula autocomplete with
parameter help.
