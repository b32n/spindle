# Spindle Sheets — remaining work

Gaps between the current spreadsheet and the feature set users expect from
Excel / Google Sheets. **Formula coverage is essentially complete** — the bulk
of the remaining work is **feature parity** in the non-formula feature set.

## Formula engine — shipped

The engine ships **229 functions** across math, logical, text, lookup,
statistical, date/time, information, financial, reference and array categories
(`packages/sheets-core/src/formula-parser/functions/`).

- [x] **Standard function library** — all categories above (229 in `catalog.ts`)
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
- [ ] **3D references and structural edits** — a 3D reference (`Sheet1:Sheet3!A1`) is
      positional and does not follow row or column edits on its sheets (normal and
      cached-only formulas alike); sheet names inside it are matched exactly, not
      case-insensitively (a 3D span and its sheet-delete shrinking).

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
      FREQUENCY family and `MIRR` (numeric text inside a reference counts as a flow), not just batch A; fixing it needs argument provenance in
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
- [ ] **`TEXTBEFORE`, `TEXTAFTER`, `TEXTSPLIT`** — deferred until Excel-computed files exist (plan S12, D8): LibreOffice 25.2 has no
      such function, so there is no trustworthy oracle. Their probe rows are in `KNOWN_FAILING`. (`omittable.ts` is where
      `TEXTSPLIT(a,,";")` would need its blank argument kept.)
- [ ] **Missing functions** — `HSTACK`, `VSTACK`, `TAKE`, `DROP`, `CHOOSECOLS`, `CHOOSEROWS`, `EXPAND`, `TOCOL`, `TOROW`, `WRAPROWS`,
      `WRAPCOLS`, `ARRAYTOTEXT`, `VALUETOTEXT` (no LibreOffice oracle and outside plan S12: S13 / post-launch); `LET` and `LAMBDA`
      with its helpers need names in scope in the evaluator (post-launch). The unprefixed names work; the `_xlfn.` spellings
      Excel writes into files are mapped on import and export (plan S4), and a formula the engine cannot evaluate keeps its saved
      result (`cachedOnly`) until its function exists.
- [x] **Text and date, batch B** — `CLEAN`, `NUMBERVALUE`, `UNICHAR`, `UNICODE`, `T`,
      `FIXED`, `DOLLAR`, `BASE`, `DECIMAL`, `ROMAN`, `ARABIC`, `ISNONTEXT`, `DAYS`, `DAYS360`,
      `YEARFRAC` (all five bases), `ISOWEEKNUM`, `WORKDAY.INTL`, `NETWORKDAYS.INTL`, and `INDEX(range,0,col)` /
      `INDEX(range,row,0)` (whole column or row, plan S12). Rows where LibreOffice differs from Excel are
      Excel-documented and unverified until real Excel files confirm them (D8). Date functions use Excel's
      1900 calendar (29 February 1900 is serial 60; `YEARFRAC` basis 1 on 1900 dates is unverified).
- [ ] **Excel's `DAYS360` at the end of February** — the US method treats the last day of February at both ends as the 30th
      (equal dates are 0 days apart); LibreOffice says -2 for 28 February to 28 February. The two rows are Excel-documented
      and unverified (`functions/date-ext.ts`).
- [x] **Array, finance and math, batch B** — `TRANSPOSE`, `XNPV`, `XIRR` (Newton with a bracketed fallback: a deliberate superset
      of Excel, which gives up in some cases), `MIRR`, `SLN`, `SYD`, `DDB`, `EFFECT`, `NOMINAL`, `RRI`, `PDURATION`, `ISPMT`,
      `FVSCHEDULE`, `QUOTIENT`, `SQRTPI`, `SUMX2MY2`, `SUMX2PY2` and `SUMXMY2` (plan S12). Where LibreOffice reports `#VALUE!` for an
      argument out of range (most finance functions) the rows carry Excel's `#NUM!` / `#DIV/0!`, tagged unverified. In the
      `formulas-core` fixture each dynamic-array probe has a column of its own, so its spill is not blocked (SA-gap).
- [x] **Math, trigonometry and statistics, batch A** — `SIN` to `ATANH`, `PI`,
      `RADIANS`, `DEGREES`, `FACT`, `FACTDOUBLE`, `COMBIN`, `PERMUT`, `SUMSQ`,
      `EVEN`, `ODD`, `ISEVEN`, `ISODD`, `CEILING.MATH`, `FLOOR.MATH`, `MAXIFS`,
      `MINIFS`, `AGGREGATE`, the dotted statistics names (`STDEV.S`, `VAR.P`,
      `PERCENTILE.INC` / `.EXC`, `QUARTILE.INC` / `.EXC`, `RANK.EQ`, `MODE.SNGL`),
      `SLOPE`, `INTERCEPT`, `RSQ`, `FORECAST` / `FORECAST.LINEAR`, `GEOMEAN` and
      `FREQUENCY` (plan S11). The conformance suite
      (`formula-conformance-oracle.json`) checks them against LibreOffice values.

## Known limitations — import

- [ ] **Loading a formula-heavy file into the editor is slow** (found in the S2b review; engine code, outside the reader, so not fixed there;
      candidate owner S20, the scale and worker work). `importFromXlsx` reads a file in 0.1 to 0.2 s in these shapes, then `setData`'s
      recalculation takes: 20,000 array formulas, each `<f t="array" ref="A{i}:Z{i+999}">` (a 276 KB zip), 110 s (5,000 took 5.8 s: quadratic;
      the time is in `getCell` and `collectDependencies`); 1,000,000 ordinary formulas (a 7 MB zip) 82 s and 2.1 GB of process memory; 1,000
      sheets each with one formula naming the next sheet's cell 33 s (400 took 4.6 s; `applyOperator` and `getSheetIdByName`). Until then a
      host should lower `maxTotalCells` and `maxSheets` and/or import in a worker (`documentation/import-loss.md`). Related, not a defect: a table of
      two million shared strings peaks at 2.8 GB of process memory in the reader (five million took 1.17 GB).

- [ ] **Blank source cells in dynamic-array results** (found in the S4 review; unowned: S5 and S11 are merged and no later plan entry covers it, so it is a known wrong value that needs a new follow-up entry). The engine spills a blank source cell as blank where Excel
      shows 0. Repro: A1 = 1, A2 empty, A3 = 3, `B1 = A1:A3` spills `1, (blank), 3` (Excel: `1, 0, 3`), and `D1 = B2&"x"` gives `"x"` (Excel: `"0x"`).
      A legacy (CSE) array formula read from a file already reads 0 (`fitToExtent`). Fix: coerce a null member of a plain range-reference spill to 0 in the
      spill path, without changing `=A1:A3*1`-style results; it is an engine semantics change, so it is not part of the interchange PR.

- [ ] **Implicit intersection of ranges inside calls in legacy formulas** (found in the SA-gap review; unowned: S5 and S11 are merged and no later plan entry covers it, so it is a known wrong value that needs a new follow-up entry). A formula saved without the array mark gets Excel's
      intersection only for range operands outside every call and group (`=A1:A3*2`, shown `=@A1:A3*2`). Excel also intersects a range passed to a scalar parameter or used
      in the arithmetic of a call's argument; the engine has no table of which parameters are scalar (the function catalog has names only), so it can return a wrong value.
      Repros (data A1:A3 = 1, 2, 3; the formula is in a file without `cm` and `t="array"`): `F5 = LEN(A1:A3)` gives 5 (Excel `#VALUE!`; in F2 Excel 1),
      `F5 = IF(A1:A3>1,1,0)` gives 1 (Excel `#VALUE!`; F1 gives 0), `F2 = SUM(A1:A3*2)` gives 12 (Excel 4), `F2 = ABS(A1:A3)` gives `#VALUE!` (Excel 2). Fix: a
      scalar/array flag per parameter in the catalog, checked against Excel's documentation, then `@` before range operands in scalar slots on import (nested
      comparison and arithmetic included) and on export the same bare-range rule only at the top level. The repros are `known divergence` tests in
      `xlsx-implicit-intersection.test.ts`. Named ranges and structured references are not intersected either.

- [x] **Textual reference shifting for cached-only formulas** (found in the S4 review; done in S4b). Inserting or deleting rows or columns now rewrites
      the references in a cached-only formula's text (`formula-parser/ref-shift.ts`, on the lenient tokenizer), by the same rule as a normal formula.
      Repro that failed: `B1 = _xlfn.LET(_xlpm.x,A1,_xlpm.x*2)`, insert a row above row 1: B2 now reads `A2`. The review of S4b also fixed the engine's
      own stable-id path: a range with a deleted boundary was `#REF!` whole (while an interior delete shrank it); both paths now share `shiftSpan`
      (Excel: the surviving rows/columns stay, `#REF!` only when all of the range is deleted), and sheet names are case-insensitive in references
      (`sheet2!A1` is `Sheet2`). Open items: the engine has no cell-move operation (copy/fill/paste rebases through the parser and drops the
      cached-only mark); two collaborators editing the structure at once: structure sync is last writer wins and the text follows the edit that
      survives; named ranges, conditional formats and data validation hold no sheet references in the engine today, so nothing there follows a delete.

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
      An unsafe file (encrypted, legacy `.xls`, a zip bomb, over the size limits) is refused with a named `ImportError`
      (`documentation/import-loss.md`, "Files that cannot be opened at all"); the DOCX reader does not use these limits yet.
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
229-function formula engine with dynamic arrays, and formula autocomplete with
parameter help.
