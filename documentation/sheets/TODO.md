# Spindle Sheets — remaining work

Gaps between the current spreadsheet and the feature set users expect from
Excel / Google Sheets. **Formula coverage is essentially complete** — the bulk
of the remaining work is **feature parity** in the non-formula feature set.

## Formula engine — shipped

The engine ships **146 functions** across math, logical, text, lookup,
statistical, date/time, information, financial, reference and array categories
(`packages/sheets-core/src/formula-parser/functions/`).

- [x] **Standard function library** — all categories above (146 in `catalog.ts`)
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

- [ ] **`AGGREGATE`** — 19 sub-functions plus the ignore-hidden / ignore-errors
      option matrix (absent from the catalog).
- [ ] **`ARRAYFORMULA`** — not implemented at all (no catalog entry, zero
      references); needs full elementwise scalar broadcast.
- [ ] **`TEXT()` date format codes** — `TEXT()` formats numbers only (decimals,
      grouping, percent, scientific, currency); date/time patterns are not
      supported (`functions/text.ts`).
- [ ] **`#SPILL!` auto-recovery** — a blocked anchor does not re-spill when the
      blocking cell is cleared; the anchor isn't a dependency of the blocker, so
      it's never marked dirty. The formula must be re-entered (`workbook.ts`).
- [ ] **Missing functions** — a handful of common functions have no catalog
      entry yet: `TRANSPOSE` (a `transpose` *helper* exists but isn't exported as
      a function), `YEARFRAC`, `TEXTSPLIT`, `TEXTBEFORE` / `TEXTAFTER`, `EVEN` /
      `ODD`, and `SUMSQ`. Also `INDEX(range, 0, col)` / `INDEX(range, row, 0)`
      (whole-column / whole-row array form) returns `#REF!` instead of the
      column/row (`functions/lookup.ts`).

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
146-function formula engine with dynamic arrays, and formula autocomplete with
parameter help.
