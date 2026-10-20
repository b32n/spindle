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

### Remaining formula items

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
- [ ] **XLSX fidelity gaps** — XLSX import/export shipped in 0.4.0
      (`exportToXlsx` / `importFromXlsx`, plus Open/Download in the toolbar), but
      charts, conditional formatting and data validation are not carried over.
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
