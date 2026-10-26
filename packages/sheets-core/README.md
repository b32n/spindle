# @b32nio/spindle-sheets-core

Core spreadsheet engine for Spindle Sheets. Zero React dependencies.

## Features

- Sparse cell storage (Map-based)
- Efficient range operations
- Formula dependency graph
- Style pooling for memory efficiency
- Event system with batching
- TypeScript strict mode

## Usage

```typescript
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';

const workbook = new WorkbookImpl('workbook_1', 'My Workbook');
const sheet = workbook.getSheet();

workbook.setCellValue(undefined, 0, 0, 'Hello');
workbook.setCellValue(undefined, 0, 1, 'World');
```

## Import / export (CSV, XLSX)

```typescript
import { exportToCSV, importFromCSV, exportToXlsx, importFromXlsx } from '@b32nio/spindle-sheets-core';

const csv = exportToCSV(workbook);                 // active sheet, or pass a sheet id
importFromCSV(csv, workbook.getSheet());           // into an existing sheet

const bytes = await exportToXlsx(workbook);        // Uint8Array (.xlsx), all sheets
// Replaces the workbook's entire contents with the file's.
const { warnings } = await importFromXlsx(bytes, workbook);
```

XLSX carries values, formulas, styles, number formats, merges, column widths,
row heights, hidden rows/columns, freeze panes and hyperlinks; charts, conditional formatting
and data validation are not carried over. The format-level readers/writers
live in `@b32nio/spindle-shared/export/csv` and `@b32nio/spindle-io/xlsx` (a dependency of this package, loaded on first
Open/Download). Everything runs locally.

## Performance

- O(1) cell lookups
- Memory efficient for sparse sheets
- Incremental formula recalculation
- Shared style objects

