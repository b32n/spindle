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

## Performance

- O(1) cell lookups
- Memory efficient for sparse sheets
- Incremental formula recalculation
- Shared style objects

