# Core Features

The core package provides advanced spreadsheet features including sorting, filtering, freeze panes, and data export capabilities.

## Sorting System

### SortManager Architecture

The `SortManager` class handles multi-column sorting with formula adjustment:

```typescript
// packages/sheets-core/src/features/sort.ts
export class SortManager {
  static sortRows(sheet: Sheet, sortOrder: SortOrder[], dataRange?: { startRow: number; endRow: number }): void
  static toggleColumnSort(column: number, currentSortOrder: SortOrder[], multiColumn: boolean = false): SortOrder[]
  static getColumnSortDirection(column: number, currentSortOrder: SortOrder[]): 'asc' | 'desc' | null
}
```

### Sort Order Structure

Multi-column sorting is supported:

```typescript
interface SortOrder {
  column: number;        // Column index to sort by
  direction: 'asc' | 'desc';  // Sort direction
}

interface SheetConfig {
  sortOrder?: SortOrder[];  // Multi-column sort order
}
```

### Sorting Algorithm

The sorting process preserves data integrity:

1. **Data Range Detection**: Automatically detects data boundaries
2. **Row Extraction**: Collects all rows in the sort range
3. **Multi-Column Comparison**: Compares values across multiple columns
4. **Formula Adjustment**: Updates cell references in formulas after sorting
5. **Row Remapping**: Moves cells to new positions

```typescript
static sortRows(sheet: Sheet, sortOrder: SortOrder[], dataRange) {
  // Extract rows to sort
  const rowsToSort = [];
  for (let row = startRow; row <= endRow; row++) {
    const rowValues = [];
    for (let col = 0; col < sheet.colCount; col++) {
      const cell = sheet.cells.get(getCellKey(row, col));
      rowValues.push(cell?.value ?? null);
    }
    rowsToSort.push({ rowIndex: row, values: rowValues });
  }

  // Sort with multi-column comparison
  rowsToSort.sort((a, b) => {
    for (const sort of sortOrder) {
      const comparison = compareValues(a.values[sort.column], b.values[sort.column]);
      if (sort.direction === 'desc') comparison = -comparison;
      if (comparison !== 0) return comparison;
    }
    return 0;
  });

  // Remap and move cells
  const rowMapping = new Map();
  rowsToSort.forEach((row, newIndex) => {
    rowMapping.set(row.rowIndex, startRow + newIndex);
  });

  // Update formulas with new row positions
  this.adjustFormulasForRowSort(sheet, rowMapping, startRow, endRow);
}
```

### Sort Toggle Logic

Column sorting follows Excel-like behavior:

```typescript
static toggleColumnSort(column: number, currentSortOrder: SortOrder[], multiColumn: boolean) {
  const existingIndex = currentSortOrder.findIndex(s => s.column === column);

  if (existingIndex >= 0) {
    // Column already in sort order
    const existingSort = currentSortOrder[existingIndex];
    if (existingSort.direction === 'asc') {
      // asc -> desc
      return currentSortOrder.map(s =>
        s.column === column ? { ...s, direction: 'desc' } : s
      );
    } else {
      // desc -> remove sort
      return currentSortOrder.filter(s => s.column !== column);
    }
  } else {
    // Add new sort
    if (multiColumn && currentSortOrder.length > 0) {
      // Add as secondary sort
      return [...currentSortOrder, { column, direction: 'asc' }];
    } else {
      // Replace existing sort
      return [{ column, direction: 'asc' }];
    }
  }
}
```

### Formula Adjustment During Sort

When rows are reordered, formulas must be updated to maintain correct references:

```typescript
private static adjustFormulaForRowSort(
  formula: string,
  rowMapping: Map<number, number>,
  sortStartRow: number,
  sortEndRow: number
): string {
  // Replace cell references within sorted range
  const adjustedExpression = expression.replace(
    /(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?/g,
    (match, startRef, endRef) => {
      const startCellRef = parseCellReference(startRef);
      if (!startCellRef) return match;

      // Adjust row if within sorted range and not absolute
      let adjustedStartRow = startCellRef.row;
      if (!startCellRef.rowAbsolute &&
          startCellRef.row >= sortStartRow &&
          startCellRef.row <= sortEndRow) {
        const newRow = rowMapping.get(startCellRef.row);
        if (newRow !== undefined) {
          adjustedStartRow = newRow;
        }
      }

      // Handle range end if present
      if (endRef) {
        // Similar adjustment for end cell
        const endResult = adjustEndReference(endRef, rowMapping, sortStartRow, sortEndRow);
        return `${startResult}:${endResult}`;
      }

      return startResult;
    }
  );

  return '=' + adjustedExpression;
}
```

## Filtering System

### FilterManager Architecture

The `FilterManager` provides comprehensive filtering capabilities:

```typescript
// packages/sheets-core/src/features/filter.ts
export class FilterManager {
  static getFilteredRows(sheet: Sheet, filters: Map<number, ColumnFilter>, dataRange?): Set<number>
  static getUniqueColumnValues(sheet: Sheet, column: number, dataRange?): Set<string | number>
  static detectColumnType(sheet: Sheet, column: number): 'text' | 'number' | 'date'
  static hasFilter(column: number, filters: Map<number, ColumnFilter>): boolean
}
```

### Filter Criteria Types

Multiple filter types are supported:

```typescript
type FilterCriteria =
  | { type: 'equals'; value: string | number }
  | { type: 'notEquals'; value: string | number }
  | { type: 'contains'; value: string }
  | { type: 'notContains'; value: string }
  | { type: 'startsWith'; value: string }
  | { type: 'endsWith'; value: string }
  | { type: 'greaterThan'; value: number }
  | { type: 'lessThan'; value: number }
  | { type: 'greaterThanOrEqual'; value: number }
  | { type: 'lessThanOrEqual'; value: number }
  | { type: 'between'; min: number; max: number }
  | { type: 'custom'; values: Set<string | number> };

interface ColumnFilter {
  column: number;
  type: 'text' | 'number' | 'date';
  criteria: FilterCriteria;
}
```

### Filter Application

Filters are applied to determine visible rows:

```typescript
static getFilteredRows(sheet: Sheet, filters: Map<number, ColumnFilter>, dataRange): Set<number> {
  const visibleRows = new Set<number>();

  // Always include header row
  if (dataRange?.startRow <= 0) {
    visibleRows.add(0);
  }

  // Check each data row against all filters
  for (let row = startRow; row <= endRow; row++) {
    let isVisible = true;

    for (const [, filter] of filters) {
      const cell = sheet.cells.get(getCellKey(row, filter.column));
      const cellValue = cell?.value ?? null;

      if (!this.matchesFilter(cellValue, filter.criteria)) {
        isVisible = false;
        break;
      }
    }

    if (isVisible) {
      visibleRows.add(row);
    }
  }

  return visibleRows;
}
```

### Filter Matching Logic

Each filter type has specific matching logic:

```typescript
private static matchesFilter(cellValue: any, criteria: FilterCriteria): boolean {
  // Handle null values
  if (cellValue === null || cellValue === undefined) {
    return criteria.type === 'equals' && criteria.value === '';
  }

  const strValue = String(cellValue);

  switch (criteria.type) {
    case 'equals':
      return strValue === String(criteria.value);

    case 'contains':
      return strValue.toLowerCase().includes(criteria.value.toLowerCase());

    case 'greaterThan':
      const numValue = Number(cellValue);
      return !isNaN(numValue) && numValue > criteria.value;

    case 'between':
      const numValue = Number(cellValue);
      return !isNaN(numValue) && numValue >= criteria.min && numValue <= criteria.max;

    case 'custom':
      return criteria.values.has(cellValue);

    // ... other criteria types
  }
}
```

### Column Type Detection

Filters adapt based on column data type:

```typescript
static detectColumnType(sheet: Sheet, column: number): 'text' | 'number' | 'date' {
  const uniqueValues = this.getUniqueColumnValues(sheet, column);
  let hasNumbers = false;
  let hasDates = false;

  for (const value of uniqueValues) {
    if (typeof value === 'number') {
      hasNumbers = true;
    } else if (typeof value === 'string') {
      // Check for date-like strings
      const date = new Date(value);
      if (!isNaN(date.getTime()) && value.match(datePattern)) {
        hasDates = true;
      }
    }
  }

  if (hasDates) return 'date';
  if (hasNumbers) return 'number';
  return 'text';
}
```

## Freeze Panes

### Freeze Configuration

Freeze panes are configured through sheet configuration:

```typescript
interface SheetConfig {
  frozenRows?: number;    // Number of rows to freeze from top
  frozenCols?: number;    // Number of columns to freeze from left
}
```

### Four-Region Architecture

Freeze panes divide the viewport into four scroll regions:

```typescript
// packages/sheets-core/src/features/freeze.ts
type FreezeRegion =
  | 'top-left'    // Frozen rows AND cols - never scrolls
  | 'top'         // Frozen rows only - scrolls horizontally
  | 'left'        // Frozen cols only - scrolls vertically
  | 'main';       // Regular scrollable area - scrolls both ways
```

### Freeze Dimensions Calculation

```typescript
export function calculateFreezeDimensions(
  frozenRows: number,
  frozenCols: number,
  rowHeights: Map<number, number>,
  colWidths: Map<number, number>,
  defaultRowHeight: number,
  defaultColWidth: number
): FreezeDimensions {
  let frozenWidth = 0;
  for (let c = 0; c < frozenCols; c++) {
    frozenWidth += colWidths.get(c) ?? defaultColWidth;
  }

  let frozenHeight = 0;
  for (let r = 0; r < frozenRows; r++) {
    frozenHeight += rowHeights.get(r) ?? defaultRowHeight;
  }

  return { frozenWidth, frozenHeight };
}
```

### Region Coordinate Mapping

Each region has different scroll behavior:

```typescript
export function getScrollForRegion(
  region: FreezeRegion,
  scrollTop: number,
  scrollLeft: number
): { effectiveScrollTop: number; effectiveScrollLeft: number } {
  switch (region) {
    case 'top-left':
      return { effectiveScrollTop: 0, effectiveScrollLeft: 0 };
    case 'top':
      return { effectiveScrollTop: 0, effectiveScrollLeft: scrollLeft };
    case 'left':
      return { effectiveScrollTop: scrollTop, effectiveScrollLeft: 0 };
    case 'main':
      return { effectiveScrollTop: scrollTop, effectiveScrollLeft: scrollLeft };
  }
}
```

## Import / Export Functionality

All formats are read and written locally (no network). The format-level
parsers/writers live in `@b32nio/spindle-shared/export/csv` (CSV) and `@b32nio/spindle-io/xlsx` (XLSX) and know
nothing about the workbook; `sheets-core` adapts the workbook model onto them.

### CSV

```typescript
import { exportToCSV, importFromCSV } from '@b32nio/spindle-sheets-core';

const csv = exportToCSV(workbook);          // active sheet, or exportToCSV(workbook, sheetId)
importFromCSV(csv, workbook.getSheet());    // writes into an existing sheet
```

Export emits formulas (not their results) for formula cells, quotes fields that
contain commas, quotes or newlines, and is bounded by the last cell with data.
Import handles quoted fields, escaped quotes, embedded newlines and CRLF.

### XLSX

```typescript
import { exportToXlsx, importFromXlsx } from '@b32nio/spindle-sheets-core';

const bytes = await exportToXlsx(workbook);   // Uint8Array, every sheet

// Replaces the workbook's entire contents with the file's.
const { warnings } = await importFromXlsx(bytes, workbook);
```

Carried over: values, formulas, cell styles, number formats, merged cells,
column widths, row heights, hidden rows/columns, freeze panes, hyperlinks, and
the active sheet. Not carried over: charts, conditional formatting, data
validation, pivot tables. Anything the importer has to drop is reported in
`warnings`. Number formats are translated to and from Excel format codes
(`xlsx-number-format.ts`).

### In the UI

`WorkbookCanvas` shows Open (`.xlsx` / `.csv`) and Download (XLSX / CSV) in its
toolbar when given the `fileActions` prop; see the
[`sheets-react` README](../../../packages/sheets-react/README.md).

## Integration with Workbook

### Feature State Management

Features are managed through sheet configuration:

```typescript
// In workbook.ts
setSortOrder(sortOrder: SortOrder[], sheetId?: string): void {
  const sheet = this.getSheet(sheetId);
  sheet.config.sortOrder = sortOrder;
  SortManager.sortRows(sheet, sortOrder);
  this.events.emit('sheetChange', { sheetId });
}

setFilter(column: number, filter: ColumnFilter, sheetId?: string): void {
  const sheet = this.getSheet(sheetId);
  sheet.config.filters = sheet.config.filters || new Map();
  sheet.config.filters.set(column, filter);
  this.events.emit('sheetChange', { sheetId });
}
```

### Rendering Integration

Features affect rendering through the canvas system:

```typescript
// Render state includes feature data
interface RenderState {
  filters?: Map<number, ColumnFilter>;
  filteredRows?: Set<number>;
  frozenRows?: number;
  frozenCols?: number;
  sortOrder?: SortOrder[];
}
```

The features system provides comprehensive spreadsheet functionality while maintaining clean separation from the core rendering and data management systems.
