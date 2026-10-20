# React Components Reference

The sheets package provides a comprehensive set of React components for building spreadsheet interfaces.

## Core Components

### WorkbookCanvas

The main spreadsheet component that orchestrates all UI elements.

```typescript
interface WorkbookCanvasProps {
  className?: string;
  style?: React.CSSProperties;
  width?: number;
  height?: number;
  rowHeight?: number;
  colWidth?: number;
  readOnly?: boolean;          // disables all editing and hides the toolbar
  fileActions?: boolean;       // show Open (.xlsx/.csv) and Download (XLSX/CSV) in the toolbar; default false
  exportFormats?: ('xlsx' | 'csv')[];   // formats offered by Download; default both
  fileName?: string;           // base name for downloads; default: the workbook name
  onFileOpened?: (info: { fileName: string; warnings: string[] }) => void;
  onFileError?: (error: Error, action: 'open' | 'download') => void;
}

export const WorkbookCanvas = memo(function WorkbookCanvas(props: WorkbookCanvasProps) {
  // Main component logic
});
```

**Features:**
- Layout management for toolbar, formula bar, canvas, and sheet tabs
- State coordination between all sub-components
- Event handling and propagation
- Modal dialog management
- Clipboard operations
- Cell editing workflow

**Layout Structure:**
```jsx
<div className="workbook-canvas">
  <Toolbar />
  <FormulaBar />
  <div className="canvas-area">
    <CanvasGrid />
    <EditOverlay /> {/* When editing */}
  </div>
  <SheetTabs />
  {/* Modal dialogs */}
</div>
```

### CanvasGrid

Low-level canvas-based grid component that handles rendering and interaction.

```typescript
interface CanvasGridProps {
  width: number;
  height: number;
  rowHeight?: number;
  colWidth?: number;
  activeCell?: CellPosition | null;
  onActiveCellChange?: (cell: CellPosition | null) => void;
  onCellEdit?: (cell: CellPosition, value: string) => void;
  onSelectionChange?: (selection: Selection) => void;
  onScroll?: (scrollTop: number, scrollLeft: number) => void;
  dimensionVersion?: number;
  editValue?: string;
  onInsertCellReference?: (reference: string) => void;
  isEditingFormula?: boolean;
  editingCell?: CellPosition | null;
  onContextMenu?: (menu: ContextMenuType) => void;
  onClipboardReady?: (handlers: ClipboardHandlers) => void;
}

export const CanvasGrid = memo(function CanvasGrid(props: CanvasGridProps) {
  // Canvas rendering and interaction logic
});
```

**Features:**
- Canvas rendering integration
- Mouse event handling (click, double-click, drag)
- Keyboard navigation
- Scroll management
- Hit testing for mouse coordinates
- Formula range highlighting
- Context menu triggering

**Canvas Integration:**
```typescript
const canvasRef = useRef<HTMLCanvasElement>(null);
const rendererRef = useRef<CanvasRenderer | null>(null);

// Initialize canvas renderer
useEffect(() => {
  rendererRef.current = new CanvasRenderer({
    canvas: canvasRef.current!,
    defaultRowHeight: rowHeight,
    defaultColWidth: colWidth,
  });
}, []);

// Update render state
useEffect(() => {
  const renderState: RenderState = {
    cells: sheet.cells,
    styles: workbook.stylePool.getAllStyles(),
    // ... other state
  };
  rendererRef.current?.updateRenderState(renderState);
}, [workbook, activeCell]);
```

### EditOverlay

DOM-based input overlay for cell editing positioned over the canvas.

```typescript
interface EditOverlayProps {
  cell: CellPosition;
  value: string;
  onChange: (value: string) => void;
  onCommit: (value: string) => void;
  onCancel: () => void;
  x: number;
  y: number;
  width: number;
  height: number;
  minWidth?: number;
  fontSize?: number;
  fontFamily?: string;
  isEditingFormula?: boolean;
  cellFormat?: CellFormat;
}

interface EditOverlayRef {
  insertAtCursor: (text: string, replaceExisting?: boolean) => void;
  getCursorPosition: () => number;
  focus: () => void;
}

export const EditOverlay = memo(forwardRef<EditOverlayRef, EditOverlayProps>(
  function EditOverlay(props, ref) {
    // DOM input overlay logic
  }
));
```

**Features:**
- Positioned DOM input over canvas cell
- Formula editing with syntax highlighting
- Cell reference insertion
- Keyboard navigation (Enter, Escape, Tab)
- Auto-sizing based on content
- Format-aware input (dates, numbers)

## UI Components

### Toolbar

Formatting and action toolbar at the top of the spreadsheet.

```typescript
interface ToolbarProps {
  // Text formatting
  onBold?: () => void;
  onItalic?: () => void;
  onUnderline?: () => void;
  onStrikethrough?: () => void;

  // Font properties
  onFontFamily?: (fontFamily: string) => void;
  onFontSize?: (fontSize: number) => void;
  onFontColor?: (color: string) => void;
  onBackgroundColor?: (color: string) => void;

  // Alignment
  onAlignLeft?: () => void;
  onAlignCenter?: () => void;
  onAlignRight?: () => void;
  onVerticalAlign?: (align: 'top' | 'middle' | 'bottom') => void;

  // Cell formatting
  onTextWrap?: () => void;
  onBorder?: (border: string) => void;
  onMergeCells?: () => void;
  onHyperlink?: (url: string) => void;

  // Data formatting
  onFormatCurrency?: () => void;
  onFormatPercentage?: () => void;
  onFormatNumber?: () => void;

  // History
  onUndo?: () => void;
  onRedo?: () => void;

  // Freeze panes
  onFreezeRows?: (rows: number) => void;
  onFreezeCols?: (cols: number) => void;
  onUnfreeze?: () => void;

  // Current selection state
  selectedFormat?: {
    bold?: boolean;
    italic?: boolean;
    fontFamily?: string;
    fontSize?: number;
    // ... other format properties
  };
  frozenRows?: number;
  frozenCols?: number;
}

export const Toolbar = memo(function Toolbar(props: ToolbarProps) {
  // Toolbar implementation
});
```

**Features:**
- Text formatting controls (bold, italic, underline)
- Font selection (family, size, color)
- Cell formatting (borders, background, alignment)
- Data formatting (currency, percentage, number)
- Freeze pane controls
- Undo/redo buttons
- Hyperlink insertion
- Merge cells functionality

**State Synchronization:**
The toolbar reflects the current selection's formatting and updates the workbook when controls are used.

### FormulaBar

Formula input bar that shows and allows editing of cell formulas/values.

```typescript
interface FormulaBarProps {
  activeCell: CellPosition | null;
  value?: string;
  onChange?: (value: string) => void;
  onCommit?: (value: string) => void;
  onCancel?: () => void;
  onFormulaChange?: (formula: string) => void;
}

export const FormulaBar = memo(function FormulaBar(props: FormulaBarProps) {
  // Formula bar implementation
});
```

**Features:**
- Displays current cell value or formula
- Allows direct editing of formulas
- Shows cell reference (e.g., "A1")
- Formula validation and syntax highlighting
- Integration with cell editing workflow

**Editing Workflow:**
```typescript
// When cell is selected, show its value/formula
useEffect(() => {
  if (activeCell) {
    const cell = workbook.getCell(activeCell.row, activeCell.col);
    const displayValue = cell?.formula || cell?.value?.toString() || '';
    setInputValue(displayValue);
  }
}, [activeCell]);

// When user edits in formula bar
const handleInputChange = (value: string) => {
  setInputValue(value);
  onFormulaChange?.(value);
};
```

### SheetTabs

Sheet navigation and management component.

```typescript
interface SheetTabsProps {
  sheets?: Sheet[];
  activeSheetId?: string;
  onSheetSelect?: (sheetId: string) => void;
  onSheetRename?: (sheetId: string, newName: string) => void;
  onSheetAdd?: () => void;
  onSheetDelete?: (sheetId: string) => void;
}

export const SheetTabs = memo(function SheetTabs(props: SheetTabsProps) {
  // Sheet tabs implementation
});
```

**Features:**
- Visual sheet tabs with active sheet indication
- Sheet creation (+ button)
- Sheet deletion (right-click menu)
- Sheet renaming (double-click)
- Drag and drop reordering (future feature)

**Sheet Management:**
```typescript
const handleSheetAdd = useCallback(() => {
  updateWorkbook(wb => {
    wb.addSheet(`Sheet${wb.sheets.size + 1}`);
  });
}, [updateWorkbook]);

const handleSheetDelete = useCallback((sheetId: string) => {
  updateWorkbook(wb => {
    wb.deleteSheet(sheetId);
  });
}, [updateWorkbook]);
```

## Context Menus

### ContextMenu

Right-click context menu for cells, rows, and columns.

```typescript
interface ContextMenuProps {
  x: number;
  y: number;
  onClose: () => void;

  // Clipboard operations
  onCopy?: () => void;
  onCut?: () => void;
  onPaste?: () => void;

  // Cell operations
  onDelete?: () => void;
  onClear?: () => void;

  // Row/Column operations
  onInsertRow?: () => void;
  onInsertRowBelow?: () => void;
  onInsertColumn?: () => void;
  onInsertColumnRight?: () => void;
  onDeleteRow?: () => void;
  onDeleteColumn?: () => void;

  // Formatting
  onFormat?: () => void;
}

export const ContextMenu = memo(function ContextMenu(props: ContextMenuProps) {
  // Context menu implementation
});
```

**Context-Aware Options:**
- **Cell context**: Copy, cut, paste, delete, clear, format
- **Row context**: Insert row above/below, delete row
- **Column context**: Insert column left/right, delete column

**Positioning:**
```typescript
const menuStyle: React.CSSProperties = {
  position: 'fixed',
  left: x,
  top: y,
  zIndex: 1000,
  // Ensure menu stays within viewport
  maxHeight: window.innerHeight - y - 10,
  overflowY: 'auto',
};
```

### HeaderContextMenu

Specialized context menu for row and column headers.

```typescript
interface HeaderContextMenuProps {
  type: 'row' | 'column';
  index: number;
  x: number;
  y: number;
  onClose: () => void;
  onDelete?: () => void;
  onInsert?: () => void;
  onResize?: (size: number) => void;
}

export const HeaderContextMenu = memo(function HeaderContextMenu(props: HeaderContextMenuProps) {
  // Header-specific context menu
});
```

**Features:**
- Row/column specific operations
- Resize options
- Insert/delete operations

## Modal Components

### FilterModal

Advanced filtering dialog for columns.

```typescript
interface FilterModalProps {
  isOpen: boolean;
  sheet: Sheet;
  column: number;
  existingFilter?: ColumnFilter;
  onClose: () => void;
  onApply: (filter: ColumnFilter) => void;
  onClear: () => void;
}

export const FilterModal = memo(function FilterModal(props: FilterModalProps) {
  // Filter modal implementation
});
```

**Features:**
- Filter by condition (equals, contains, greater than, etc.)
- Custom value lists
- Date range filtering
- Number range filtering
- Text pattern matching

**Filter Types:**
```typescript
const TEXT_FILTER_OPTIONS = [
  { value: 'equals', label: 'Equals' },
  { value: 'contains', label: 'Contains' },
  { value: 'startsWith', label: 'Starts with' },
  // ... more options
];

const NUMBER_FILTER_OPTIONS = [
  { value: 'greaterThan', label: 'Greater than' },
  { value: 'lessThan', label: 'Less than' },
  { value: 'between', label: 'Between' },
  // ... more options
];
```

### FormatCellsModal

Cell formatting dialog.

```typescript
interface FormatCellsModalProps {
  isOpen: boolean;
  currentFormat?: CellFormat;
  sampleValue?: number;
  onClose: () => void;
  onApply: (format: CellFormat) => void;
}

export const FormatCellsModal = memo(function FormatCellsModal(props: FormatCellsModalProps) {
  // Format cells modal implementation
});
```

**Features:**
- Number formatting (decimal places, separators)
- Currency formatting (symbol, position)
- Date/time formatting patterns
- Custom format strings
- Preview of formatted values

### HyperlinkModal

Hyperlink insertion dialog.

```typescript
interface HyperlinkModalProps {
  isOpen: boolean;
  initialUrl?: string;
  onClose: () => void;
  onApply: (url: string, text?: string) => void;
}

export const HyperlinkModal = memo(function HyperlinkModal(props: HyperlinkModalProps) {
  // Hyperlink modal implementation
});
```

**Features:**
- URL validation
- Link text customization
- Preview of hyperlink
- Edit existing hyperlinks

## Utility Components

### FormulaReferenceOverlay

Shows cell references when editing formulas.

```typescript
interface FormulaReferenceOverlayProps {
  ranges: FormulaRangeHighlight[];
  viewport: Viewport;
  // ... positioning props
}

export const FormulaReferenceOverlay = memo(function FormulaReferenceOverlay(props: FormulaReferenceOverlayProps) {
  // Formula reference highlighting
});
```

**Features:**
- Visual highlighting of referenced ranges
- Color-coded reference groups
- Dynamic updates during formula editing

## Component Architecture Patterns

### Memoization

All components use `React.memo` for performance:

```typescript
export const WorkbookCanvas = memo(function WorkbookCanvas(props) {
  // Component logic
});
```

### Context Usage

Components access workbook state through context:

```typescript
const { workbook, updateWorkbook } = useWorkbook();

// Read state
const sheets = workbook.sheets;

// Update state
updateWorkbook(wb => {
  wb.setCellValue(undefined, row, col, value);
});
```

### Ref-Based Communication

Performance-critical communication uses refs:

```typescript
const editOverlayRef = useRef<EditOverlayRef>(null);

// Direct method calls
editOverlayRef.current?.insertAtCursor(reference);
```

### Event-Driven Updates

Components subscribe to workbook events for reactive updates:

```typescript
useEffect(() => {
  const handleCellChange = () => {
    // Trigger re-render
    setUpdateTrigger(prev => prev + 1);
  };

  workbook.on('cellChange', handleCellChange);
  return () => workbook.off('cellChange', handleCellChange);
}, [workbook]);
```

The component library provides a complete, performant React interface for spreadsheet functionality with consistent patterns and comprehensive features.
