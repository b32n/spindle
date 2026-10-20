# @b32nio/spindle-sheets-react

React components for Spindle Sheets with Canvas-based rendering.

## Features

- **Canvas-based rendering**: High-performance grid rendering with pixel-perfect scroll synchronization
- **Optimized React components**: Minimal re-renders with memoization
- **Context-based state management**: Clean separation of state and UI
- **TypeScript strict mode**: Full type safety
- **Retina display support**: Crisp rendering on high-DPI displays

## Components

### WorkbookCanvas (main component)
The main spreadsheet component using canvas rendering. Provides:
- Cell grid with headers
- Selection and active cell highlighting
- Cell editing with DOM overlay
- Toolbar and formula bar
- Sheet tabs

### CanvasGrid
Low-level canvas grid component for custom implementations.

### EditOverlay
DOM-based input overlay for cell editing.

## Usage

```typescript
import { Workbook, WorkbookProvider } from '@b32nio/spindle-sheets-react';
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';

const workbook = new WorkbookImpl('workbook_1', 'My Workbook');

function App() {
  return (
    <WorkbookProvider workbook={workbook}>
      <Workbook width={800} height={600} />
    </WorkbookProvider>
  );
}
```

## Open / Download (XLSX, CSV)

`WorkbookCanvas` can show **Open** (a local `.xlsx` / `.csv` file picker) and a
**Download** menu (XLSX, CSV) in its toolbar. It is opt-in:

```tsx
<WorkbookCanvas
  fileActions
  exportFormats={['xlsx', 'csv']}           // optional; default is both
  fileName="Budget"                         // optional; defaults to the workbook name
  onFileOpened={({ fileName, warnings }) => {}}
  onFileError={(error, action) => {}}       // action: 'open' | 'download'
/>
```

Opening a file replaces the whole workbook (a CSV becomes a single-sheet
workbook); CSV download exports the active sheet. Files are read locally via
`<input type="file">`; nothing is uploaded. For a host-owned button use the
exported helpers `exportWorkbookFile(workbook, 'xlsx')` and
`importWorkbookFile(workbook, file)`, or the `Toolbar` props `onOpenFile` /
`onDownload` / `downloadFormats`.

## Architecture

The rendering architecture uses a canvas-based approach:

```
WorkbookCanvas (React)
├── Toolbar (DOM)
├── FormulaBar (DOM)
├── CanvasGrid (Canvas)
│   ├── Cell rendering
│   ├── Grid lines
│   ├── Row/Column headers
│   ├── Selection highlighting
│   └── Fill handle
├── EditOverlay (DOM - shown during editing)
└── SheetTabs (DOM)
```

Canvas handles all grid rendering for optimal performance, while DOM is used for interactive elements like inputs and buttons.
