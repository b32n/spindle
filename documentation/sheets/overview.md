# Spindle Sheets — Architecture Overview

Spindle Sheets is a high-performance spreadsheet library for React, designed with a modular architecture that separates core functionality from framework-specific implementations.

## Architecture

The library is structured as a monorepo with two main packages:

```mermaid
graph TD
    A[Spindle] --> B[@b32nio/spindle-sheets-core]
    A --> C[@b32nio/spindle-sheets-react]
    B --> D[Zero React dependencies]
    C --> E[React components]
    B --> F[Framework agnostic]
    C --> G[Canvas-based rendering]
```

### Core Package (@b32nio/spindle-sheets-core)

The core package contains all spreadsheet logic and is framework-agnostic:

- **Workbook and Sheet models**: Data structures for spreadsheet state
- **Formula engine**: Parsing, dependency graph, and calculation
- **Rendering system**: Canvas-based rendering components
- **Features**: Sorting, filtering, freeze panes, export
- **Collaboration**: Real-time synchronization providers
- **Storage**: Sparse cell storage, style/format pooling

### Sheets Package (@b32nio/spindle-sheets-react)

The sheets package provides React components that use the core package:

- **WorkbookCanvas**: Main spreadsheet component
- **Context management**: React context for state management
- **UI components**: Toolbar, formula bar, sheet tabs
- **Modal dialogs**: Cell formatting, filtering, hyperlinks

## Design Principles

### Separation of Concerns

The core package is completely framework-agnostic, containing all business logic. The sheets package provides the React interface. This allows:

- Framework portability (future support for Vue, Angular, etc.)
- Easier testing of core functionality
- Clear boundaries between UI and business logic

### Performance-First Design

All architectural decisions prioritize performance:

- **Canvas rendering**: Hardware-accelerated graphics for smooth scrolling
- **Sparse storage**: Only non-empty cells consume memory
- **Style pooling**: Shared style objects reduce memory usage
- **Incremental updates**: Only changed regions are re-rendered
- **Formula optimization**: Dependency graph enables efficient recalculation

### Type Safety

TypeScript strict mode is enforced throughout:

- Comprehensive type definitions in `packages/sheets-core/src/types.ts`
- Generic constraints for type safety
- Interface segregation for clean APIs

## Package Structure

```
spindle/
├── packages/
│   ├── sheets-core/                # Framework-agnostic core
│   │   ├── src/
│   │   │   ├── canvas/             # Canvas rendering system
│   │   │   ├── collab/             # Real-time sync (Yjs binding)
│   │   │   ├── export/             # CSV import / export
│   │   │   ├── features/           # Filter, sort, freeze
│   │   │   ├── formula-parser/     # Formula parsing & functions
│   │   │   ├── utils/              # Shared helpers
│   │   │   ├── workbook.ts         # Main workbook model
│   │   │   └── types.ts            # Type definitions
│   │   └── package.json
│   └── sheets-react/               # React components
│       ├── src/
│       │   ├── components/         # React components
│       │   ├── context/            # React context
│       │   └── hooks/              # React hooks
│       └── package.json
├── examples/
│   └── sheets-demo/                # Standalone demo
└── documentation/                  # Documentation
```

## Key Interfaces

### Workbook

The main API surface for spreadsheet operations:

```typescript
interface Workbook {
  id: string;
  name: string;
  sheets: Map<string, Sheet>;
  activeSheetId: string;

  // Cell operations
  setCellValue(sheetId: string | undefined, row: number, col: number, value: CellValue): void;
  getCellValue(sheetId: string | undefined, row: number, col: number): CellValue;

  // Sheet management
  addSheet(name: string): Sheet;
  deleteSheet(sheetId: string): void;

  // Features
  setSortOrder(sortOrder: SortOrder[], sheetId?: string): void;
  setFilter(column: number, filter: ColumnFilter, sheetId?: string): void;

  // Serialization
  getData(): WorkbookData;
  setData(data: WorkbookData): void;
}
```

### Sheet

Represents a single worksheet:

```typescript
interface Sheet {
  id: string;
  name: string;
  cells: Map<string, Cell>;  // Sparse storage: key format "row:col"
  config: SheetConfig;
  rowCount: number;
  colCount: number;
}
```

## Usage Patterns

### Basic Usage

```typescript
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';
import { WorkbookProvider, WorkbookCanvas } from '@b32nio/spindle-sheets-react';

const workbook = new WorkbookImpl('workbook_1', 'My Workbook');

// In React
function App() {
  return (
    <WorkbookProvider workbook={workbook}>
      <WorkbookCanvas width={800} height={600} />
    </WorkbookProvider>
  );
}
```

### Advanced Usage

```typescript
import { WorkbookImpl, exportToCSV } from '@b32nio/spindle-sheets-core';
import { WebSocketProvider } from '@b32nio/spindle-transport-websocket';

// Load existing data
workbook.setData(workbookData);

// Set up real-time collaboration (see the Collaboration guide)
const provider = new WebSocketProvider({ url: 'wss://collab.example.com' });
await workbook.attachCollab(provider, {
  userId: 'u_42',
  displayName: 'Bharat',
  color: '#4ecdc4',
});

// Add event listeners
workbook.on('cellChange', (event) => {
  console.log('Cell changed:', event.payload);
});

// Export data
const csvData = exportToCSV(workbook);
```

## Extension Points

The modular architecture provides clear extension points:

- **Formula functions**: Add new functions to the formula parser
- **Export formats**: Implement new export formats
- **Collaboration providers**: Add new real-time sync backends
- **Rendering customization**: Extend or replace rendering components
- **Features**: Add new spreadsheet features

See the contributing documentation for implementation details.
