# Spindle Documentation

Spindle is a suite of high-performance productivity components for React, providing Google Docs-like document editing and Google Sheets-like spreadsheet functionality.

## Packages

| Package | Description |
|---------|-------------|
| `@b32nio/spindle-docs-core` | Framework-agnostic document engine |
| `@b32nio/spindle-docs-react` | React document editor with true page layout |
| `@b32nio/spindle-sheets-core` | Framework-agnostic spreadsheet engine |
| `@b32nio/spindle-sheets-react` | React spreadsheet with canvas rendering |
| `@b32nio/spindle-slides-core` | Framework-agnostic presentation engine |
| `@b32nio/spindle-slides-react` | React slides editor with DOM rendering |
| `@b32nio/spindle-shared` | Shared utilities, types, and React UI |
| `@b32nio/spindle-transport-websocket` | WebSocket `CollabProvider` for real-time collaboration |

---

## 🤝 Real-Time Collaboration

Documents and spreadsheets sync live between users via Yjs CRDTs.

- **[Collaboration Guide](collaboration.md)** — `attachCollab`, transports, offline persistence, connection status, and an end-to-end encryption recipe

---

## 💬 Comments

Comment threads on spreadsheet cells and document text — replies, resolve, and @-mentions.

- **[Comments Guide](comments.md)** — the comment model, cell/text anchoring, @-mentions, and the `onCommentEvent` notification callback

---

## 📄 Document Editor (Docs)

A Google Docs-like document editor with true page-based layout and real-time pagination.

### Documentation

- **[Overview](docs/overview.md)** - High-level tour of the two packages and design principles
- **[Architecture](docs/architecture.md)** - Core and React package internals, extension points
- **[Components](docs/components.md)** - Complete guide to React components
- **[Data Structures](docs/data-structures.md)** - Type definitions and interfaces

### Quick Start

```typescript
import { DocumentImpl } from '@b32nio/spindle-docs-core';
import { DocumentProvider, DocumentEditor } from '@b32nio/spindle-docs-react';

// Create document
const doc = new DocumentImpl('doc_1', 'My Document');

function App() {
  return (
    <DocumentProvider document={doc}>
      <DocumentEditor showToolbar={true} showRuler={true} />
    </DocumentProvider>
  );
}
```

### Key Features

- **True Page Layout**: Line-level pagination for accurate print preview
- **Rich Text Editing**: Full formatting with ProseMirror-powered editing
- **Headers & Footers**: Dynamic fields (page numbers, date, title)
- **Tables**: Full table support with cell formatting
- **Images**: Block and inline images with resizing
- **Undo/Redo**: Complete history management

---

## 📊 Spreadsheet (Sheets)

A Google Sheets-like spreadsheet with high-performance canvas rendering.

### Documentation

- **[Overview](sheets/overview.md)** - High-level architecture and design principles
- **[Architecture](sheets/core/architecture.md)** - Workbook & Sheet models, sparse storage, and the React layer
- **[Rendering](sheets/core/rendering.md)** - Canvas rendering, viewport optimization
- **[Formula Engine](sheets/core/formulas.md)** - Formula parser, dependency graph, 146 built-in functions
- **[Features](sheets/core/features.md)** - Sorting, filtering, freeze panes
- **[Collaboration](sheets/core/collaboration.md)** - Sheet-specific real-time sync internals
- **[Components](sheets/components.md)** - All React components
- **[Data Structures](sheets/data-structures.md)** - Type definitions

### Quick Start

```typescript
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';
import { WorkbookProvider, WorkbookCanvas } from '@b32nio/spindle-sheets-react';

// Create workbook
const workbook = new WorkbookImpl('workbook_1', 'My Workbook');

function App() {
  return (
    <WorkbookProvider workbook={workbook}>
      <WorkbookCanvas width={800} height={600} />
    </WorkbookProvider>
  );
}
```

### Key Features

- **Canvas Rendering**: 60fps performance for large datasets
- **Formula Engine**: 146 functions, dynamic arrays, topological recalculation, autocomplete
- **Sparse Storage**: Efficient memory usage
- **Freeze Panes**: Lock rows/columns while scrolling
- **Sorting & Filtering**: Column operations
- **CSV Export/Import**: Data interchange

See the [issue tracker](https://github.com/b32n/spindle/issues) for planned work.

---

## 🖼️ Presentations (Slides)

A collaborative presentation editor with DOM rendering, a scene/element model,
and present mode.

### Documentation

- **[Overview](slides/overview.md)** - High-level tour and design principles
- **[Architecture](slides/core/architecture.md)** - Scene model, geometry, and the React layer
- **[Collaboration](slides/core/collaboration.md)** - Slides-specific real-time sync
- **[Data Structures](slides/data-structures.md)** - Type definitions
- **[Components](slides/components.md)** - React components

### Quick Start

```typescript
import { DeckImpl } from '@b32nio/spindle-slides-core';
import { DeckProvider, SlidesEditor } from '@b32nio/spindle-slides-react';

// Create a deck
const deck = new DeckImpl('deck_1', 'My Deck');

function App() {
  return (
    <DeckProvider deck={deck}>
      <SlidesEditor />
    </DeckProvider>
  );
}
```

### Key Features

- **Element Model**: text, shapes (~18 SVG presets), images, lines/connectors, and **tables**
- **Tables**: per-cell rich text, cell/row/column/range selection, paste a spreadsheet range as a table
- **Interactions**: drag/resize/rotate at 60fps, smart-guide snapping, align/distribute, groups, z-order
- **Themes & Layouts**: symbolic PPTX color slots, layout gallery with placeholder materialization
- **Present Mode**: fullscreen 16:9, keyboard nav, cross-fade, and presenter view (timer, notes, next-slide preview)
- **Comments & Collaboration**: element-anchored threads and Yjs real-time sync

See the [issue tracker](https://github.com/b32n/spindle/issues) for planned work.

---

## Architecture Overview

```
spindle/
├── packages/
│   ├── docs-core/          # Document engine
│   │   ├── document.ts     # Document model
│   │   ├── blocks/         # Block types
│   │   └── prosemirror/    # PM integration
│   │
│   ├── docs-react/         # Document React components
│   │   ├── components/     # UI components
│   │   ├── context/        # React context
│   │   └── core/           # Layout engine
│   │
│   ├── sheets-core/        # Spreadsheet engine
│   │   ├── workbook.ts     # Workbook model
│   │   ├── canvas/         # Rendering
│   │   └── formula-parser/ # Formulas
│   │
│   ├── sheets-react/       # Spreadsheet React components
│   │   ├── components/     # UI components
│   │   └── context/        # React context
│   │
│   ├── slides-core/        # Presentation engine
│   │   ├── scene/          # Element/scene model
│   │   ├── theme/          # Themes & layouts
│   │   └── collab/         # Yjs collaboration
│   │
│   ├── slides-react/       # Slides React components
│   │   ├── components/     # UI components
│   │   └── interactions/   # Gestures & stores
│   │
│   ├── shared/             # Shared utilities + React UI
│   └── transport-websocket/ # WebSocket CollabProvider
│
├── examples/
│   ├── docs-demo/          # Document editor demo
│   ├── sheets-demo/        # Spreadsheet demo
│   ├── slides-demo/        # Presentation demo
│   └── collab-server/      # Local collaboration relay
│
└── documentation/          # This documentation
    ├── docs/               # Document editor docs
    ├── sheets/             # Spreadsheet docs
    └── slides/             # Presentation docs
```

---

## Design Principles

### Performance-First
- Canvas rendering for spreadsheets at 60fps
- Virtual rendering for documents
- Sparse storage for memory efficiency

### Framework-Agnostic Core
- Core packages work without React
- Easy to port to Vue, Angular, etc.
- Clear separation of concerns

### Type Safety
- Full TypeScript with strict mode
- Comprehensive type definitions
- No `any` types

### Extensibility
- Plugin architecture
- Custom block types
- Custom formula functions

---

## Development Setup

```bash
# Install dependencies
npm install

# Build all packages
npm run build

# Run tests
npm run test

# Type checking
npm run type-check

# Start dev server
npm run dev
```

---

## Examples

- **[Docs Demo](../examples/docs-demo/)** - Complete document editor
- **[Sheets Demo](../examples/sheets-demo/)** - Complete spreadsheet
- **[Slides Demo](../examples/slides-demo/)** - Complete presentation editor

---

## Contributing

- **[Extending the Library](sheets/contributing/extending.md)** - Add new features

---

**Last Updated**: July 2026  
**Version**: 0.3.0
