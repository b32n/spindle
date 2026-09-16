# Spindle Docs — Architecture

Spindle Docs is a high-performance document editor library for React, providing a Google Docs-like editing experience with true page-based layout and real-time pagination. For a high-level tour of the two packages and the design principles behind them, start with the **[Overview](overview.md)**; this page covers the internals of each package.

## Core Package (@b32nio/spindle-docs-core)

The core package provides the framework-agnostic document engine:

### Document Model

```typescript
import { DocumentImpl, createDocument } from '@b32nio/spindle-docs-core';

// Create a new document
const doc = new DocumentImpl('doc_1', 'My Document');

// Access sections and blocks
const sections = doc.getSections();
const blocks = sections[0].blocks;

// Modify content
doc.setTitle('New Title');
doc.setSectionBlocks(sectionId, newBlocks);
```

### Key Concepts

- **Document**: The root container with title, sections, and page configuration
- **Section**: A page configuration boundary containing blocks with optional header/footer
- **Block**: A content element (paragraph, heading, list-item, table, image, etc.)
- **InlineContent**: Text runs, images, and links within blocks

### ProseMirror Integration

The core package provides full ProseMirror integration:

```typescript
import { docsSchema, createPlugins, createCommands } from '@b32nio/spindle-docs-core';

// Schema defines document structure
const schema = docsSchema;

// Create editing commands
const commands = createCommands(schema);

// Available commands
commands.toggleBold();
commands.toggleItalic();
commands.setHeading(2);
commands.insertTable(3, 3);
// ... and more
```

### Block Types

The following block types are supported:

| Type | Description |
|------|-------------|
| `paragraph` | Standard text paragraph with inline content |
| `heading` | Heading levels 1-6 with content |
| `list-item` | Bullet or numbered list item with nesting |
| `table` | Table with rows, cells, and column widths |
| `image` | Block-level image with alignment and caption |
| `horizontal-rule` | Horizontal divider |
| `page-break` | Force page break in layout |

### Style Pools

Efficient style management through pooling:

```typescript
// Text styles are deduplicated
const styleId = doc.getTextStylePool().getOrCreate({
  bold: true,
  fontSize: 14,
  color: '#333333',
});

// Paragraph styles too
const paraStyleId = doc.getParagraphStylePool().getOrCreate({
  alignment: 'center',
  lineSpacing: { type: 'double' },
});
```

### Event System

```typescript
// Listen for document changes
doc.on('documentChange', ({ payload }) => {
  console.log('Document changed:', payload);
});

// Other events: blockChange, blockAdd, blockDelete, 
// selectionChange, pageConfigChange, historyChange
```

### History (Undo/Redo)

```typescript
// Record history checkpoint
doc.recordHistory('Made important change');

// Undo/Redo
if (doc.canUndo()) doc.undo();
if (doc.canRedo()) doc.redo();
```

## React Package (@b32nio/spindle-docs-react)

### True Layout Architecture

The React package implements a sophisticated layout engine for paginated document editing:

```
┌──────────────────────────────────────────────────────────────┐
│                     TrueLayoutEditor                          │
├──────────────────────────────────────────────────────────────┤
│                                                               │
│  ┌─────────────────┐    ┌─────────────────────────────────┐ │
│  │  Hidden Editor  │───▶│     FlowBlocks Converter        │ │
│  │  (ProseMirror)  │    │  (PM Doc → Layout Blocks)       │ │
│  └─────────────────┘    └─────────────────────────────────┘ │
│          │                            │                      │
│          │                            ▼                      │
│          │              ┌─────────────────────────────────┐ │
│          │              │        DOM Measurer             │ │
│          │              │   (Block → Line Measurements)   │ │
│          │              └─────────────────────────────────┘ │
│          │                            │                      │
│          │                            ▼                      │
│          │              ┌─────────────────────────────────┐ │
│          │              │      True Layout Engine         │ │
│          │              │   (Pagination & Line Breaking)  │ │
│          │              └─────────────────────────────────┘ │
│          │                            │                      │
│          │                            ▼                      │
│  ┌───────┴───────────────────────────────────────────────┐ │
│  │                                                         │ │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────────┐│ │
│  │  │   Page 1    │  │   Page 2    │  │    Page N       ││ │
│  │  │ ┌─────────┐ │  │ ┌─────────┐ │  │  ┌─────────┐   ││ │
│  │  │ │ Header  │ │  │ │ Header  │ │  │  │ Header  │   ││ │
│  │  │ ├─────────┤ │  │ ├─────────┤ │  │  ├─────────┤   ││ │
│  │  │ │         │ │  │ │         │ │  │  │         │   ││ │
│  │  │ │ Content │ │  │ │ Content │ │  │  │ Content │   ││ │
│  │  │ │         │ │  │ │         │ │  │  │         │   ││ │
│  │  │ ├─────────┤ │  │ ├─────────┤ │  │  ├─────────┤   ││ │
│  │  │ │ Footer  │ │  │ │ Footer  │ │  │  │ Footer  │   ││ │
│  │  │ └─────────┘ │  │ └─────────┘ │  │  └─────────┘   ││ │
│  │  └─────────────┘  └─────────────┘  └─────────────────┘│ │
│  │                      DOM Painter                       │ │
│  └───────────────────────────────────────────────────────┘ │
│          ▲                            ▲                      │
│          │                            │                      │
│  ┌───────┴─────────┐    ┌─────────────┴───────────────────┐ │
│  │  Input Bridge   │    │      Selection Overlay          │ │
│  │ (Events → PM)   │    │   (Caret & Selection Rects)     │ │
│  └─────────────────┘    └─────────────────────────────────┘ │
│                                                               │
└──────────────────────────────────────────────────────────────┘
```

### Component Flow

1. **Hidden ProseMirror Editor**: A hidden editor instance handles all text editing, providing battle-tested text input handling, undo/redo, and cursor management.

2. **FlowBlocks Converter**: Converts ProseMirror document nodes into FlowBlocks - an intermediate representation optimized for layout.

3. **DOM Measurer**: Measures each block's dimensions by rendering to a hidden container. Captures line-level measurements for precise pagination.

4. **True Layout Engine**: Computes page layout with line-level pagination:
   - Assigns blocks to pages
   - Splits blocks at line boundaries when needed
   - Respects widow/orphan control
   - Handles page breaks

5. **DOM Painter**: Renders content to page containers:
   - Creates page DOM elements
   - Paints block fragments at correct positions
   - Handles headers and footers

6. **Input Bridge**: Forwards user input from the visible pages to the hidden ProseMirror editor:
   - Maps click positions to ProseMirror positions
   - Handles keyboard events
   - Manages focus

7. **Selection Overlay**: Renders caret and selection highlighting as a separate layer above the content.

### Page Configuration

```typescript
interface PageConfig {
  width: number;   // Page width in pixels (at 96 DPI)
  height: number;  // Page height in pixels
  margins: {
    top: number;
    bottom: number;
    left: number;
    right: number;
    header?: number;  // Header area offset from top
    footer?: number;  // Footer area offset from bottom
  };
}

// Standard page sizes (96 DPI)
const PAGE_SIZES = {
  LETTER: { w: 816, h: 1056 },   // 8.5" x 11"
  LEGAL: { w: 816, h: 1344 },    // 8.5" x 14"
  A4: { w: 794, h: 1123 },       // 210mm x 297mm
  // ... more sizes
};
```

### Headers and Footers

Headers and footers support dynamic fields:

```typescript
const header: HeaderFooterContent = {
  blocks: [
    {
      type: 'paragraph',
      content: [
        { type: 'text', text: 'Document Title - Page ' },
        { type: 'dynamicField', fieldType: 'pageNumber' },
        { type: 'text', text: ' of ' },
        { type: 'dynamicField', fieldType: 'totalPages' },
      ],
      alignment: 'center',
    },
  ],
  differentFirstPage: true,
  firstPageBlocks: [], // Empty header on first page
};
```

### Design Principles

1. **Single Source of Truth**: ProseMirror manages all document state. The visual pages are a derived view.

2. **Line-Level Pagination**: Unlike block-level pagination, content is broken at natural line boundaries for proper flow.

3. **Separation of Concerns**: 
   - Core handles document structure
   - React handles rendering and interaction
   - Layout engine handles pagination

4. **Performance-First**:
   - Measurements are cached
   - Only visible pages are painted
   - Selection overlay is efficient

5. **Print Fidelity**: What you see is what you print - true WYSIWYG editing.

## Basic Usage

```tsx
import { DocumentImpl } from '@b32nio/spindle-docs-core';
import { DocumentProvider, DocumentEditor } from '@b32nio/spindle-docs-react';

// Create document
const doc = new DocumentImpl('doc_1', 'My Document');

function App() {
  return (
    <DocumentProvider document={doc}>
      <DocumentEditor
        showToolbar={true}
        showRuler={true}
      />
    </DocumentProvider>
  );
}
```

### Using TrueLayoutEditor Directly

```tsx
import { TrueLayoutEditor } from '@b32nio/spindle-docs-react';
import type { Block } from '@b32nio/spindle-docs-core';

function Editor({ blocks }: { blocks: Block[] }) {
  const handleDocChange = (newBlocks: Block[]) => {
    // Save blocks
  };
  
  return (
    <TrueLayoutEditor
      initialBlocks={blocks}
      pageConfig={{
        width: 816,
        height: 1056,
        margins: { top: 96, bottom: 96, left: 96, right: 96 },
      }}
      zoom={1}
      editable={true}
      onDocChange={handleDocChange}
      pageGap={24}
      minLinesAtBreak={2}
    />
  );
}
```

## Extension Points

### Custom Block Rendering

The DOM Painter can be extended to handle custom block types.

### Custom Commands

Add new ProseMirror commands for custom formatting or block types.

### Custom Dynamic Fields

Extend the dynamic field system for headers/footers with custom field types.

## Performance Considerations

- **Lazy Measurement**: Blocks are only measured when content changes
- **Virtual Rendering**: Only visible pages are fully rendered
- **Debounced Layout**: Layout computation is debounced during rapid typing
- **Style Pooling**: Shared styles reduce memory usage

## Related Documentation

- [Component Reference](./components.md) - Complete guide to React components
- [Data Structures](./data-structures.md) - Type definitions and interfaces
- [Contributing](../sheets/contributing/extending.md) - Guide for extending the library
