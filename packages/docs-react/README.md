# @b32nio/spindle-docs-react

React components for Spindle Docs with paginated document rendering.

## Features

- **Paginated rendering**: True page-based layout with headers, footers, and page breaks
- **ProseMirror-based editing**: Rich text editing with formatting toolbar
- **Line-level pagination**: Content breaks at line boundaries, not block boundaries
- **Header/footer support**: Dynamic fields for page numbers, total pages, document title
- **Tables**: Basic table support with cell editing
- **Ruler**: Interactive margin adjustment via draggable ruler

## Components

### DocumentEditor (main component)
The main document editor component. Provides:
- Formatting toolbar
- Horizontal and vertical rulers
- Paginated content display
- Header/footer editing

### TrueLayoutEditor
Lower-level component for custom implementations. Handles:
- Page layout computation
- DOM-based content rendering
- Selection overlay
- Input bridging between visible content and hidden ProseMirror

## Usage

```typescript
import { DocumentProvider, DocumentEditor } from '@b32nio/spindle-docs-react';
import { DocumentImpl } from '@b32nio/spindle-docs-core';

const doc = new DocumentImpl('doc_1', 'My Document');

function App() {
  return (
    <DocumentProvider document={doc}>
      <DocumentEditor 
        width={window.innerWidth} 
        height={window.innerHeight}
        showToolbar={true}
        showRuler={true}
      />
    </DocumentProvider>
  );
}
```

## Architecture

The rendering uses a hidden ProseMirror editor with visible paginated output:

```
DocumentEditor (React)
├── Toolbar (DOM)
├── Ruler (DOM)
├── TrueLayoutEditor
│   ├── Hidden ProseMirror (handles editing)
│   ├── Page containers (visible, paginated)
│   │   ├── Header
│   │   ├── Content fragments
│   │   └── Footer
│   ├── Selection overlay (caret + selection rects)
│   └── Input bridge (forwards events to hidden editor)
└── PageSetupModal
```

## Hooks

```typescript
import { useDocument, useSections, useHistory } from '@b32nio/spindle-docs-react';

// Access document model
const { document, zoom, setZoom } = useDocument();

// Get sections
const sections = useSections();

// Undo/redo
const { canUndo, canRedo, undo, redo } = useHistory();
```

## Limitations

- No spell-check (browser spellcheck is disabled)
- No track changes / suggestions mode (comment threads are supported)
- Tables support column & row resize but not cell merge/split
- No find & replace, word count, or export (PDF/DOCX)
- Multi-section documents render only the first section in the UI

Real-time collaboration, comments, and image upload (file / drag-drop / URL)
**are** supported. See
[`documentation/docs/TODO.md`](../../documentation/docs/TODO.md) for the full
remaining-work list.
