# @b32nio/spindle-docs-core

Core document editor engine for Spindle Docs. Zero React dependencies.

## Features

- Document model with blocks (paragraphs, headings, lists, tables, images)
- ProseMirror schema and state management
- Text and paragraph style pooling for memory efficiency
- Event system with batching
- Undo/redo history
- TypeScript strict mode

## Usage

```typescript
import { DocumentImpl, type Block } from '@b32nio/spindle-docs-core';

const doc = new DocumentImpl('doc_1', 'My Document');

// Get sections and blocks
const sections = doc.getSections();
const blocks = sections[0].blocks;

// Update page config
doc.setSectionPageConfig(sections[0].id, {
  size: { w: 816, h: 1056 },
  margins: { top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 },
  orientation: 'portrait',
});

// Serialize/deserialize
const data = doc.getData();
doc.setData(data);
```

## Block Types

- `paragraph` - Basic text paragraph
- `heading` - Heading levels 1-6
- `list-item` - Bulleted or numbered list items
- `table` - Tables with rows and cells
- `image` - Images with optional captions

## ProseMirror Integration

The package includes ProseMirror schema and utilities for rich text editing:

```typescript
import { 
  docsSchema, 
  createPlugins, 
  blocksToPmDoc, 
  proseMirrorToDocument 
} from '@b32nio/spindle-docs-core';

// Convert blocks to ProseMirror document
const pmDoc = blocksToPmDoc(blocks, schema);

// Convert ProseMirror document back to blocks
const docModel = proseMirrorToDocument(pmNode, existingDoc);
```

## Limitations

- Single-section editing: multi-section documents live in the data model, but
  the editor UI renders only the first section (and collaboration is
  single-section for now)
- No export/import (PDF, DOCX) and no find & replace

Real-time collaboration (Yjs, via `attachCollab` / the `/collab` subpath) and
comment threads **are** supported. See
[`documentation/docs/TODO.md`](../../documentation/docs/TODO.md) for the full
remaining-work list.
