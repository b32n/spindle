# Spindle

High-performance spreadsheet, docs and slides libraries for React, optimized for performance and bundle size.

## Inspiration

For the past decade, I've dreamed of building my own encrypted drive application—primarily for personal use, but one that others could benefit from too. When I finally began development, I discovered a significant gap in the open-source ecosystem: there was no comprehensive library offering full-featured editors for spreadsheets, documents, and presentations.

Existing solutions were either:
- **Heavy and bloated** with unnecessary features
- **Incomplete** supporting only one document type
- **Proprietary or restrictive** with mixed licensing models

This inspired me to create **Spindle**—a forever-free, open-source library that delivers the full editing experience you'd expect from commercial suites like Google Workspace or Microsoft Office, but built for developers who value performance, flexibility, and complete control over their tools.

## Quick Start

### Installation

Add the packages to your React project:

```bash
npm install @b32nio/spindle-sheets-core @b32nio/spindle-sheets-react
```

### Basic Usage - Spreadsheets

Here's how to add a spreadsheet to your React application:

```tsx
import React, { useState } from 'react';
import { WorkbookProvider, WorkbookCanvas } from '@b32nio/spindle-sheets-react';
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';

function MySpreadsheet() {
  // 1. Create a workbook instance
  const [workbook] = useState(() => {
    const wb = new WorkbookImpl('my-workbook', 'My Spreadsheet');

    // Optional: Pre-populate with data
    const initialData = {
      id: 'my-workbook',
      name: 'My Spreadsheet',
      activeSheetId: 'sheet_1',
      sheets: [{
        id: 'sheet_1',
        name: 'Sheet1',
        cells: [
          { key: '0:0', cell: { value: 'Hello' } },
          { key: '0:1', cell: { value: 'World' } },
        ],
        config: {
          defaultRowHeight: 20,
          defaultColWidth: 100,
        },
        rowCount: 1000,
        colCount: 100,
      }],
      selection: {
        ranges: [],
        activeCell: { row: 0, col: 0 },
      },
    };

    wb.setData(initialData);
    return wb;
  });

  // Example: Access and save workbook data
  const saveToDb = () => {
    const workbookData = workbook.getData();
    // Send to your backend API
    console.log('Saving workbook data:', workbookData);
    // Example: fetch('/api/workbooks', { method: 'POST', body: JSON.stringify(workbookData) })
  };

  return (
    <div style={{ width: '100%', height: '500px' }}>
      <button onClick={saveToDb} style={{ marginBottom: '10px', padding: '8px 16px' }}>
        Save Workbook
      </button>

      {/* 2. Wrap your app with WorkbookProvider */}
      <WorkbookProvider workbook={workbook}>
        {/* 3. Render the WorkbookCanvas component */}
        <WorkbookCanvas
          width={800}
          height={400}
          rowHeight={20}
          colWidth={100}
        />
      </WorkbookProvider>
    </div>
  );
}

export default MySpreadsheet;
```

### Key Components - Sheets

- **`WorkbookImpl`**: The core spreadsheet engine that manages data and state
- **`WorkbookProvider`**: React context provider that manages workbook state
- **`WorkbookCanvas`**: The main UI component that renders the spreadsheet

### Basic Usage - Documents

Here's how to add a document editor to your React application:

```bash
npm install @b32nio/spindle-docs-core @b32nio/spindle-docs-react
```

```tsx
import React, { useState } from 'react';
import { DocumentImpl, type DocumentData } from '@b32nio/spindle-docs-core';
import { DocumentProvider, DocumentEditor } from '@b32nio/spindle-docs-react';

// Document data in JSON format (can be loaded from backend/database)
const initialDocumentData: DocumentData = {
  id: 'my-document',
  title: 'My Document',
  defaultPageConfig: {
    size: { w: 816, h: 1056 }, // Letter size at 96 DPI
    margins: { top: 96, right: 96, bottom: 96, left: 96 },
    orientation: 'portrait',
  },
  textStylePool: {
    'style_bold': { bold: true },
  },
  paragraphStylePool: {},
  sections: [
    {
      id: 'section_1',
      pageConfig: {
        size: { w: 816, h: 1056 },
        margins: { top: 96, right: 96, bottom: 96, left: 96 },
        orientation: 'portrait',
      },
      blocks: [
        {
          id: 'block_1',
          type: 'heading',
          level: 1,
          content: [{ type: 'text', text: 'Welcome to Spindle Docs' }],
        },
        {
          id: 'block_2',
          type: 'paragraph',
          content: [
            { type: 'text', text: 'This is a ' },
            { type: 'text', text: 'rich text', styleId: 'style_bold' },
            { type: 'text', text: ' document editor.' },
          ],
        },
      ],
    },
  ],
  createdAt: new Date().toISOString(),
  updatedAt: new Date().toISOString(),
};

function MyDocumentEditor() {
  // 1. Create a document instance and load data
  const [doc] = useState(() => {
    const document = new DocumentImpl();
    document.setData(initialDocumentData);
    return document;
  });

  // Example: Access and save document data
  const saveToDb = () => {
    const documentData = doc.getData();
    // Send to your backend API
    console.log('Saving document data:', documentData);
    // Example: fetch('/api/documents', { method: 'POST', body: JSON.stringify(documentData) })
  };

  return (
    <div style={{ width: '100%', height: '800px' }}>
      <button onClick={saveToDb} style={{ marginBottom: '10px', padding: '8px 16px' }}>
        Save Document
      </button>

      {/* 2. Wrap your app with DocumentProvider */}
      <DocumentProvider document={doc}>
        {/* 3. Render the DocumentEditor component */}
        <DocumentEditor
          showToolbar={true}
          showRuler={true}
        />
      </DocumentProvider>
    </div>
  );
}

export default MyDocumentEditor;
```

### Key Components - Docs

- **`DocumentImpl`**: The core document engine that manages content and state
- **`DocumentProvider`**: React context provider that manages document state
- **`DocumentEditor`**: The main UI component with toolbar, rulers, and paginated editing

### Features Included

**Spreadsheets:**
- ✅ Virtual scrolling for performance
- ✅ Cell editing with keyboard navigation
- ✅ Formula engine — 146 functions, dynamic arrays (spill), and topological recalculation
- ✅ Formula autocomplete with parameter help
- ✅ Customizable dimensions and styling
- ✅ Sparse data storage (only stores non-empty cells)

**Documents:**
- ✅ True page layout with line-level pagination
- ✅ Rich text formatting (bold, italic, fonts, colors)
- ✅ Headers & footers with dynamic fields (page numbers, date)
- ✅ Tables with cell formatting
- ✅ Images (block and inline)
- ✅ Undo/redo history

**Slides:**
- ✅ Element model: text, shapes, images, lines/connectors, and tables
- ✅ Drag/resize/rotate at 60fps with smart-guide snapping
- ✅ Themes & layouts with placeholder materialization
- ✅ Present mode with keyboard nav and presenter view
- ✅ Element-anchored comments

**Collaboration:**
- ✅ Real-time multi-user editing for docs, sheets, and slides (Yjs CRDT)
- ✅ Live remote cursors, selections, and presence
- ✅ Pluggable transport — ships an in-memory and a WebSocket provider
- ✅ Offline persistence (IndexedDB) — edits survive refresh/offline
- ✅ Encryption-ready — opaque byte payloads wrap cleanly for E2EE

## 🤝 Real-Time Collaboration

Both documents and spreadsheets sync live between users. The library owns
a Yjs CRDT internally; you supply a transport (`CollabProvider`) and call
`attachCollab`:

```ts
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';
import { WebSocketProvider } from '@b32nio/spindle-transport-websocket';

const workbook = new WorkbookImpl('wb_1', 'Quarterly Plan');
workbook.setData(savedJson);

const provider = new WebSocketProvider({ url: 'wss://collab.example.com' });

await workbook.attachCollab(
  provider,
  { userId: 'u_42', displayName: 'Alice', color: '#4ecdc4' },
  { roomId: 'quarterly-plan', persistenceKey: 'wb:quarterly-plan' },
);
```

Documents work identically (`document.attachCollab(...)`). The React
editors render remote cursors and selections automatically — no extra
wiring. Because `CollabProvider` payloads are opaque byte arrays,
end-to-end encryption is a simple wrapping layer.

See the **[Collaboration Guide](./documentation/collaboration.md)** for
transports, offline persistence, connection status, and a complete E2EE
recipe. A reference relay server lives in
[`examples/collab-server`](./examples/collab-server/).

## Development

Help is highly appreciated in improving performance, adding features and also porting to other frontend frameworks.

P.S. I am not a react pro in any way. Any support in improving code is highly welcome. 

```bash
# Install dependencies
npm install

# Build all packages
npm run build

# Run tests
npm run test

# Type check
npm run type-check

# Lint
npm run lint
```

