# @b32nio/spindle-shared

Shared utilities for Spindle packages, used across the docs, sheets, and slides
cores and their React layers.

## Exports

- **EventEmitter**: Generic event emitter for event-driven architectures
- **Collaboration**: Yjs-based collaboration types and an in-memory
  `CollabProvider` (the WebSocket transport lives in
  `@b32nio/spindle-transport-websocket`)
- **`@b32nio/spindle-shared/react`**: shared React UI — notably
  `ResponsiveToolbar` (overflow/menu toolbar used by all three editors)

- **`@b32nio/spindle-shared/export/csv`**: CSV reader/writer (`exportCsv`/`importCsv`,
  `detectCsvLoss`), dependency-free.
- **`@b32nio/spindle-shared/export/xlsx`**, **`/export/docx`**: the format types and
  import-loss helpers, plus **deprecated** async shims (`exportXlsx`/`importXlsx`,
  `exportDocx`/`importDocx`, `detect*Loss`) that load `@b32nio/spindle-io` on first
  call. The implementations (and `exceljs` / `docx` / `jszip` / `fast-xml-parser`) moved
  to [`@b32nio/spindle-io`](../io/README.md); install it and import from
  `@b32nio/spindle-io/xlsx` / `/docx` instead. It is an optional peer dependency here.

## Usage

```typescript
import { EventEmitter } from '@b32nio/spindle-shared';

// Create an event emitter with specific event types
type MyEventTypes = 'change' | 'update' | 'delete';
const emitter = new EventEmitter<MyEventTypes>();

emitter.on('change', (data) => {
  console.log('Change event:', data.payload);
});

emitter.emit('change', { some: 'data' });
```

