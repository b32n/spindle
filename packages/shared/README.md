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

- **`@b32nio/spindle-shared/export/csv`**, **`/export/xlsx`**, **`/export/docx`**:
  format-level readers/writers (`exportCsv`/`importCsv`, `exportXlsx`/`importXlsx`,
  `exportDocx`/`importDocx`) independent of any editor model. The editor cores
  adapt their models onto these; use them directly for headless conversion.
  `xlsx` and `docx` pull in `exceljs` / `docx` + `jszip` + `fast-xml-parser`, so
  import them only where needed.

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

