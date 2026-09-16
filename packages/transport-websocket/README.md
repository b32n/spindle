# @b32nio/spindle-transport-websocket

A WebSocket-backed [`CollabProvider`](https://github.com/b32n/spindle) for
Spindle real-time collaboration. Plug it into a `WorkbookImpl`, `DocumentImpl`,
or `DeckImpl` via `attachCollab` to sync edits, presence, and remote cursors
across clients over a WebSocket relay.

Part of [Spindle](https://spindle.b32n.io) — open-source spreadsheet,
document, and slide editing libraries for React.

## Installation

```bash
npm install @b32nio/spindle-transport-websocket
```

## Usage

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

The `roomId` is appended to the base `url` as a path segment. The provider
reconnects automatically with exponential backoff, and queued edits flush once
the socket reopens. Because payloads are opaque byte arrays, wrapping them for
end-to-end encryption is a thin layer on top.

### Options

| Option | Default | Description |
| --- | --- | --- |
| `url` | — | Base WebSocket URL, no trailing slash (e.g. `wss://collab.example.com`). |
| `WebSocketImpl` | global `WebSocket` | Custom `WebSocket` constructor (useful in Node). |
| `minReconnectDelayMs` | `250` | Minimum reconnect backoff delay. |
| `maxReconnectDelayMs` | `10000` | Maximum reconnect backoff delay. |

A reference relay server lives in
[`examples/collab-server`](https://github.com/b32n/spindle/tree/master/examples/collab-server).
See the [Collaboration Guide](https://github.com/b32n/spindle/blob/master/documentation/collaboration.md)
for transports, offline persistence, and an end-to-end encryption recipe.

## License

MIT
