// @b32nio/spindle-transport-websocket
//
// WebSocketProvider — a CollabProvider implementation that relays
// 'doc' / 'awareness' payloads over a WebSocket to a peer-relay server
// (see examples/collab-server for a reference implementation).
//
// Wire shape:
//   - URL: <baseUrl>/<roomId>
//   - Each message: 1 byte channel tag + opaque payload
//       tag = 0 → 'doc'
//       tag = 1 → 'awareness'
//
// The server doesn't interpret payloads — it just broadcasts a peer's
// message to every other peer in the same room. That's enough for
// y-protocols/sync + y-protocols/awareness to do their thing.

export { WebSocketProvider, type WebSocketProviderOptions } from './provider';
