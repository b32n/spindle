# spindle-collab-server

Minimal WebSocket relay server for `@b32nio/spindle-transport-websocket`. Acts as
a dumb broadcast hub: it never inspects payloads, only groups connected
sockets by room id (the URL path) and forwards each peer's messages to the
others in the same room.

## Run it

```bash
npm install
npm run start
# → listening on ws://localhost:1234/<roomId>
```

Set `PORT=...` to use a different port.

## Wire protocol

Clients connect with `ws://host:PORT/<roomId>`. Each WebSocket message is
binary:

```
| 1 byte channel tag | opaque payload |
```

* `tag = 0` → the `doc` channel (y-protocols/sync frames produced by the
  Yjs binding inside `@b32nio/spindle-docs-core/collab` or `.../sheets-core/collab`).
* `tag = 1` → the `awareness` channel (y-protocols/awareness updates).

The server doesn't decode either; it just rebroadcasts the frame.

## What's missing for production

* **Auth.** The server accepts any connection. Add auth on the HTTP upgrade
  request (e.g., a JWT in a header or query string).
* **Persistence.** When all peers leave a room, the in-memory state is lost
  and a fresh peer joining gets an empty Y.Doc. Plug in y-websocket's
  filesystem persistence or any KV store keyed by room id.
* **TLS.** Run behind a reverse proxy or use `https`/`wss` directly.
* **Rate limiting / size caps.** A misbehaving peer can flood a room.
