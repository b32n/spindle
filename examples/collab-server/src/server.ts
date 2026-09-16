// Reference WebSocket relay server for Spindle collaboration.
//
// Wire protocol (matches @b32nio/spindle-transport-websocket):
//   Connect:  ws://host:PORT/<roomId>
//   Each frame: 1 byte channel tag + opaque payload.
//     tag 0 = doc        (durable — logged and replayed to newcomers)
//     tag 1 = awareness   (ephemeral — fanned out only)
//     tag 2 = sync        (server → client control: "replay complete")
//   The server reads only the 1-byte tag; the payload stays opaque, so this
//   remains zero-knowledge / E2EE-friendly.
//
// Why a log + replay (not a bare relay): seeding a CRDT is a write, and if
// every peer seeds independently the document duplicates. This server is the
// room's authority — on connect it replays the room's doc log to the newcomer
// and then sends a `sync` frame, so the client knows whether the room already
// had state (join, don't seed) or was empty (creator, seed). The same replay
// runs on reconnect, so a peer that dropped re-converges for free.
//
// This reference keeps the log in memory, dedups identical re-pushed frames,
// and evicts a room after it has been idle (peerless) for a grace window.
// A production deployment should swap the in-memory `rooms` map for durable,
// content-agnostic storage (any KV/blob store) and add auth on the upgrade —
// the protocol above is unchanged.

import { createServer, type Server } from 'node:http';
import { WebSocketServer, type WebSocket, type RawData } from 'ws';

const CHANNEL_DOC = 0;
const CHANNEL_SYNC = 2;

interface Room {
  peers: Set<WebSocket>;
  /** Ordered log of every doc frame, replayed verbatim to each newcomer. */
  docLog: Buffer[];
  /** base64 of every logged frame, so an identical re-push isn't appended twice. */
  logged: Set<string>;
  /** Pending eviction timer once the room goes idle (empty). */
  idleTimer?: ReturnType<typeof setTimeout>;
}

function roomIdFromUrl(url: string | undefined): string {
  if (!url) return '';
  const path = url.split('?')[0].replace(/^\/+/, '');
  try {
    return decodeURIComponent(path);
  } catch {
    return path;
  }
}

function toBuffer(raw: RawData): Buffer {
  if (Buffer.isBuffer(raw)) return raw;
  if (Array.isArray(raw)) return Buffer.concat(raw);
  return Buffer.from(raw as ArrayBuffer);
}

export interface CollabServerOptions {
  /** Port to listen on. Defaults to $PORT or 1234. Pass 0 for an ephemeral port. */
  port?: number;
  /** Called once the server is listening (receives the resolved port). */
  onListening?: (port: number) => void;
  /**
   * How long to keep an idle (peerless) room's log before evicting it, in ms.
   * A brief drop/reconnect within this window is still replayed its state; a
   * truly abandoned room is reclaimed instead of leaking forever. Default 5 min.
   */
  emptyRoomTtlMs?: number;
}

/**
 * Create and start a relay server. Returns the underlying http.Server so a
 * caller (or a test) can read `.address()` and `.close()` it.
 */
export function createCollabServer(options: CollabServerOptions = {}): Server {
  const rooms = new Map<string, Room>();
  const emptyRoomTtlMs = options.emptyRoomTtlMs ?? 5 * 60_000;

  const httpServer = createServer((_req, res) => {
    res.writeHead(200, { 'content-type': 'text/plain' });
    res.end('spindle-collab-server: connect via WebSocket at /<roomId>\n');
  });

  const wss = new WebSocketServer({ server: httpServer });

  wss.on('connection', (ws, req) => {
    const roomId = roomIdFromUrl(req.url);
    if (!roomId) {
      ws.close(1008, 'roomId required (use ws://host/<roomId>)');
      return;
    }

    let room = rooms.get(roomId);
    if (!room) {
      room = { peers: new Set(), docLog: [], logged: new Set() };
      rooms.set(roomId, room);
    }
    // A returning peer cancels a pending eviction so its state isn't reclaimed.
    if (room.idleTimer) {
      clearTimeout(room.idleTimer);
      room.idleTimer = undefined;
    }
    room.peers.add(ws);

    // Replay the room's existing doc state, then signal that the newcomer is
    // synced. connect() on the client resolves on this frame.
    for (const frame of room.docLog) {
      if (ws.readyState === ws.OPEN) ws.send(frame, { binary: true });
    }
    if (ws.readyState === ws.OPEN) ws.send(Buffer.from([CHANNEL_SYNC]), { binary: true });

    ws.on('message', (raw, isBinary) => {
      const frame = toBuffer(raw);
      if (frame.length < 1) return;
      // Log doc frames (durable); awareness is fanned out but never stored.
      // Dedup identical frames: a persistence-enabled client re-pushes its full
      // state on every attach, which is byte-identical when nothing changed —
      // appending each copy would grow the log without bound. Comparing bytes is
      // content-agnostic, so this stays E2EE-friendly.
      if (frame[0] === CHANNEL_DOC) {
        const copy = Buffer.from(frame);
        const key = copy.toString('base64');
        if (!room!.logged.has(key)) {
          room!.logged.add(key);
          room!.docLog.push(copy);
        }
      }
      for (const peer of room!.peers) {
        if (peer === ws || peer.readyState !== peer.OPEN) continue;
        try {
          peer.send(raw, { binary: isBinary });
        } catch (err) {
          // eslint-disable-next-line no-console
          console.warn('[collab-server] send failed:', err);
        }
      }
    });

    ws.on('close', () => {
      room!.peers.delete(ws);
      // Keep the room's doc log when it goes idle so a peer that briefly drops
      // (or the last peer leaving and soon rejoining) is replayed its state —
      // but only for a grace window, then evict, so abandoned rooms don't leak
      // forever. A production deployment should move `rooms` to durable storage
      // with its own eviction + compaction policy. Payloads stay opaque (the
      // dedup above compares bytes only), so this remains E2EE-friendly.
      if (room!.peers.size === 0 && !room!.idleTimer) {
        room!.idleTimer = setTimeout(() => {
          if (room!.peers.size === 0) rooms.delete(roomId);
        }, emptyRoomTtlMs);
        // Don't keep the process alive just to evict an idle room.
        room!.idleTimer.unref?.();
      }
    });

    ws.on('error', (err) => {
      // eslint-disable-next-line no-console
      console.warn('[collab-server] socket error:', err.message);
    });
  });

  const port = options.port ?? Number(process.env.PORT ?? 1234);
  httpServer.listen(port, () => {
    const resolved = (httpServer.address() as { port: number } | null)?.port ?? port;
    options.onListening?.(resolved);
    // eslint-disable-next-line no-console
    console.log(`[collab-server] listening on ws://localhost:${resolved}/<roomId>`);
  });

  return httpServer;
}

// Start when run directly (npm start / tsx src/server.ts).
createCollabServer();
