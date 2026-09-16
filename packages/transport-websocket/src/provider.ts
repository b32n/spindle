// WebSocketProvider: a CollabProvider over a WebSocket relay server.
//
// Contract (see CollabProvider.connect): connect() must not resolve until the
// room's existing 'doc' state has been delivered to the doc handlers, so a
// binding can decide whether to seed. The server satisfies this by replaying
// the room's opaque 'doc' log to a newcomer and then sending a one-byte
// CHANNEL_SYNC control frame; connect() resolves on that frame. The replay
// also runs on every reconnect, so a peer that was offline re-converges for
// free (applying replayed updates is idempotent).
//
// Reconnection: on an unintentional close, the provider reopens with
// exponential backoff. Handlers are guarded so a superseded socket's late
// events can never mutate state that now belongs to a newer socket.

import type {
  CollabChannel,
  CollabMessageHandler,
  CollabProvider,
  CollabStatus,
  CollabStatusHandler,
} from '@b32nio/spindle-shared';

const CHANNEL_DOC = 0;
const CHANNEL_AWARENESS = 1;
// Server → client control frame: "the room's state has been replayed to you".
const CHANNEL_SYNC = 2;

function channelToByte(channel: CollabChannel): number {
  return channel === 'doc' ? CHANNEL_DOC : CHANNEL_AWARENESS;
}

function byteToChannel(byte: number): CollabChannel | null {
  if (byte === CHANNEL_DOC) return 'doc';
  if (byte === CHANNEL_AWARENESS) return 'awareness';
  return null;
}

export interface WebSocketProviderOptions {
  /**
   * Base WebSocket URL (no trailing slash). The roomId passed to connect()
   * is appended as a path segment. Example: `ws://localhost:1234`.
   */
  url: string;
  /** Optional WebSocket constructor (for Node tests). Defaults to global. */
  WebSocketImpl?: typeof WebSocket;
  /** Min reconnect delay in ms. Default 250. */
  minReconnectDelayMs?: number;
  /** Max reconnect delay in ms. Default 10_000. */
  maxReconnectDelayMs?: number;
  /**
   * How long to wait after the socket opens for the server's CHANNEL_SYNC
   * frame before resolving connect() anyway. A conforming server replies
   * immediately, so this only matters against a relay that doesn't implement
   * replay — there the provider degrades to best-effort rather than hanging.
   * Default 4000.
   */
  syncTimeoutMs?: number;
}

export class WebSocketProvider implements CollabProvider {
  private readonly url: string;
  private readonly Ctor: typeof WebSocket;
  private readonly minDelay: number;
  private readonly maxDelay: number;
  private readonly syncTimeoutMs: number;

  private socket: WebSocket | null = null;
  private roomId: string | null = null;
  /** True between intentional disconnect() and a fresh connect(). */
  private intentionallyClosed = false;
  /** Outbound messages queued while the socket is (re-)connecting. */
  private pendingOutbound: Uint8Array[] = [];
  /** Channel → handler set. Survives disconnects. */
  private subscribers: Map<CollabChannel, Set<CollabMessageHandler>> = new Map();
  /** Reconnect bookkeeping. */
  private reconnectAttempt = 0;
  private reconnectTimer: ReturnType<typeof setTimeout> | null = null;
  /** Promise resolver for the in-flight connect(). */
  private connectResolve: (() => void) | null = null;
  private connectReject: ((err: unknown) => void) | null = null;
  /** True once the current connect() has opened at least one socket. */
  private opened = false;
  /** Fallback timer that resolves connect() if no CHANNEL_SYNC frame arrives. */
  private syncTimer: ReturnType<typeof setTimeout> | null = null;
  /** Connection status + subscribers. */
  private status: CollabStatus = 'offline';
  private statusHandlers: Set<CollabStatusHandler> = new Set();

  constructor(options: WebSocketProviderOptions) {
    this.url = options.url.replace(/\/+$/, '');
    this.Ctor =
      options.WebSocketImpl ??
      (typeof WebSocket !== 'undefined' ? WebSocket : (undefined as unknown as typeof WebSocket));
    if (!this.Ctor) {
      throw new Error(
        'WebSocketProvider: no global WebSocket available — pass options.WebSocketImpl',
      );
    }
    this.minDelay = options.minReconnectDelayMs ?? 250;
    this.maxDelay = options.maxReconnectDelayMs ?? 10_000;
    this.syncTimeoutMs = options.syncTimeoutMs ?? 4000;
  }

  connect(roomId: string): Promise<void> {
    if (this.roomId && this.roomId !== roomId) {
      // Already connected to a different room — disconnect first.
      this.disconnect();
    }
    this.roomId = roomId;
    this.intentionallyClosed = false;
    this.opened = false;

    // Cancel any pending reconnect/sync timers from a prior connection to this
    // same room. Otherwise a backoff timer armed by an earlier drop fires after
    // openSocket() below and tears down the just-established socket to open yet
    // another — a spurious disconnect/reconnect blip right after connect().
    if (this.reconnectTimer !== null) {
      clearTimeout(this.reconnectTimer);
      this.reconnectTimer = null;
    }
    this.clearSyncTimer();

    // Supersede any still-pending connect() so its promise can't hang forever.
    if (this.connectReject) {
      this.connectReject(new Error('WebSocketProvider: connect() superseded by a newer connect()'));
      this.connectResolve = null;
      this.connectReject = null;
    }

    return new Promise<void>((resolve, reject) => {
      this.connectResolve = resolve;
      this.connectReject = reject;
      this.openSocket();
    });
  }

  disconnect(): void {
    this.intentionallyClosed = true;
    this.roomId = null;
    // Settle a still-pending connect() so an awaiting caller (e.g. attachCollab
    // during a React unmount mid-connect) doesn't hang forever — with the socket
    // detached below there is no onclose left to reject it and no timer to
    // resolve it.
    if (this.connectReject) {
      this.connectReject(new Error('WebSocketProvider: connect() aborted by disconnect()'));
      this.connectResolve = null;
      this.connectReject = null;
    }
    if (this.reconnectTimer !== null) {
      clearTimeout(this.reconnectTimer);
      this.reconnectTimer = null;
    }
    this.clearSyncTimer();
    if (this.socket) {
      this.detachSocket(this.socket);
      try {
        this.socket.close();
      } catch {
        /* ignore */
      }
      this.socket = null;
    }
    this.pendingOutbound = [];
    this.setStatus('offline');
  }

  send(channel: CollabChannel, payload: Uint8Array): void {
    const framed = new Uint8Array(payload.length + 1);
    framed[0] = channelToByte(channel);
    framed.set(payload, 1);
    if (this.socket && this.socket.readyState === 1 /* OPEN */) {
      this.socket.send(framed);
    } else {
      // Buffer until the socket opens; flushed in onopen.
      this.pendingOutbound.push(framed);
    }
  }

  onMessage(channel: CollabChannel, handler: CollabMessageHandler): () => void {
    let set = this.subscribers.get(channel);
    if (!set) {
      set = new Set();
      this.subscribers.set(channel, set);
    }
    set.add(handler);
    return () => {
      set!.delete(handler);
    };
  }

  getStatus(): CollabStatus {
    return this.status;
  }

  onStatusChange(handler: CollabStatusHandler): () => void {
    this.statusHandlers.add(handler);
    return () => {
      this.statusHandlers.delete(handler);
    };
  }

  // --- Internals ----------------------------------------------------------

  private setStatus(status: CollabStatus): void {
    if (this.status === status) return;
    this.status = status;
    for (const h of this.statusHandlers) {
      try {
        h(status);
      } catch (err) {
        // eslint-disable-next-line no-console
        console.error('[WebSocketProvider] status handler threw:', err);
      }
    }
  }

  /** Detach every handler from a socket so its late events become no-ops. */
  private detachSocket(ws: WebSocket): void {
    ws.onopen = null;
    ws.onmessage = null;
    ws.onclose = null;
    // Keep a no-op error handler rather than nulling it: closing a socket that
    // is still CONNECTING emits an 'error', which would otherwise be unhandled.
    ws.onerror = () => {};
  }

  private clearSyncTimer(): void {
    if (this.syncTimer !== null) {
      clearTimeout(this.syncTimer);
      this.syncTimer = null;
    }
  }

  /** Resolve the in-flight connect(), if any. Called on the sync frame. */
  private resolveConnect(): void {
    this.clearSyncTimer();
    if (this.connectResolve) {
      const r = this.connectResolve;
      this.connectResolve = null;
      this.connectReject = null;
      r();
    }
  }

  private openSocket(): void {
    if (!this.roomId) return;
    // Retire any previous socket so its handlers can't fight the new one, and
    // close it so a superseded/reconnected socket doesn't linger open.
    if (this.socket) {
      this.detachSocket(this.socket);
      try {
        this.socket.close();
      } catch {
        /* ignore */
      }
    }

    const url = `${this.url}/${encodeURIComponent(this.roomId)}`;
    const ws = new this.Ctor(url);
    ws.binaryType = 'arraybuffer';
    this.socket = ws;
    this.setStatus('connecting');

    ws.onopen = () => {
      if (this.socket !== ws) return; // superseded
      this.reconnectAttempt = 0;
      this.setStatus('connected');
      // The socket opened, so a later drop (before the sync frame) should
      // reconnect, not reject connect(). Mark it opened rather than nulling
      // connectReject, so a superseding connect() can still reject the pending
      // promise instead of leaving it to hang.
      this.opened = true;
      // Flush any messages queued while the socket was opening.
      for (const m of this.pendingOutbound) {
        try {
          ws.send(m);
        } catch {
          // If a single send fails, the socket will close and we'll reconnect
          // — drop the rest to avoid noisy errors.
          break;
        }
      }
      this.pendingOutbound = [];
      // Do NOT resolve connect() yet — wait for the server to replay the room
      // and send CHANNEL_SYNC. Arm a fallback so a non-conforming relay can't
      // hang the caller forever.
      if (this.connectResolve && this.syncTimer === null) {
        this.syncTimer = setTimeout(() => {
          this.syncTimer = null;
          this.resolveConnect();
        }, this.syncTimeoutMs);
      }
    };

    ws.onmessage = (event) => {
      if (this.socket !== ws) return; // superseded
      const data = event.data;
      let bytes: Uint8Array | null = null;
      if (data instanceof ArrayBuffer) bytes = new Uint8Array(data);
      else if (data instanceof Uint8Array) bytes = data;
      // Browsers occasionally hand back Blob if binaryType wasn't honored;
      // we set it to 'arraybuffer' above, so this is mostly defensive.
      if (!bytes || bytes.length < 1) return;
      if (bytes[0] === CHANNEL_SYNC) {
        // Initial (or post-reconnect) replay is complete.
        this.resolveConnect();
        return;
      }
      const channel = byteToChannel(bytes[0]);
      if (!channel) return;
      const payload = bytes.slice(1);
      const handlers = this.subscribers.get(channel);
      if (!handlers) return;
      for (const h of handlers) {
        try {
          h(payload);
        } catch (err) {
          // Don't let a single handler error tear down the whole dispatch.
          // eslint-disable-next-line no-console
          console.error('[WebSocketProvider] handler threw:', err);
        }
      }
    };

    ws.onerror = () => {
      // Surface as a connect() rejection only via onclose, which always
      // follows. Browser WebSocket onerror carries no useful detail.
    };

    ws.onclose = () => {
      if (this.socket !== ws) return; // superseded — a newer socket owns us now
      this.socket = null;
      this.clearSyncTimer();
      if (this.intentionallyClosed) {
        this.setStatus('offline');
        return;
      }
      if (this.connectReject && !this.opened) {
        // The very first open never succeeded.
        const rej = this.connectReject;
        this.connectResolve = null;
        this.connectReject = null;
        this.setStatus('offline');
        rej(new Error(`WebSocketProvider: failed to connect to ${url}`));
        return;
      }
      // Lost an established connection — show offline, then scheduleReconnect
      // flips us back to 'connecting' on its next openSocket().
      this.setStatus('offline');
      this.scheduleReconnect();
    };
  }

  private scheduleReconnect(): void {
    if (this.intentionallyClosed) return;
    const delay = Math.min(
      this.maxDelay,
      this.minDelay * 2 ** this.reconnectAttempt,
    );
    this.reconnectAttempt += 1;
    this.reconnectTimer = setTimeout(() => {
      this.reconnectTimer = null;
      this.openSocket();
    }, delay);
  }
}
