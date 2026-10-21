import { WebSocket } from 'ws';

export const COLLAB_WS = 'ws://127.0.0.1:1234';

/**
 * Resolve once the reference relay holds at least `minBytes` bytes of document frames (the creator's seed, not just its tiny SyncStep1 handshake) for
 * `room` (i.e. the room's creator has pushed its seed).
 *
 * Why: a peer that joins a brand-new room before the creator's seed reached the
 * relay finds it empty and seeds the demo content too, duplicating it (documented
 * limitation in the collab bindings: "two peers creating the same brand-new room
 * within one round-trip can still both seed"). Tests that add a second peer
 * must therefore wait for the first peer's seed. We probe like a joiner does: the
 * relay replays its doc log (tag 0) and then sends a sync marker (tag 2).
 */
export async function waitForRoomSeed(
  room: string,
  minBytes = 1000,
  timeoutMs = 10_000,
  wsBase = COLLAB_WS
): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const bytes = await docBytes(room, wsBase);
    if (bytes >= minBytes) return;
    await new Promise((r) => setTimeout(r, 50));
  }
  throw new Error(`relay never received ${minBytes} bytes of doc frames for room ${room}`);
}

function docBytes(room: string, wsBase: string): Promise<number> {
  return new Promise((resolve, reject) => {
    const ws = new WebSocket(`${wsBase}/${encodeURIComponent(room)}`);
    let bytes = 0;
    ws.on('message', (data: Buffer) => {
      const tag = data[0];
      if (tag === 0) bytes += data.length;
      if (tag === 2) {
        ws.close();
        resolve(bytes);
      }
    });
    ws.on('error', reject);
  });
}

/**
 * A room name unique to this test *run*. `testInfo.testId` is stable across runs and
 * repeats, and the relay keeps a room's log for a while after it empties, so reusing
 * it would replay (and accumulate) edits from earlier runs.
 */
export function roomName(prefix: string, testId: string): string {
  return `${prefix}-${testId}-${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
}
