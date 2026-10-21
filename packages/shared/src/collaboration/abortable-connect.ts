import type { CollabProvider } from './types';

/** The error an aborted attach rejects with (`name === 'AbortError'`). */
export function createAttachAbortError(): Error {
  const err = new Error('Collaboration attach was aborted (detachCollab() was called while it was in flight).');
  err.name = 'AbortError';
  return err;
}

export function isAttachAbortError(err: unknown): boolean {
  return err instanceof Error && err.name === 'AbortError';
}

/**
 * `provider.connect(roomId)` that rejects with an AbortError as soon as `signal`
 * aborts (without waiting for a hung connection) and never resolves "successfully"
 * for an already-aborted signal. The caller owns tearing down what it set up.
 */
export async function connectAbortable(
  provider: CollabProvider,
  roomId: string,
  signal?: AbortSignal
): Promise<void> {
  if (!signal) {
    await provider.connect(roomId);
    return;
  }
  if (signal.aborted) throw createAttachAbortError();
  let onAbort: (() => void) | undefined;
  const aborted = new Promise<never>((_, reject) => {
    onAbort = () => reject(createAttachAbortError());
    signal.addEventListener('abort', onAbort, { once: true });
  });
  const connecting = provider.connect(roomId);
  connecting.catch(() => undefined); // a late failure after an abort is not an unhandled rejection
  try {
    await Promise.race([connecting, aborted]);
  } finally {
    if (onAbort) signal.removeEventListener('abort', onAbort);
  }
  if (signal.aborted) throw createAttachAbortError();
}
