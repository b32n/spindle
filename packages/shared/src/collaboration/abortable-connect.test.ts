import { connectAbortable } from './abortable-connect';
import type { CollabProvider } from './types';

function provider(connect: () => Promise<void>): CollabProvider {
  return { connect, disconnect: () => undefined, send: () => undefined, onMessage: () => () => undefined };
}

describe('connectAbortable', () => {
  it('connects normally without a signal', async () => {
    await expect(connectAbortable(provider(async () => undefined), 'r')).resolves.toBeUndefined();
  });

  it('rejects immediately for an already-aborted signal, without connecting', async () => {
    const connect = jest.fn(async () => undefined);
    const ac = new AbortController();
    ac.abort();
    await expect(connectAbortable(provider(connect), 'r', ac.signal)).rejects.toMatchObject({ name: 'AbortError' });
    expect(connect).not.toHaveBeenCalled();
  });

  it('rejects as soon as the signal aborts, even if connect() never settles', async () => {
    const ac = new AbortController();
    const p = connectAbortable(provider(() => new Promise<void>(() => undefined)), 'r', ac.signal);
    ac.abort();
    await expect(p).rejects.toMatchObject({ name: 'AbortError' });
  });

  it('does not leak an unhandled rejection when connect() fails after an abort', async () => {
    const ac = new AbortController();
    let fail!: (e: Error) => void;
    const p = connectAbortable(provider(() => new Promise<void>((_, rej) => (fail = rej))), 'r', ac.signal);
    ac.abort();
    await expect(p).rejects.toMatchObject({ name: 'AbortError' });
    fail(new Error('late'));
    await new Promise((r) => setTimeout(r, 0));
  });
});
