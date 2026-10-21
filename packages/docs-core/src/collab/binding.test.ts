import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { DocumentImpl } from '../document';
import { createParagraphFromText } from '../blocks/paragraph';
import { getYDocFields } from './y-schema';

const identity = (n: string): CollabIdentity => ({ userId: n, displayName: n, color: '#111' });

function seeded(): DocumentImpl {
  const d = new DocumentImpl('room', 'Doc');
  d.setSectionBlocks(d.getSections()[0].id, [
    createParagraphFromText('hello'),
    createParagraphFromText('world'),
  ]);
  return d;
}

afterEach(() => __resetInMemoryRooms());

describe('docs collab seeding marker', () => {
  it('a joiner does not resurrect content deleted after the room was seeded', async () => {
    __resetInMemoryRooms();
    const a = seeded();
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    // The seed sets a persistent marker.
    expect(ha.ydoc.getMap('__spindle').get('seeded')).toBe(true);

    // Peers delete the entire body.
    const content = getYDocFields(ha.ydoc).content;
    ha.ydoc.transact(() => content.delete(0, content.length));
    expect(content.length).toBe(0);
    // Marker survives the deletion.
    expect(ha.ydoc.getMap('__spindle').get('seeded')).toBe(true);

    // B joins with the SAME initial data. The room is empty but was seeded, so
    // B must not re-hydrate its copy (that would resurrect the deleted content).
    const b = seeded();
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    expect(getYDocFields(hb.ydoc).content.length).toBe(0);

    ha.detach();
    hb.detach();
  });

  it('a fresh creator still seeds an empty room', async () => {
    __resetInMemoryRooms();
    const a = seeded();
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    expect(getYDocFields(ha.ydoc).content.length).toBeGreaterThan(0);
    ha.detach();
  });
});

// A provider whose connect() stays pending until released — an attach in flight.
class GatedProvider extends InMemoryProvider {
  disconnected = false;
  private release!: () => void;
  private gate = new Promise<void>((r) => (this.release = r));
  override async connect(roomId: string): Promise<void> {
    await this.gate;
    return super.connect(roomId);
  }
  override disconnect(): void {
    this.disconnected = true;
    super.disconnect();
  }
  open(): void {
    this.release();
  }
}

describe('detachCollab() cancels an in-flight attachCollab()', () => {
  it('the aborted attach never seeds the room, leaves no socket, and rejects with AbortError', async () => {
    __resetInMemoryRooms();
    const first = seeded();
    const gated = new GatedProvider();
    const pending = first.attachCollab(gated, identity('A'), { roomId: 'room' });
    const settled = pending.then(() => 'resolved', (e: Error) => e.name);
    first.detachCollab();
    // The gate opens only AFTER the detach — the old code seeded the room right here.
    gated.open();
    expect(await settled).toBe('AbortError');
    expect(gated.disconnected).toBe(true);
    expect(first.getCollabHandle()).toBeNull();

    // The room was never seeded, so the next attach is the one and only creator.
    const second = seeded();
    const h2 = await second.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const joiner = new DocumentImpl('joiner', 'Doc');
    const h3 = await joiner.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    expect(getYDocFields(h3.ydoc).content.length).toBe(2);
    h2.detach();
    h3.detach();
  });

  it('a detach/attach pair (React StrictMode) seeds exactly once', async () => {
    __resetInMemoryRooms();
    const d = seeded();
    const gated = new GatedProvider();
    const stale = d.attachCollab(gated, identity('A'), { roomId: 'room' }).catch((e: Error) => e.name);
    d.detachCollab();
    const handle = await d.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    gated.open();
    expect(await stale).toBe('AbortError');
    expect(getYDocFields(handle.ydoc).content.length).toBe(2);
    handle.detach();
  });
});
