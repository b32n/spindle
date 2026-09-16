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
