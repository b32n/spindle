import * as Y from 'yjs';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { DeckImpl } from '../deck';
import type { RichTextDoc } from '../text/model';
import type { TextElement } from '../scene/types';
import { COLLAB_SCHEMA_VERSION, CollabSchemaVersionError, hydrateYDocFromData, serializeYDocToData } from './y-schema';
import { getCollabBlockers } from './blockers';

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });

const rich: RichTextDoc = {
  type: 'doc',
  content: [
    {
      type: 'paragraph',
      attrs: {
        bullet: { kind: 'auto', startAt: 3, scheme: 'alphaUpper', color: { kind: 'theme', slot: 'accent2' }, sizePct: 90 },
        lineHeightPx: 22,
        spaceBeforePct: 25,
        spaceAfterPct: 10,
        marginLeft: 36,
        firstLineIndent: -18,
        tabStops: [{ pos: 72, align: 'r' }],
      },
      content: [{ type: 'text', text: 'one' }],
    },
    { type: 'paragraph', attrs: { bullet: { kind: 'char', char: '★' } }, content: [{ type: 'text', text: 'two' }] },
  ],
};

function build(): DeckImpl {
  const d = new DeckImpl('room', 'Deck');
  d.addElement(d.getActiveSlideId(), { type: 'text', x: 0, y: 0, w: 200, h: 80, richText: rich });
  return d;
}

describe('paragraph attributes in the collaboration layout', () => {
  it('has no collab blockers and survives hydrate -> serialize', () => {
    const d = build();
    expect(getCollabBlockers(d.getData())).toEqual([]);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, d.getData());
    const back = serializeYDocToData(ydoc, { elementIds: [] }, d.getActiveSlideId());
    const el = back.slides[0].elements[0] as TextElement;
    expect(el.richText?.content[0].attrs).toMatchObject(rich.content[0].attrs as object);
    expect(el.richText?.content[1].attrs).toMatchObject({ bullet: { kind: 'char', char: '★' } });
  });

  it('syncs between two clients', async () => {
    __resetInMemoryRooms();
    const a = build();
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new DeckImpl('b');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    const el = b.getData().slides[0].elements[0] as TextElement;
    expect(el.richText?.content[0].attrs).toMatchObject({ bullet: { startAt: 3 }, marginLeft: 36, tabStops: [{ pos: 72, align: 'r' }] });
    ha.detach();
    hb.detach();
  });

  it('stamps the bumped version and a client of the previous version would refuse the room', () => {
    expect(COLLAB_SCHEMA_VERSION).toBeGreaterThanOrEqual(2);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, build().getData());
    expect(ydoc.getMap('meta').get('schemaVersion')).toBe(COLLAB_SCHEMA_VERSION);
    const err = new CollabSchemaVersionError(2, 1);
    expect(err.message).toContain('version 2');
    expect(err.message).toContain('up to version 1');
  });
});
