import * as Y from 'yjs';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { DeckImpl } from '../deck';
import type { RichTextDoc } from '../text/model';
import type { TextElement } from '../scene/types';
import { COLLAB_SCHEMA_VERSION, CollabSchemaVersionError, assertSchemaVersion, hydrateYDocFromData, serializeYDocToData } from './y-schema';
import { getCollabBlockers } from './blockers';

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });

const rich: RichTextDoc = {
  type: 'doc',
  content: [
    {
      type: 'paragraph',
      content: [
        { type: 'text', text: 'E=mc', marks: [{ type: 'caps', attrs: { kind: 'small' } }, { type: 'charSpacing', attrs: { px: 1.5 } }] },
        { type: 'text', text: '2', marks: [{ type: 'baseline', attrs: { pct: 30 } }, { type: 'lang', attrs: { lang: 'fr-FR' } }] },
        { type: 'text', text: ' page ' },
        { type: 'field', attrs: { kind: 'slidenum', format: null }, content: [{ type: 'text', text: '‹#›', marks: [{ type: 'bold' }, { type: 'baseline', attrs: { pct: 30 } }] }] },
        { type: 'field', attrs: { kind: 'datetime', format: 'time' }, content: [{ type: 'text', text: '2:30 PM' }] },
      ],
    },
  ],
};

function build(): DeckImpl {
  const d = new DeckImpl('room', 'Deck');
  d.addElement(d.getActiveSlideId(), { type: 'text', x: 0, y: 0, w: 200, h: 80, richText: rich });
  return d;
}

const inlines = (el: TextElement) => el.richText!.content[0].content!;

describe('run marks and fields in the collaboration layout', () => {
  it('has no collab blockers and survives hydrate -> serialize', () => {
    const d = build();
    expect(getCollabBlockers(d.getData())).toEqual([]);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, d.getData());
    const back = serializeYDocToData(ydoc, { elementIds: [] }, d.getActiveSlideId());
    const runs = inlines(back.slides[0].elements[0] as TextElement);
    expect(runs[0].marks).toEqual(expect.arrayContaining([expect.objectContaining({ type: 'caps', attrs: { kind: 'small' } }), expect.objectContaining({ type: 'charSpacing', attrs: { px: 1.5 } })]));
    expect(runs[1].marks).toEqual(expect.arrayContaining([expect.objectContaining({ type: 'baseline', attrs: { pct: 30 } }), expect.objectContaining({ type: 'lang', attrs: { lang: 'fr-FR' } })]));
    expect(runs[3]).toMatchObject({ type: 'field', attrs: { kind: 'slidenum' } });
    // The formatting of a field survives because it sits on the cached text.
    expect(runs[3].type === 'field' && runs[3].content![0]).toMatchObject({ text: '‹#›', marks: expect.arrayContaining([expect.objectContaining({ type: 'bold' }), expect.objectContaining({ type: 'baseline' })]) });
    expect(runs[4]).toMatchObject({ type: 'field', attrs: { kind: 'datetime', format: 'time' }, content: [{ text: '2:30 PM' }] });
  });

  it('syncs between two clients', async () => {
    __resetInMemoryRooms();
    const a = build();
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new DeckImpl('b');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    const runs = inlines(b.getData().slides[0].elements[0] as TextElement);
    expect(runs[1].marks).toEqual(expect.arrayContaining([expect.objectContaining({ type: 'baseline' })]));
    expect(runs[3]).toMatchObject({ type: 'field', attrs: { kind: 'slidenum' } });
    ha.detach();
    hb.detach();
  });

  it('stamps v3, and a client of v2 refuses the room (refuse-newer-room)', () => {
    expect(COLLAB_SCHEMA_VERSION).toBe(3);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, build().getData());
    expect(ydoc.getMap('meta').get('schemaVersion')).toBe(3);
    // What the previous release (supports up to 2) does with this room:
    expect(() => assertSchemaVersion(ydoc, 2)).toThrow(CollabSchemaVersionError);
    expect(() => assertSchemaVersion(ydoc, 3)).not.toThrow();
    expect(() => assertSchemaVersion(ydoc, 2)).toThrow(/version 3.*up to version 2/);
  });
});
