import * as Y from 'yjs';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { DocumentImpl } from '../document';
import { createParagraphFromText } from '../blocks/paragraph';
import type { NamedStyle } from '../types';
import { getYDocFields, hydrateYDocFromData, serializeYDocToData } from './y-schema';

const identity = (n: string): CollabIdentity => ({ userId: n, displayName: n, color: '#111' });

const STYLES: Record<string, NamedStyle> = {
  Callout: { name: 'Callout', kind: 'paragraph', basedOn: 'Normal', text: { bold: true, color: '#c00000' }, paragraph: { leftIndent: 48, keepWithNext: true } },
  AccentWord: { name: 'Accent Word', kind: 'character', text: { smallCaps: true } },
};

function seeded(): DocumentImpl {
  const d = new DocumentImpl('room', 'Doc');
  d.setSectionBlocks(d.getSections()[0].id, [{ ...createParagraphFromText('hello'), namedStyleId: 'Callout' }]);
  d.setNamedStyles(STYLES);
  return d;
}

afterEach(() => __resetInMemoryRooms());

describe('named styles in the Y document', () => {
  it('hydrate then serialize keeps the styles and the style references', () => {
    const data = seeded().getData();
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, data);
    expect([...getYDocFields(ydoc).namedStyles.keys()].sort()).toEqual(['AccentWord', 'Callout']);
    const out = serializeYDocToData(ydoc);
    expect(out.namedStyles).toEqual(STYLES);
    expect(out.sections[0].blocks[0]).toMatchObject({ type: 'paragraph', namedStyleId: 'Callout' });
  });

  it('a document without named styles serializes without the field', () => {
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, new DocumentImpl('d', 'D').getData());
    expect(serializeYDocToData(ydoc).namedStyles).toBeUndefined();
  });

  it('ignores a prototype-polluting style id from a peer', () => {
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, seeded().getData());
    getYDocFields(ydoc).namedStyles.set('__proto__', { name: 'x', kind: 'paragraph' });
    const out = serializeYDocToData(ydoc);
    expect(Object.keys(out.namedStyles ?? {}).sort()).toEqual(['AccentWord', 'Callout']);
    expect(({} as Record<string, unknown>).name).toBeUndefined();
  });

  it('a late joiner receives the styles from the room, and later changes reach peers', async () => {
    const a = seeded();
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new DocumentImpl('room', 'Doc');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    expect(b.getNamedStyles()).toEqual(STYLES);

    a.setNamedStyles({ ...STYLES, Quote: { name: 'Quote', kind: 'paragraph' } });
    expect(Object.keys(getYDocFields(hb.ydoc).namedStyles.toJSON()).sort()).toEqual(['AccentWord', 'Callout', 'Quote']);
    ha.detach();
    hb.detach();
  });
});

describe('DocumentImpl named styles', () => {
  it('round-trips through getData / setData and can be cleared', () => {
    const d = seeded();
    const copy = new DocumentImpl('copy', 'Copy');
    copy.setData(d.getData());
    expect(copy.getNamedStyles()).toEqual(STYLES);
    copy.setNamedStyles(undefined);
    expect(copy.getNamedStyles()).toBeUndefined();
    expect(copy.getData()).not.toHaveProperty('namedStyles');
  });

  it('keeps a block reference when its style is removed (the export drops it)', () => {
    const d = seeded();
    d.setNamedStyles({});
    expect(d.getNamedStyles()).toBeUndefined();
    expect(d.getData().sections[0].blocks[0]).toMatchObject({ namedStyleId: 'Callout' });
  });
});
