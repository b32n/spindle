import * as Y from 'yjs';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { encodeOpaqueEntry, hashBytes, type OpaquePart } from '@b32nio/spindle-shared/opaque';
import { DeckImpl } from '../deck';
import { createOpaqueElement } from '../scene/elements';
import type { OpaqueElement } from '../scene/types';
import { COLLAB_SCHEMA_VERSION, CollabSchemaVersionError, assertSchemaVersion, getDeckYTypes, hydrateYDocFromData, readSchemaVersion, serializeYDocToData, yMapToElement } from './y-schema';

// Opaque content in collaboration (plan K11, schema v10): the element, the slide's transition and timing, the part map.

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });
const TIMING = (spid: string): string =>
  `<p:timing xmlns:p="urn:p"><p:tnLst><p:par><p:cTn id="1"><p:childTnLst><p:par><p:cTn id="2" presetClass="entr"><p:childTnLst><p:set><p:cBhvr><p:cTn id="3"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:set></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
async function part(text: string, contentType = 'application/xml'): Promise<OpaquePart> {
  const bytes = new TextEncoder().encode(text);
  return { hash: await hashBytes(bytes), contentType, bytes };
}
const settle = (): Promise<void> => new Promise((r) => setTimeout(r, 30));

async function chartDeck(): Promise<{ d: DeckImpl; chartPart: OpaquePart; slide: string }> {
  const chartPart = await part('<chart/>', 'application/vnd.chart+xml');
  const d = new DeckImpl('room', 'Deck');
  const slide = d.getActiveSlideId();
  const data = d.getData();
  const chart = createOpaqueElement({
    containerId: slide,
    index: 'a0',
    id: 'chart',
    kind: 'chart',
    partRefs: [{ hash: chartPart.hash, partName: 'ppt/charts/chart1.xml', role: 'chart' }],
    alt: 'Sales',
    title: 'Sales 2024',
    poster: 'data:image/png;base64,AA==',
    props: { progId: 'x' },
    pptxId: '3',
    x: 10,
    y: 20,
    w: 300,
    h: 200,
  });
  data.slides[0] = { ...data.slides[0], opaque: { transition: '<p:transition xmlns:p="urn:p"/>', timing: TIMING('3') }, elements: [chart] };
  data.opaqueParts = [chartPart];
  d.setData(data);
  return { d, chartPart, slide };
}

afterEach(() => jest.restoreAllMocks());

describe('schema v10', () => {
  it('stamps 10, and a client that reads only up to 9 refuses the room', () => {
    expect(COLLAB_SCHEMA_VERSION).toBe(10);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, new DeckImpl('d').getData());
    expect(readSchemaVersion(ydoc)).toBe(10);
    expect(() => assertSchemaVersion(ydoc, 9)).toThrow(CollabSchemaVersionError);
    expect(() => assertSchemaVersion(ydoc, 9)).toThrow(/version 10.*up to version 9/s);
  });

  it('a current client refuses a room a newer one wrote, and never reads its parts', async () => {
    __resetInMemoryRooms();
    const { d } = await chartDeck();
    const ha = await d.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    ha.ydoc.getMap('meta').set('schemaVersion', COLLAB_SCHEMA_VERSION + 1);
    const b = new DeckImpl('b');
    await expect(b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' })).rejects.toBeInstanceOf(CollabSchemaVersionError);
    expect(b.getOpaqueParts()).toEqual([]);
    ha.detach();
  });

  it('hydration writes the parts and the slide record, and reading the doc returns them with the element exactly', async () => {
    const { d } = await chartDeck();
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, d.getData());
    const y = getDeckYTypes(ydoc);
    expect([...y.opaqueParts.keys()]).toEqual([d.getOpaqueParts()[0].hash]);
    const back = serializeYDocToData(ydoc, { elementIds: [] }, d.getActiveSlideId(), d.getOpaqueParts());
    expect(back.slides[0].opaque).toEqual(d.getData().slides[0].opaque);
    expect(back.slides[0].elements[0]).toEqual(d.getData().slides[0].elements[0]);
    expect(back.opaqueParts).toEqual(d.getOpaqueParts());
    expect(serializeYDocToData(ydoc, { elementIds: [] }, '', []).opaqueParts).toBeUndefined();
    // a part is written once, and the element map keeps type, props and refs
    hydrateYDocFromData(ydoc, { ...d.getData(), opaqueParts: [...d.getOpaqueParts(), ...d.getOpaqueParts()] });
    expect(y.opaqueParts.size).toBe(1);
    const el = yMapToElement('chart', y.elements.get('chart') as Y.Map<unknown>) as OpaqueElement;
    expect(el).toMatchObject({ type: 'opaque', kind: 'chart', pptxId: '3', title: 'Sales 2024', props: { progId: 'x' } });
  });
});

describe('an opaque deck in a room', () => {
  it('reaches a joiner whole: element, parts, transition and timing', async () => {
    __resetInMemoryRooms();
    const { d, slide, chartPart } = await chartDeck();
    const ha = await d.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    expect(d.getOpaquePart(chartPart.hash)?.bytes).toEqual(chartPart.bytes);
    const b = new DeckImpl('b');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    expect(b.getOpaqueParts()).toEqual([chartPart]);
    expect(b.getData().slides[0].elements).toEqual(d.getData().slides[0].elements);
    expect(b.getSlide(slide)?.opaque).toEqual(d.getSlide(slide)?.opaque);
    expect(b.getElement('chart')).toMatchObject({ type: 'opaque', kind: 'chart' });
    ha.detach();
    hb.detach();
  });

  it('deleting the shape a timing names prunes it for every peer, and undo brings it back', async () => {
    __resetInMemoryRooms();
    const { d, slide } = await chartDeck();
    const ha = await d.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new DeckImpl('b');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    d.deleteElement('chart');
    expect(d.getSlide(slide)?.opaque).toEqual({ transition: '<p:transition xmlns:p="urn:p"/>' });
    expect(b.getSlide(slide)?.opaque).toEqual({ transition: '<p:transition xmlns:p="urn:p"/>' });
    expect(b.getElement('chart')).toBeUndefined();
    ha.undoManager.undo();
    expect(d.getSlide(slide)?.opaque?.timing).toBe(TIMING('3'));
    expect(b.getSlide(slide)?.opaque?.timing).toBe(TIMING('3'));
    ha.detach();
    hb.detach();
  });

  it('a peer\'s late parts reach the engine, once verified', async () => {
    __resetInMemoryRooms();
    const { d } = await chartDeck();
    const ha = await d.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new DeckImpl('b');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    const late = await part('<late/>');
    const changes: unknown[] = [];
    b.on('deckChange', (e) => changes.push(e.payload));
    ha.ydoc.getMap<Uint8Array>('opaqueParts').set(late.hash, encodeOpaqueEntry(late));
    await settle();
    expect(b.getOpaquePart(late.hash)?.bytes).toEqual(late.bytes);
    expect(changes).toContainEqual({ keys: ['opaqueParts'] });
    ha.detach();
    hb.detach();
  });

  it('a part whose bytes do not hash to its key is never taken, and does not stop the rest', async () => {
    __resetInMemoryRooms();
    const { d, chartPart } = await chartDeck();
    const ha = await d.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const forged = await part('<forged/>');
    ha.ydoc.getMap<Uint8Array>('opaqueParts').set('f'.repeat(64), encodeOpaqueEntry(forged));
    const b = new DeckImpl('b');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    await settle();
    expect(b.getOpaqueParts()).toEqual([chartPart]);
    ha.detach();
    hb.detach();
  });

  it('a deck without parts attaches and joins without hashing anything', async () => {
    __resetInMemoryRooms();
    const digest = jest.spyOn(globalThis.crypto.subtle, 'digest');
    const a = new DeckImpl('room');
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'plain' });
    const b = new DeckImpl('b');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'plain' });
    expect(digest).not.toHaveBeenCalled();
    expect(b.getOpaqueParts()).toEqual([]);
    ha.detach();
    hb.detach();
  });

  it('without SubtleCrypto a peer\'s late parts are left out instead of breaking the session, and a detached deck ignores them', async () => {
    __resetInMemoryRooms();
    const { d } = await chartDeck();
    const ha = await d.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new DeckImpl('b');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    const late = await part('<late/>');
    const real = Object.getOwnPropertyDescriptor(globalThis, 'crypto') as PropertyDescriptor;
    Object.defineProperty(globalThis, 'crypto', { value: undefined, configurable: true });
    try {
      ha.ydoc.getMap<Uint8Array>('opaqueParts').set(late.hash, encodeOpaqueEntry(late));
      await settle();
    } finally {
      Object.defineProperty(globalThis, 'crypto', real);
    }
    expect(b.getOpaquePart(late.hash)).toBeUndefined();
    // a part that arrives and is verified after the deck detached changes nothing
    const later = await part('<later/>');
    ha.ydoc.getMap<Uint8Array>('opaqueParts').set(later.hash, encodeOpaqueEntry(later));
    hb.detach();
    await settle();
    expect(b.getOpaquePart(later.hash)).toBeUndefined();
    ha.detach();
  });

  it('a joiner reads the room\'s parts again when more arrived than it read (a rejected entry among them)', async () => {
    __resetInMemoryRooms();
    const { d } = await chartDeck();
    const ha = await d.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const forged = await part('<forged/>');
    ha.ydoc.getMap<Uint8Array>('opaqueParts').set('e'.repeat(64), encodeOpaqueEntry(forged));
    const b = new DeckImpl('b');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    await settle();
    expect(b.getOpaqueParts()).toHaveLength(1);
    ha.detach();
    hb.detach();
  });
});
