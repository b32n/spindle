import * as Y from 'yjs';
import { DeckImpl } from '../deck';
import { COLLAB_SCHEMA_VERSION, CollabSchemaVersionError, assertSchemaVersion, hydrateYDocFromData, serializeYDocToData } from './y-schema';
import { getCollabBlockers } from './blockers';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import type { LineElement } from '../scene/types';

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });

describe('elbow and curved connectors in collab (schema v7)', () => {
  it('the schema is at least 7 and a v6 client refuses a v7 room', () => {
    expect(COLLAB_SCHEMA_VERSION).toBeGreaterThanOrEqual(7);
    const ydoc = new Y.Doc();
    ydoc.getMap('meta').set('schemaVersion', 7);
    expect(() => assertSchemaVersion(ydoc, 6)).toThrow(CollabSchemaVersionError);
    expect(() => assertSchemaVersion(ydoc, 7)).not.toThrow();
  });

  it('connector kind, adjustments and bound sites survive hydrate and serialize', () => {
    const d = new DeckImpl('room', 'Deck');
    const s = d.getActiveSlideId();
    const a = d.addElement(s, { type: 'shape', shape: 'ellipse', x: 0, y: 0, w: 100, h: 100 }).id;
    const b = d.addElement(s, { type: 'shape', shape: 'triangle', x: 300, y: 200, w: 100, h: 100 }).id;
    const l = d.addElement(s, {
      type: 'line', connector: 'elbow', adjustments: [0.3],
      startBind: { elementId: a, anchor: 'se' }, endBind: { elementId: b, anchor: 'nw' },
      x: 0, y: 0, w: 10, h: 10,
    }).id;
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, d.getData());
    const back = serializeYDocToData(ydoc, { elementIds: [] }, s);
    const el = back.slides[0].elements.find((e) => e.id === l);
    expect(el).toMatchObject({ connector: 'elbow', adjustments: [0.3], startBind: { elementId: a, anchor: 'se' }, endBind: { elementId: b, anchor: 'nw' } });
  });

  it('changing a connector kind is not a collab blocker', () => {
    const d = new DeckImpl('room', 'Deck');
    const s = d.getActiveSlideId();
    d.addElement(s, { type: 'line', connector: 'curved', x: 0, y: 0, w: 10, h: 10 });
    expect(getCollabBlockers(d.getData())).toEqual([]);
  });

  it('propagates routing changes between two bound decks, including the reset to straight', async () => {
    __resetInMemoryRooms();
    const a = new DeckImpl('room', 'Deck');
    const s = a.getActiveSlideId();
    const p = a.addElement(s, { type: 'shape', shape: 'ellipse', x: 0, y: 0, w: 100, h: 100 }).id;
    const q = a.addElement(s, { type: 'shape', shape: 'triangle', x: 300, y: 200, w: 100, h: 100 }).id;
    const l = a.addElement(s, { type: 'line', startBind: { elementId: p, anchor: 'se' }, endBind: { elementId: q, anchor: 'nw' }, x: 0, y: 0, w: 10, h: 10 }).id;
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new DeckImpl('joiner');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    const peer = (): LineElement => b.getElement(l) as LineElement;
    expect(peer().connector).toBeUndefined();

    a.setConnectorKind(l, 'elbow');
    a.updateElement(l, { adjustments: [0.3] });
    expect(peer()).toMatchObject({ connector: 'elbow', adjustments: [0.3], startBind: { elementId: p, anchor: 'se' }, endBind: { elementId: q, anchor: 'nw' } });

    a.setConnectorKind(l, 'curved');
    expect(peer().connector).toBe('curved');
    expect(peer()).not.toHaveProperty('adjustments');

    b.setConnectorKind(l, 'straight');
    expect(a.getElement(l)).not.toHaveProperty('connector');
    expect(a.getElement(l)).toStrictEqual(peer());

    a.setConnectorKind(l, 'elbow');
    // Deliberately untyped: exercises the `{connector: undefined}` reset path.
    a.updateElement(l, { connector: undefined, adjustments: undefined } as never);
    expect(peer().connector).toBeUndefined();
    ha.detach();
    hb.detach();
  });
});
