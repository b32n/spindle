import * as Y from 'yjs';
import { DeckImpl } from '../deck';
import { COLLAB_SCHEMA_VERSION, CollabSchemaVersionError, assertSchemaVersion, hydrateYDocFromData, serializeYDocToData } from './y-schema';
import { getCollabBlockers } from './blockers';

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
});
