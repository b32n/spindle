import * as Y from 'yjs';
import { DeckImpl } from '../deck';
import { COLLAB_SCHEMA_VERSION, CollabSchemaVersionError, assertSchemaVersion, hydrateYDocFromData, serializeYDocToData } from './y-schema';
import { getCollabBlockers } from './blockers';
import type { CustomGeometry, ShapeElement } from '../scene/types';

const GEOM: CustomGeometry = {
  paths: [
    {
      w: 100,
      h: 80,
      commands: [
        { op: 'move', x: 0, y: 80 },
        { op: 'cubic', x1: 10, y1: 0, x2: 90, y2: 0, x: 100, y: 80 },
        { op: 'arc', wR: 5, hR: 5, stAng: 0, swAng: 90 },
        { op: 'close' },
      ],
    },
    { w: 10, h: 10, fill: 'none', stroke: false, commands: [{ op: 'move', x: 1, y: 1 }, { op: 'line', x: 9, y: 9 }] },
  ],
};

describe('custom geometry in collab (schema v8)', () => {
  it('the schema is at least 8 and a v7 client refuses a v8 room', () => {
    expect(COLLAB_SCHEMA_VERSION).toBeGreaterThanOrEqual(8);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, new DeckImpl('d').getData());
    expect(ydoc.getMap('meta').get('schemaVersion')).toBe(COLLAB_SCHEMA_VERSION);
    expect(() => assertSchemaVersion(ydoc, 7)).toThrow(CollabSchemaVersionError);
    expect(() => assertSchemaVersion(ydoc, 7)).toThrow(/up to version 7/);
    expect(() => assertSchemaVersion(ydoc, COLLAB_SCHEMA_VERSION)).not.toThrow();
  });

  it('a path shape survives hydrate and serialize', () => {
    const d = new DeckImpl('room', 'Deck');
    const s = d.getActiveSlideId();
    const id = d.addElement(s, { type: 'shape', shape: 'path', path: GEOM, x: 0, y: 0, w: 100, h: 80 }).id;
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, d.getData());
    const back = serializeYDocToData(ydoc, { elementIds: [] }, s);
    const el = back.slides[0].elements.find((e) => e.id === id) as ShapeElement;
    expect(el.shape).toBe('path');
    expect(el.path).toEqual(GEOM);
  });

  it('a path shape is not a collab blocker', () => {
    const d = new DeckImpl('room', 'Deck');
    d.addElement(d.getActiveSlideId(), { type: 'shape', shape: 'path', path: GEOM, x: 0, y: 0, w: 100, h: 80 });
    expect(getCollabBlockers(d.getData())).toEqual([]);
  });
});
