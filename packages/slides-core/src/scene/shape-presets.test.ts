import * as Y from 'yjs';
import {
  SHAPE_CATEGORIES,
  SHAPE_PRESETS,
  SHAPE_PRESET_INFO,
  isShapePreset,
  isUnknownShape,
  resolveAdjustments,
  shapePrst,
  shapeFromPrst,
  unknownShape,
} from './shape-presets';
import { createShapeElement } from './elements';
import { DeckImpl } from '../deck';
import {
  COLLAB_SCHEMA_VERSION,
  CollabSchemaVersionError,
  assertSchemaVersion,
  hydrateYDocFromData,
  serializeYDocToData,
} from '../collab/y-schema';
import { getCollabBlockers } from '../collab/blockers';
import type { ShapeElement } from './types';

describe('preset table', () => {
  it('holds about 70 presets, each in a known category, listed once in picker order', () => {
    expect(SHAPE_PRESETS.length).toBeGreaterThanOrEqual(70);
    expect(new Set(SHAPE_PRESETS).size).toBe(SHAPE_PRESETS.length);
    expect(SHAPE_PRESETS.length).toBe(Object.keys(SHAPE_PRESET_INFO).length);
    const cats = new Set(SHAPE_CATEGORIES.map((c) => c.id));
    for (const p of SHAPE_PRESETS) expect(cats.has(SHAPE_PRESET_INFO[p].category)).toBe(true);
    const order = SHAPE_PRESETS.map((p) => SHAPE_CATEGORIES.findIndex((c) => c.id === SHAPE_PRESET_INFO[p].category));
    expect(order).toEqual([...order].sort((x, y) => x - y));
  });

  it('keeps the original 18 presets', () => {
    for (const p of ['rect', 'roundRect', 'ellipse', 'triangle', 'rightTriangle', 'diamond', 'pentagon', 'hexagon', 'octagon', 'star5', 'arrowRight', 'arrowLeft', 'chevron', 'parallelogram', 'trapezoid', 'plus', 'heart', 'cloud']) {
      expect(isShapePreset(p)).toBe(true);
    }
  });

  it('has sane guide specs: default inside the range, ordered range', () => {
    for (const p of SHAPE_PRESETS) {
      for (const g of SHAPE_PRESET_INFO[p].adj ?? []) {
        expect(g.min).toBeLessThanOrEqual(g.def);
        expect(g.def).toBeLessThanOrEqual(g.max);
      }
    }
  });
});

describe('identity helpers', () => {
  it('tells presets from unknown ones, including inherited object keys', () => {
    expect(isShapePreset('rect')).toBe(true);
    expect(isShapePreset('rightArrow2')).toBe(false);
    expect(isShapePreset('toString')).toBe(false);
    expect(isUnknownShape('unknown:flowChartDocument')).toBe(true);
    expect(isUnknownShape('unknown:')).toBe(false);
    expect(isUnknownShape('rect')).toBe(false);
  });

  it('round-trips the original prst name', () => {
    expect(unknownShape('flowChartDocument')).toBe('unknown:flowChartDocument');
    expect(shapePrst('unknown:flowChartDocument')).toBe('flowChartDocument');
    expect(shapePrst('star6')).toBe('star6');
  });

  it('maps OOXML names to presets and back, keeping the three legacy names exportable', () => {
    expect(shapeFromPrst('rtTriangle')).toBe('rightTriangle');
    expect(shapeFromPrst('rightArrow')).toBe('rightArrow');
    expect(shapeFromPrst('star5')).toBe('star5');
    expect(shapeFromPrst('ribbon')).toBe('unknown:ribbon');
    expect(shapeFromPrst('constructor')).toBe('unknown:constructor');
    expect(shapePrst('rightTriangle')).toBe('rtTriangle');
    expect(shapePrst('arrowRight')).toBe('rightArrow');
    expect(shapePrst('arrowLeft')).toBe('leftArrow');
    for (const p of SHAPE_PRESETS) expect(shapeFromPrst(shapePrst(p))).toBe(p === 'arrowRight' ? 'rightArrow' : p === 'arrowLeft' ? 'leftArrow' : p);
  });
});

describe('resolveAdjustments', () => {
  it('fills defaults, clamps to the range and ignores junk', () => {
    expect(resolveAdjustments('rect')).toEqual([]);
    expect(resolveAdjustments('roundRect')).toEqual([0.15]);
    expect(resolveAdjustments('roundRect', [0.9])).toEqual([0.5]);
    expect(resolveAdjustments('roundRect', [-1])).toEqual([0]);
    expect(resolveAdjustments('roundRect', [NaN])).toEqual([0.15]);
    expect(resolveAdjustments('rightArrow', [0.2])).toEqual([0.2, 0.5]);
    expect(resolveAdjustments('rightArrow', [0.2, 7, 99])).toEqual([0.2, 1]);
    expect(resolveAdjustments('wedgeRectCallout', [-5, 5])).toEqual([-1, 2]);
  });

  it('never interprets an unknown shape or an unrecognised name', () => {
    expect(resolveAdjustments('unknown:foo', [1, 2, 3])).toEqual([]);
    expect(resolveAdjustments('bogus' as never, [1])).toEqual([]);
  });
});

describe('unknown preset passthrough', () => {
  const unknown = (): ShapeElement =>
    createShapeElement({ id: 'u', containerId: 's', index: 'a0', x: 0, y: 0, w: 10, h: 10, shape: unknownShape('flowChartDocument'), adjustments: [12345, -7] });

  it('keeps name and raw adjustments on the element', () => {
    const el = unknown();
    expect(el.shape).toBe('unknown:flowChartDocument');
    expect(el.adjustments).toEqual([12345, -7]);
  });

  it('round-trips through the Y layout unchanged, with new presets, and raises no blocker', () => {
    const d = new DeckImpl('d');
    const slideId = d.getActiveSlideId();
    const known = d.addElement(slideId, { type: 'shape', shape: 'wedgeRectCallout', adjustments: [-0.3, 0.7], x: 1, y: 2, w: 30, h: 20 } as never) as ShapeElement;
    const odd = d.addElement(slideId, { type: 'shape', shape: unknownShape('flowChartDocument'), adjustments: [12345, -7], x: 5, y: 6, w: 30, h: 20 } as never) as ShapeElement;
    expect(getCollabBlockers(d.getData())).toEqual([]);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, d.getData());
    const back = serializeYDocToData(ydoc, { elementIds: [] }, slideId);
    const els = back.slides.find((s) => s.id === slideId)!.elements as ShapeElement[];
    expect(els.find((e) => e.id === known.id)).toMatchObject({ shape: 'wedgeRectCallout', adjustments: [-0.3, 0.7] });
    expect(els.find((e) => e.id === odd.id)).toMatchObject({ shape: 'unknown:flowChartDocument', adjustments: [12345, -7] });
  });
});

describe('collaboration schema v5', () => {
  it('a v4 client refuses a v5 room (refuse-newer-room)', () => {
    expect(COLLAB_SCHEMA_VERSION).toBeGreaterThanOrEqual(5);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, new DeckImpl('d').getData());
    expect(ydoc.getMap('meta').get('schemaVersion')).toBe(COLLAB_SCHEMA_VERSION);
    expect(() => assertSchemaVersion(ydoc, 4)).toThrow(CollabSchemaVersionError);
    expect(() => assertSchemaVersion(ydoc, 4)).toThrow(/up to version 4/);
    expect(() => assertSchemaVersion(ydoc, COLLAB_SCHEMA_VERSION)).not.toThrow();
  });
});
