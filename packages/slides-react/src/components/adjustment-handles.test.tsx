import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { DeckImpl, createShapeElement, getBuiltinTheme, type ShapeElement } from '@b32nio/spindle-slides-core';
import { NodeRegistry } from '../interactions/node-registry';
import { TransientStore } from '../interactions/transient-store';
import { createAdjustGesture, flipInBox, type GestureContext } from '../interactions/gesture';

let selected: string[] = [];
let current: unknown;
let transient = new TransientStore();
let deckStub: { getElement: (id: string) => unknown; on: () => () => void };

jest.mock('../hooks', () => ({
  ...jest.requireActual('../hooks'),
  useSelection: () => ({ slideId: 's1', elementIds: selected }),
  useElement: () => current,
  useTheme: () => getBuiltinTheme('default'),
  useEditingId: () => null,
}));
jest.mock('../context/DeckContext', () => ({
  useDeckContext: () => ({
    deck: deckStub,
    transient,
    nodes: new NodeRegistry(),
    connectors: { subscribe: () => () => undefined, get: () => ({ edit: null }) },
  }),
}));

import { SelectionOverlay } from './SelectionOverlay';
import { ElementView } from './elements/ElementView';

beforeEach(() => {
  jest.spyOn(React, 'useSyncExternalStore').mockImplementation(((_sub: unknown, get: () => unknown) => get()) as never);
  transient = new TransientStore();
});
afterEach(() => jest.restoreAllMocks());

const shape = (patch: Partial<ShapeElement> = {}): ShapeElement =>
  ({ ...createShapeElement({ containerId: 's1', index: 'a0', shape: 'roundRect', x: 100, y: 50, w: 400, h: 200 }), id: 'a', ...patch }) as ShapeElement;

const overlay = (el: unknown) => {
  deckStub = { getElement: (id) => (id === 'a' ? el : undefined), on: () => () => undefined };
  selected = ['a'];
  return renderToStaticMarkup(<SelectionOverlay scale={1} />);
};
const adjustTags = (html: string) => html.match(/<div data-adjust="\d+"[^>]*>/g) ?? [];

describe('SelectionOverlay adjustment handles', () => {
  it('draws a labelled handle for a single selected shape with adjustments', () => {
    const tags = adjustTags(overlay(shape()));
    expect(tags).toHaveLength(1);
    expect(tags[0]).toContain('aria-label="Adjust shape"');
    // roundRect at the default 0.15 of the short side (200): x = 30 → left 30 - 5
    expect(tags[0]).toContain('left:25px');
  });

  it('draws one handle per guide for a two-guide preset and none for plain shapes, text, or a multi-selection', () => {
    expect(adjustTags(overlay(shape({ shape: 'rightArrow' })))).toHaveLength(2);
    expect(adjustTags(overlay(shape({ shape: 'rect' })))).toHaveLength(0);
    expect(adjustTags(overlay({ ...shape(), type: 'text' }))).toHaveLength(0);
    deckStub = { getElement: () => shape(), on: () => () => undefined };
    selected = ['a', 'b'];
    expect(adjustTags(renderToStaticMarkup(<SelectionOverlay scale={1} />))).toHaveLength(0);
  });

  it('follows the live values during a drag and mirrors for flips', () => {
    transient.set({ liveAdjustments: new Map([['a', [0.5]]]) });
    expect(adjustTags(overlay(shape()))[0]).toContain('left:95px'); // 0.5 * 200 - 5
    transient.clear();
    expect(adjustTags(overlay(shape({ flipH: true })))[0]).toContain('left:365px'); // 400 - 30 - 5
  });
});

describe('flipInBox', () => {
  it('mirrors only the flipped axes', () => {
    const box = { w: 100, h: 50 };
    expect(flipInBox({ x: 10, y: 5 }, box)).toEqual({ x: 10, y: 5 });
    expect(flipInBox({ x: 10, y: 5 }, { ...box, flipH: true })).toEqual({ x: 90, y: 5 });
    expect(flipInBox({ x: 10, y: 5 }, { ...box, flipV: true })).toEqual({ x: 10, y: 45 });
  });
});

describe('ElementView shapes', () => {
  const path = (html: string) => html.match(/<path d="([^"]*)"/)?.[1];

  it('redraws an interactive shape with the live adjustments, and a static one never subscribes', () => {
    current = shape({ adjustments: [0] });
    const rest = path(renderToStaticMarkup(<ElementView elementId="a" interactive />));
    transient.set({ liveAdjustments: new Map([['a', [0.5]]]) });
    const live = path(renderToStaticMarkup(<ElementView elementId="a" interactive />));
    expect(live).toBeDefined();
    expect(live).not.toEqual(rest);
    const stat = path(renderToStaticMarkup(<ElementView elementId="a" />));
    expect(stat).toEqual(rest);
  });
});

describe('createAdjustGesture', () => {
  const make = (patch: Partial<ShapeElement> = {}) => {
    const deck = new DeckImpl();
    const slideId = deck.getActiveSlideId()!;
    const el = deck.addElement(slideId, { type: 'shape', shape: 'roundRect', x: 100, y: 50, w: 400, h: 200, ...patch } as never) as ShapeElement;
    const store = new TransientStore();
    const ctx: GestureContext = { deck, nodes: new NodeRegistry(), transient: store, slideId, slideSize: { w: 1280, h: 720 }, snapThreshold: 5 };
    return { deck, el, store, ctx };
  };

  it('previews through the transient store and commits one undo entry on release', () => {
    const { deck, el, store, ctx } = make();
    const g = createAdjustGesture(ctx, el.id, 0)!;
    g.onMove({ x: 100 + 60, y: 50 });
    expect(store.get().liveAdjustments?.get(el.id)).toEqual([0.3]);
    expect((deck.getElement(el.id) as ShapeElement).adjustments).toBeUndefined();
    g.onMove({ x: 100 + 9999, y: 50 });
    g.onEnd();
    expect((deck.getElement(el.id) as ShapeElement).adjustments).toEqual([0.5]);
    expect(store.get().liveAdjustments).toBeNull();
    deck.undo();
    expect((deck.getElement(el.id) as ShapeElement).adjustments).toBeUndefined();
  });

  it('maps through rotation and flips', () => {
    const { el, store, ctx } = make({ rotation: 90, flipH: true });
    const g = createAdjustGesture(ctx, el.id, 0)!;
    // local (340, 0) is (60, 0) once mirrored; about the centre (300, 150) that is (140, -100), turned 90deg clockwise: (100, 140)
    g.onMove({ x: 400, y: 290 });
    expect(store.get().liveAdjustments?.get(el.id)?.[0]).toBeCloseTo(0.3);
  });

  it('commits nothing when the drag changed nothing, and ignores unmapped moves', () => {
    const { deck, el, ctx } = make();
    const g = createAdjustGesture(ctx, el.id, 0)!;
    g.onMove({ x: 100 + 30, y: 50 }); // the default value
    g.onEnd();
    expect((deck.getElement(el.id) as ShapeElement).adjustments).toBeUndefined();
    const none = createAdjustGesture(ctx, el.id, 7)!;
    none.onMove({ x: 1, y: 1 });
    none.onEnd();
    expect((deck.getElement(el.id) as ShapeElement).adjustments).toBeUndefined();
  });

  it('is not available for a missing element or a non-shape', () => {
    const { deck, ctx } = make();
    expect(createAdjustGesture(ctx, 'nope', 0)).toBeNull();
    const text = deck.addElement(ctx.slideId, { type: 'text', x: 0, y: 0, w: 10, h: 10 } as never);
    expect(createAdjustGesture(ctx, text.id, 0)).toBeNull();
  });
});
