import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { DeckImpl, getBuiltinTheme } from '@b32nio/spindle-slides-core';
import { ConnectorStore } from '../interactions/connector-store';
import { NodeRegistry } from '../interactions/node-registry';
import { TransientStore } from '../interactions/transient-store';
import { nearestAnchor } from '../interactions/connector-hit';

let deck: DeckImpl;
let transient = new TransientStore();
let connectors = new ConnectorStore();
let nodes = new NodeRegistry();
let selected: string[] = [];

jest.mock('../hooks', () => ({
  ...jest.requireActual('../hooks'),
  useElement: (id: string) => deck.getElement(id),
  useTheme: () => getBuiltinTheme('default'),
  useSelection: () => ({ slideId: 's1', elementIds: selected }),
  useEditingId: () => null,
}));
jest.mock('../context/DeckContext', () => ({
  useDeckContext: () => ({ deck, transient, nodes, connectors }),
}));

import { ConnectorView } from './elements/ConnectorView';
import { ConnectionPointsOverlay } from './ConnectionPointsOverlay';
import { SelectionOverlay } from './SelectionOverlay';

const K = 0.5 - Math.SQRT1_2 / 2;
let slide: string;
let tri: string;
let ell: string;
let line: string;

beforeEach(() => {
  jest.spyOn(React, 'useSyncExternalStore').mockImplementation(((_sub: unknown, get: () => unknown) => get()) as never);
  transient = new TransientStore();
  connectors = new ConnectorStore();
  nodes = new NodeRegistry();
  selected = [];
  deck = new DeckImpl('room', 'Deck');
  slide = deck.getActiveSlideId();
  ell = deck.addElement(slide, { type: 'shape', shape: 'ellipse', x: 0, y: 0, w: 100, h: 100 }).id;
  tri = deck.addElement(slide, { type: 'shape', shape: 'triangle', x: 300, y: 200, w: 100, h: 100 }).id;
  line = deck.addElement(slide, { type: 'line', connector: 'elbow', startBind: { elementId: ell, anchor: 'e' }, endBind: { elementId: tri, anchor: 'n' }, x: 0, y: 0, w: 1, h: 1 }).id;
});
afterEach(() => jest.restoreAllMocks());

const view = (interactive = true): string => renderToStaticMarkup(<ConnectorView elementId={line} interactive={interactive} />);
const pathOf = (html: string): string => /<path d="([^"]+)"/.exec(html)![1];

describe('ConnectorView', () => {
  it('routes an elbow between the bound sites, leaving each shape outward', () => {
    const html = view();
    expect(html).toContain(`data-element-id="${line}"`);
    expect(html).toContain('pointer-events:auto');
    // e of the ellipse (100, 50) to the triangle apex (350, 200): out east, in from above.
    expect(html).toContain('translate(100px, 50px)');
    expect(pathOf(html)).toContain('L');
  });

  it('is inert and unregistered when not interactive or locked, and renders nothing for a missing element', () => {
    expect(view(false)).toContain('pointer-events:none');
    deck.updateElement(line, { locked: true });
    expect(view()).toContain('pointer-events:none');
    expect(renderToStaticMarkup(<ConnectorView elementId="nope" interactive />)).toBe('');
  });

  it('follows a live frame for a bound shape, flips included', () => {
    const before = view();
    transient.set({ liveFrames: new Map([[ell, { x: 0, y: 100, w: 100, h: 100, rotation: 0 }]]) });
    const during = view();
    expect(during).not.toBe(before);
    expect(during).toContain('translate(100px, 150px)');
    // A live frame for an element the deck no longer holds still resolves.
    deck.updateElement(line, { endBind: { elementId: 'ghost', anchor: 'n' } } as never);
    transient.set({ liveFrames: new Map([['ghost', { x: 300, y: 200, w: 100, h: 100, rotation: 0 }]]) });
    expect(view()).toContain('height:150px'); // (100, 50) to the ghost's live apex (350, 200)
    transient.clear();
    expect(view()).toContain('<svg');
  });

  it('a binding to a flipped triangle attaches to the flipped apex', () => {
    deck.updateElement(tri, { flipV: true } as never);
    // apex now at the bottom (350, 300)
    expect(view()).toContain('height:250px'); // from (100, 50) down to (350, 300), not (350, 200)
  });

  it('overrides the dragged end live and drops its direction', () => {
    connectors.setEdit({ elementId: line, end: 'end', point: { x: 500, y: 400 }, snap: null, overElementId: null });
    expect(view()).toContain('translate(100px, 50px)');
    connectors.setEdit({ elementId: line, end: 'start', point: { x: 0, y: 0 }, snap: null, overElementId: null });
    expect(view()).toContain('translate(0px, 0px)');
    connectors.setEdit({ elementId: 'other', end: 'start', point: { x: 900, y: 900 }, snap: null, overElementId: null });
    expect(view()).toContain('translate(100px, 50px)');
  });

  it('registers its node when interactive', () => {
    const spy = jest.spyOn(nodes, 'register');
    jest.spyOn(React, 'useCallback').mockImplementation(((fn: unknown) => fn) as never);
    type Ref = (n: HTMLDivElement | null) => void;
    const refOf = (interactive: boolean): Ref => (ConnectorView({ elementId: line, interactive }) as React.ReactElement<{ ref: Ref }>).props.ref;
    const refFn = refOf(true);
    const node = {} as HTMLDivElement;
    refFn(node);
    expect(spy).toHaveBeenCalledWith(line, node);
    const un = jest.spyOn(nodes, 'unregister');
    refFn(null);
    expect(un).toHaveBeenCalledWith(line);
    spy.mockClear();
    refOf(false)(node);
    expect(spy).not.toHaveBeenCalled();
  });
});

describe('nearestAnchor on shape-specific sites', () => {
  it('snaps to an ellipse diagonal site on the outline, not the box corner', () => {
    const hit = nearestAnchor(deck, slide, { x: 100 * K + 1, y: 100 * K + 1 }, line, 10);
    expect(hit).toMatchObject({ elementId: ell, anchor: 'nw' });
    expect(hit!.point.x).toBeCloseTo(100 * K);
    expect(nearestAnchor(deck, slide, { x: 1, y: 1 }, line, 10)).toBeNull();
  });

  it('snaps to a triangle side midpoint and follows a flip; skips the excluded element and lines', () => {
    expect(nearestAnchor(deck, slide, { x: 325, y: 250 }, line, 10)).toMatchObject({ elementId: tri, anchor: 'nw', point: { x: 325, y: 250 } });
    deck.updateElement(tri, { flipV: true } as never);
    expect(nearestAnchor(deck, slide, { x: 350, y: 300 }, line, 10)).toMatchObject({ elementId: tri, anchor: 'n' });
    expect(nearestAnchor(deck, slide, { x: 350, y: 300 }, tri, 10)).toBeNull();
  });
});

describe('ConnectionPointsOverlay sites', () => {
  it('shows only the triangle sites, at the outline positions', () => {
    connectors.setHover(tri);
    const html = renderToStaticMarkup(<ConnectionPointsOverlay scale={1} />);
    expect(html.match(/data-connect-anchor="/g)).toHaveLength(6);
    expect(html).not.toContain('data-connect-anchor="e"');
    connectors.setHover(null);
    expect(renderToStaticMarkup(<ConnectionPointsOverlay scale={1} />)).toBe('');
  });
});

describe('SelectionOverlay connector tips', () => {
  it('puts the tips on the bound shape-specific sites and follows a live frame', () => {
    selected = [line];
    const tips = (html: string): string[] => html.match(/left:[-\d.]+px;top:[-\d.]+px/g) ?? [];
    const base = tips(renderToStaticMarkup(<SelectionOverlay scale={1} />));
    expect(base).toHaveLength(2);
    transient.set({ liveFrames: new Map([[ell, { x: 0, y: 100, w: 100, h: 100, rotation: 0 }]]) });
    expect(tips(renderToStaticMarkup(<SelectionOverlay scale={1} />))[0]).not.toBe(base[0]);
    transient.clear();
    deck.deleteElements([ell]);
    expect(renderToStaticMarkup(<SelectionOverlay scale={1} />)).toContain('data-');
  });
});
