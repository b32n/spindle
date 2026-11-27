import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { DeckImpl, SHAPE_CATEGORIES, PICKER_PRESETS, getBuiltinTheme, unknownShape, type ShapeElement } from '@b32nio/spindle-slides-core';
import { DeckProvider } from '../context/DeckContext';

const addElement = jest.fn(() => ({ id: 'new' }));
const setSelection = jest.fn();
let menu = { inMenu: false, openModal: jest.fn(), closeModal: jest.fn(), closeMenu: jest.fn() };

jest.mock('../hooks', () => ({
  ...jest.requireActual('../hooks'),
  useEditingId: () => null,
  useDeck: () => ({ addElement, setSelection, getActiveSlideId: () => 's1', getSlideSize: () => ({ w: 1280, h: 720 }) }),
}));
jest.mock('@b32nio/spindle-shared/react', () => ({
  ...jest.requireActual('@b32nio/spindle-shared/react'),
  useToolbarMenu: () => menu,
}));

import { ShapePicker } from './ShapePicker';
import { ShapeView } from './elements/ShapeView';

type P = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
const find = (node: React.ReactNode, pred: (e: React.ReactElement<P>) => boolean, out: React.ReactElement<P>[] = []) => {
  React.Children.forEach(node, (c) => {
    if (!React.isValidElement<P>(c)) return;
    if (pred(c)) out.push(c);
    find(c.props.children, pred, out);
  });
  return out;
};

let slots: unknown[] = [];
let n = 0;
beforeEach(() => {
  slots = [];
  addElement.mockClear();
  setSelection.mockClear();
  menu = { inMenu: false, openModal: jest.fn(), closeModal: jest.fn(), closeMenu: jest.fn() };
  jest.spyOn(React, 'useState').mockImplementation(((init: unknown) => {
    const i = n++;
    if (!(i in slots)) slots[i] = init;
    return [slots[i], (v: unknown) => { slots[i] = typeof v === 'function' ? (v as (p: unknown) => unknown)(slots[i]) : v; }];
  }) as never);
});
afterEach(() => jest.restoreAllMocks());
const render = () => {
  n = 0;
  return ShapePicker();
};

describe('ShapePicker', () => {
  it('opens a popover with every preset under its category, each button labelled', () => {
    const closed = render();
    expect(find(closed, (e) => e.props.title === 'Shapes')).toHaveLength(1);
    find(closed, (e) => e.props.title === 'Shapes')[0].props.onClick();
    const open = render();
    const popover = find(open, (e) => typeof e.props.onClose === 'function')[0];
    expect(popover).toBeDefined();
    const html = renderToStaticMarkup(popover.props.children);
    for (const c of SHAPE_CATEGORIES) expect(html).toContain(`aria-label="${c.label}"`);
    const buttons = find(popover.props.children, (e) => e.type === 'button');
    expect(buttons).toHaveLength(PICKER_PRESETS.length);
    expect(buttons.every((b) => typeof b.props['aria-label'] === 'string' && b.props['aria-label'].length > 2)).toBe(true);
    const labels = buttons.map((b) => b.props['aria-label']);
    expect(new Set(labels).size).toBe(labels.length);
    expect(labels).toContain('Right arrow');
    expect(labels).not.toContain('Arrow right');
    // a hover highlight on the button
    const style = { background: '' };
    buttons[0].props.onMouseEnter({ currentTarget: { style } });
    expect(style.background).toBe('#f1f4f8');
    buttons[0].props.onMouseLeave({ currentTarget: { style } });
    expect(style.background).toBe('#fff');
    // picking one inserts it, selects it and closes the popover
    const star = buttons.find((b) => b.props['data-shape'] === 'star6')!;
    star.props.onClick();
    expect(addElement).toHaveBeenCalledWith('s1', expect.objectContaining({ type: 'shape', shape: 'star6' }));
    expect(setSelection).toHaveBeenCalledWith({ slideId: 's1', elementIds: ['new'] });
    const after = render();
    expect(find(after, (e) => typeof e.props.onClose === 'function')).toHaveLength(0);
    // the popover closes itself on an outside click
    find(after, (e) => e.props.title === 'Shapes')[0].props.onClick();
    find(render(), (e) => typeof e.props.onClose === 'function')[0].props.onClose();
    expect(find(render(), (e) => typeof e.props.onClose === 'function')).toHaveLength(0);
  });

  it('draws bracket icons with a separate outline path', () => {
    find(render(), (e) => e.props.title === 'Shapes')[0].props.onClick();
    const popover = find(render(), (e) => typeof e.props.onClose === 'function')[0];
    const brace = find(popover.props.children, (e) => e.props['data-shape'] === 'leftBrace')[0];
    expect(renderToStaticMarkup(brace.props.children)).toContain('fill="none"');
    const ellipse = find(popover.props.children, (e) => e.props['data-shape'] === 'ellipse')[0];
    expect(renderToStaticMarkup(ellipse.props.children)).toContain('<ellipse');
  });

  it('in the overflow menu opens the same grid in a modal and closes the menu after a pick', () => {
    menu.inMenu = true;
    const row = render();
    (row as React.ReactElement<P>).props.onClick();
    expect(menu.openModal).toHaveBeenCalledWith('Insert shape', expect.anything());
    const grid = (menu.openModal.mock.calls[0] as unknown[])[1] as React.ReactElement<P>;
    find(grid, (e) => e.props['data-shape'] === 'wedgeRectCallout')[0].props.onClick();
    expect(addElement).toHaveBeenCalledWith('s1', expect.objectContaining({ shape: 'wedgeRectCallout' }));
    expect(menu.closeMenu).toHaveBeenCalled();
  });
});

describe('ShapeView', () => {
  const theme = getBuiltinTheme();
  const shape = (extra: Partial<ShapeElement>): ShapeElement =>
    ({ id: 'e', type: 'shape', x: 0, y: 0, w: 100, h: 60, rotation: 0, containerId: 's', index: 'a0', shape: 'rect', fill: { kind: 'solid', color: { kind: 'rgb', hex: '#ff0000' } }, stroke: { color: { kind: 'rgb', hex: '#000000' }, width: 2 }, ...extra }) as ShapeElement;
  const view = (el: ShapeElement): string =>
    renderToStaticMarkup(
      <DeckProvider deck={new DeckImpl('d')}>
        <ShapeView el={el} theme={theme} />
      </DeckProvider>,
    );

  it('fills an open-outline preset without stroking the closing edge, and strokes the outline without filling', () => {
    const html = view(shape({ shape: 'leftBracket' }));
    const paths = html.match(/<path [^>]*>/g)!;
    expect(paths).toHaveLength(2);
    expect(paths[0]).toContain('stroke="none"');
    expect(paths[0]).not.toContain('fill="none"');
    expect(paths[1]).toContain('fill="none"');
    expect(paths[1]).toContain('stroke="');
  });

  it('draws an ordinary preset as one path and an unknown preset as a rounded rectangle', () => {
    expect(view(shape({ shape: 'snip1Rect' })).match(/<path /g)).toHaveLength(1);
    const unknown = view(shape({ shape: unknownShape('flowChartDocument'), adjustments: [99999] }));
    expect(unknown).toBe(view(shape({ shape: 'roundRect' })));
  });
  describe('custom geometry', () => {
    const geom = {
      paths: [
        { w: 10, h: 10, commands: [{ op: 'move', x: 0, y: 0 }, { op: 'line', x: 10, y: 10 }, { op: 'close' }] },
        { w: 10, h: 10, fill: 'none', commands: [{ op: 'move', x: 0, y: 5 }, { op: 'line', x: 10, y: 5 }] },
        { w: 10, h: 10, fill: 'darken', stroke: false, commands: [{ op: 'move', x: 5, y: 0 }, { op: 'line', x: 5, y: 10 }] },
      ],
    } as ShapeElement['path'];

    it('draws each path scaled to the box, honouring fill modes and the stroke flag', () => {
      const html = view(shape({ shape: 'path', path: geom }));
      const paths = html.match(/<path [^>]*>/g)!;
      // path 1: fill + stroke; path 2: no fill (fill="none") + stroke; path 3: fill + shade, no stroke
      expect(paths).toHaveLength(6);
      expect(paths[0]).toContain('d="M0 0 L100 60 Z"');
      expect(paths[0]).toContain('stroke="none"');
      expect(paths[1]).toContain('fill="none"');
      expect(paths[1]).toContain('stroke-width="2"');
      expect(paths[2]).toContain('fill="none"');
      expect(paths[3]).toContain('stroke-width="2"');
      expect(paths[4]).toContain('stroke="none"');
      expect(paths[5]).toContain('fill="#000000"');
      expect(paths[5]).toContain('fill-opacity="0.4"');
    });

    it('falls back to a rectangle when the geometry is missing or unusable', () => {
      for (const path of [undefined, { paths: [] }, { paths: 7 }] as Array<ShapeElement['path']>) {
        const html = view(shape({ shape: 'path', path }));
        expect(html).toContain('d="M0 0 L100 0 L100 60 L0 60 Z"');
      }
    });

    it('ignores `path` on a preset shape', () => {
      const html = view(shape({ shape: 'rect', path: geom }));
      expect(html).toContain('d="M0 0 L100 0 L100 60 L0 60 Z"');
    });
  });
});
