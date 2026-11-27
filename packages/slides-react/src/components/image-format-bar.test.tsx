import React from 'react';
import type { ImageElement } from '@b32nio/spindle-slides-core';
import { Popover } from './Popover';
import { ToolbarDivider } from './toolbarUI';
import { ImageCropFields } from './ImageCropFields';

const updateElement = jest.fn();
let selected: string[] = ['i'];
let element: Partial<ImageElement> | undefined;
let imageHooks: Record<string, unknown> = {};
let menu = { inMenu: false, openModal: jest.fn(), closeModal: jest.fn(), closeMenu: jest.fn() };

jest.mock('../hooks', () => ({
  useDeck: () => ({ updateElement }),
  useSelection: () => ({ slideId: 's', elementIds: selected }),
  useElement: () => element,
}));
jest.mock('@b32nio/spindle-shared/react', () => ({
  ...jest.requireActual('@b32nio/spindle-shared/react'),
  useToolbarMenu: () => menu,
  useImageHooks: () => imageHooks,
}));

import { MenuRow } from '@b32nio/spindle-shared/react';
import { ImageFormatBar, LiveImageCropFields } from './ImageFormatBar';

type P = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
const find = (node: React.ReactNode, pred: (e: React.ReactElement<P>) => boolean, out: React.ReactElement<P>[] = []) => {
  React.Children.forEach(node, (c) => {
    if (!React.isValidElement<P>(c)) return;
    if (pred(c)) out.push(c);
    find(c.props.children, pred, out);
  });
  return out;
};
const byTitle = (t: string) => (e: React.ReactElement<P>) => e.props.title === t;

// A minimal hook harness: state persists across renders, so a click can be followed by a re-render.
let slots: unknown[] = [];
let n = 0;
const render = () => {
  n = 0;
  return ImageFormatBar();
};

beforeEach(() => {
  slots = [];
  updateElement.mockClear();
  imageHooks = {};
  selected = ['i'];
  element = { id: 'i', type: 'image', src: 'x', crop: { l: 10, t: 0, r: 0, b: 0 } };
  menu = { inMenu: false, openModal: jest.fn(), closeModal: jest.fn(), closeMenu: jest.fn() };
  jest.spyOn(React, 'useRef').mockImplementation(() => ({ current: null }));
  jest.spyOn(React, 'useState').mockImplementation(((init: unknown) => {
    const i = n++;
    if (!(i in slots)) slots[i] = typeof init === 'function' ? (init as () => unknown)() : init;
    return [slots[i], (v: unknown) => { slots[i] = typeof v === 'function' ? (v as (p: unknown) => unknown)(slots[i]) : v; }];
  }) as never);
});
afterEach(() => jest.restoreAllMocks());

describe('ImageFormatBar crop control', () => {
  it('renders nothing without a single selected image', () => {
    selected = [];
    expect(render()).toBeNull();
    selected = ['i'];
    element = { id: 'i', type: 'shape' } as never;
    expect(render()).toBeNull();
    element = undefined;
    expect(render()).toBeNull();
  });

  it('inline: the button toggles an anchored popover with live fields', () => {
    const [btn] = find(render(), byTitle('Crop image'));
    expect(btn.props.active).toBe(true);
    expect(find(render(), (e) => e.type === Popover)).toHaveLength(0);
    btn.props.onClick();
    const [pop] = find(render(), (e) => e.type === Popover);
    const [fields] = find(pop, (e) => e.type === ImageCropFields);
    fields.props.onChange({ l: 5, t: 0, r: 0, b: 0 });
    expect(updateElement).toHaveBeenCalledWith('i', { crop: { l: 5, t: 0, r: 0, b: 0 } });
    pop.props.onClose();
    expect(find(render(), (e) => e.type === Popover)).toHaveLength(0);
    // toggling: a second click closes
    find(render(), byTitle('Crop image'))[0].props.onClick();
    find(render(), byTitle('Crop image'))[0].props.onClick();
    expect(find(render(), (e) => e.type === Popover)).toHaveLength(0);
    expect(menu.openModal).not.toHaveBeenCalled();
  });

  it('an uncropped image leaves the button inactive', () => {
    element = { id: 'i', type: 'image', src: 'x' };
    expect(find(render(), byTitle('Crop image'))[0].props.active).toBe(false);
  });

  it('the popover closes when the selection moves to another image', () => {
    find(render(), byTitle('Crop image'))[0].props.onClick();
    expect(find(render(), (e) => e.type === Popover)).toHaveLength(1);
    selected = ['j'];
    element = { id: 'j', type: 'image', src: 'y' };
    expect(find(render(), (e) => e.type === Popover)).toHaveLength(0);
  });

  it('in the overflow menu: a row opens a modal (no popover, no toolbar button)', () => {
    menu.inMenu = true;
    const tree = render();
    expect(find(tree, byTitle('Crop image'))).toHaveLength(0);
    expect(find(tree, (e) => e.type === Popover)).toHaveLength(0);
    // The Crop row is the first child: a leading divider would make the overflow surface collapse it as empty.
    const kids = React.Children.toArray((tree as React.ReactElement<P>).props.children) as React.ReactElement<P>[];
    expect(kids[0].type).toBe(MenuRow);
    expect(kids[0].props.label).toBe('Crop image');
    const [row] = find(tree, (e) => e.props.label === 'Crop image');
    row.props.onClick();
    expect(menu.openModal).toHaveBeenCalledTimes(1);
    const [title, content] = menu.openModal.mock.calls[0];
    expect(title).toBe('Crop image');
    expect((content as React.ReactElement<P>).type).toBe(LiveImageCropFields);
    expect((content as React.ReactElement<P>).props.id).toBe('i');
  });

  describe('overflow menu keeps every other image control', () => {
    const row = (label: string) => find(render(), (e) => e.type === MenuRow && e.props.label === label)[0];
    beforeEach(() => { menu.inMenu = true; });

    it('lists replace, fit and both flips, plus the hidden file input', () => {
      const labels = find(render(), (e) => e.type === MenuRow).map((e) => e.props.label);
      expect(labels).toEqual(['Crop image', 'Replace image', 'Image fit', 'Flip horizontal', 'Flip vertical']);
      expect(find(render(), (e) => e.type === 'input' && e.props.type === 'file')).toHaveLength(1);
      expect(find(render(), (e) => e.type === ToolbarDivider)).toHaveLength(0);
    });

    it('Replace image clicks the file input', () => {
      const click = jest.fn();
      (React.useRef as jest.Mock).mockImplementation(() => ({ current: { click } }));
      row('Replace image').props.onClick();
      expect(click).toHaveBeenCalledTimes(1);
    });

    it('Replace image tolerates a missing file input ref', () => {
      expect(() => row('Replace image').props.onClick()).not.toThrow();
    });

    it('Replace with URL appears only when the host can resolve images', () => {
      expect(row('Replace with URL')).toBeUndefined();
      imageHooks = { resolveImage: jest.fn() };
      expect(row('Replace with URL')).toBeDefined();
      const prompt = jest.fn().mockReturnValue(null);
      (globalThis as P).window = { prompt };
      row('Replace with URL').props.onClick();
      expect(prompt).toHaveBeenCalledWith('Image URL', 'x');
      expect(updateElement).not.toHaveBeenCalled();
      delete (globalThis as P).window;
    });

    it('Image fit cycles stretch, fit, crop to fill and shows the current value', () => {
      expect(row('Image fit').props.hint).toBe('Stretch');
      row('Image fit').props.onClick();
      expect(updateElement).toHaveBeenLastCalledWith('i', { fit: 'contain' });
      element = { ...element, fit: 'contain' };
      expect(row('Image fit').props.hint).toBe('Fit');
      row('Image fit').props.onClick();
      expect(updateElement).toHaveBeenLastCalledWith('i', { fit: 'cover' });
      element = { ...element, fit: 'cover' };
      expect(row('Image fit').props.hint).toBe('Crop to fill');
      row('Image fit').props.onClick();
      expect(updateElement).toHaveBeenLastCalledWith('i', { fit: 'fill' });
    });

    it('Image fit recovers from an unknown stored value', () => {
      element = { ...element, fit: 'bogus' as never };
      expect(row('Image fit').props.hint).toBeUndefined();
      row('Image fit').props.onClick();
      expect(updateElement).toHaveBeenLastCalledWith('i', { fit: 'fill' });
    });

    it('flip rows toggle and show On only when active', () => {
      expect(row('Flip horizontal').props.hint).toBeUndefined();
      expect(row('Flip vertical').props.hint).toBeUndefined();
      row('Flip horizontal').props.onClick();
      expect(updateElement).toHaveBeenLastCalledWith('i', { flipH: true });
      row('Flip vertical').props.onClick();
      expect(updateElement).toHaveBeenLastCalledWith('i', { flipV: true });
      element = { ...element, flipH: true, flipV: true };
      expect(row('Flip horizontal').props.hint).toBe('On');
      expect(row('Flip vertical').props.hint).toBe('On');
      row('Flip horizontal').props.onClick();
      expect(updateElement).toHaveBeenLastCalledWith('i', { flipH: false });
      row('Flip vertical').props.onClick();
      expect(updateElement).toHaveBeenLastCalledWith('i', { flipV: false });
    });
  });
});

describe('LiveImageCropFields', () => {
  it('reads the current crop from the model and writes edits back', () => {
    element = { id: 'i', type: 'image', src: 'x', crop: { l: 1, t: 2, r: 3, b: 4 } };
    const out = LiveImageCropFields({ id: 'i' }) as React.ReactElement<P>;
    expect(out.type).toBe(ImageCropFields);
    expect(out.props.crop).toEqual({ l: 1, t: 2, r: 3, b: 4 });
    expect(out.props.idPrefix).toBe('crop-i');
    out.props.onChange({ l: 0, t: 0, r: 0, b: 0 });
    expect(updateElement).toHaveBeenCalledWith('i', { crop: { l: 0, t: 0, r: 0, b: 0 } });
  });
  it('renders nothing once the image is gone or is not an image', () => {
    element = undefined;
    expect(LiveImageCropFields({ id: 'i' })).toBeNull();
    element = { id: 'i', type: 'shape' } as never;
    expect(LiveImageCropFields({ id: 'i' })).toBeNull();
  });
});
