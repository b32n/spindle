import React from 'react';
import type { BodyStyle } from '@b32nio/spindle-slides-core';
import { Popover } from './Popover';
import { TextBodyFields } from './TextBodyFields';

const updateElement = jest.fn();
const refit = jest.fn();
let selected: string[] = ['t'];
let editingId: string | null = null;
let element: Record<string, unknown> | undefined;
let menu = { inMenu: false, openModal: jest.fn(), closeModal: jest.fn(), closeMenu: jest.fn() };

jest.mock('../hooks', () => ({
  useDeck: () => ({ updateElement }),
  useSelection: () => ({ slideId: 's', elementIds: selected }),
  useEditingId: () => editingId,
  useElement: () => element,
}));
let editingView: unknown = null;
const applyFormatMock = jest.fn();
jest.mock('../context/DeckContext', () => ({ useDeckContext: () => ({ editing: { getView: () => editingView } }) }));
const insertFieldMock = jest.fn();
jest.mock('../interactions/formatting', () => ({
  applyFormat: (...a: unknown[]) => applyFormatMock(...a),
  insertField: (...a: unknown[]) => insertFieldMock(...a),
}));
jest.mock('../text-fit', () => ({ refitElement: (...a: unknown[]) => refit(...a) }));
jest.mock('@b32nio/spindle-shared/react', () => ({
  ...jest.requireActual('@b32nio/spindle-shared/react'),
  useToolbarMenu: () => menu,
}));

import { MenuRow } from '@b32nio/spindle-shared/react';
import { ParagraphFields } from './ParagraphFields';
import { RunFields } from './RunFields';
import { TextOptionsButton, LiveTextBodyFields, LiveParagraphFields, LiveRunFields } from './TextOptionsButton';

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

let slots: unknown[] = [];
let n = 0;
const render = () => {
  n = 0;
  return TextOptionsButton();
};

beforeEach(() => {
  slots = [];
  updateElement.mockClear();
  refit.mockClear();
  selected = ['t'];
  editingId = null;
  element = { id: 't', type: 'text', bodyStyle: { autofit: { mode: 'shrink' } } };
  menu = { inMenu: false, openModal: jest.fn(), closeModal: jest.fn(), closeMenu: jest.fn() };
  jest.spyOn(React, 'useState').mockImplementation(((init: unknown) => {
    const i = n++;
    if (!(i in slots)) slots[i] = typeof init === 'function' ? (init as () => unknown)() : init;
    return [slots[i], (v: unknown) => { slots[i] = typeof v === 'function' ? (v as (p: unknown) => unknown)(slots[i]) : v; }];
  }) as never);
  (globalThis as { requestAnimationFrame?: unknown }).requestAnimationFrame = (cb: () => void) => { cb(); return 0; };
});
afterEach(() => jest.restoreAllMocks());

describe('TextOptionsButton', () => {
  it('renders nothing without a single text box (or the text of a shape that is not being edited)', () => {
    selected = [];
    expect(render()).toBeNull();
    selected = ['t', 'u'];
    expect(render()).toBeNull();
    selected = ['t'];
    element = undefined;
    expect(render()).toBeNull();
    element = { id: 't', type: 'image' };
    expect(render()).toBeNull();
    element = { id: 't', type: 'shape' };
    expect(render()).toBeNull();
    editingId = 't';
    expect(render()).not.toBeNull();
  });

  it('inline: the button toggles an anchored popover with live fields', () => {
    const [btn] = find(render(), byTitle('Text box options'));
    expect(btn.props.active).toBe(false);
    btn.props.onClick();
    const [pop] = find(render(), (e) => e.type === Popover);
    expect(find(pop, (e) => e.type === LiveTextBodyFields)[0].props.id).toBe('t');
    expect(find(pop, (e) => e.type === LiveParagraphFields)[0].props.id).toBe('t');
    expect(find(pop, (e) => e.type === LiveRunFields)[0].props.id).toBe('t');
    pop.props.onClose();
    expect(find(render(), (e) => e.type === Popover)).toHaveLength(0);
    find(render(), byTitle('Text box options'))[0].props.onClick();
    find(render(), byTitle('Text box options'))[0].props.onClick();
    expect(find(render(), (e) => e.type === Popover)).toHaveLength(0);
  });

  it('the popover closes when the selection moves to another box', () => {
    find(render(), byTitle('Text box options'))[0].props.onClick();
    selected = ['u'];
    element = { id: 'u', type: 'text' };
    expect(find(render(), (e) => e.type === Popover)).toHaveLength(0);
  });

  it('in the overflow menu: a single row opens a modal and hints the autofit mode', () => {
    menu.inMenu = true;
    const hints: Array<[Record<string, unknown> | undefined, unknown]> = [
      [{ autofit: { mode: 'shrink' } }, 'Shrink'],
      [{ autofit: { mode: 'resize' } }, 'Resize'],
      [{ autofit: { mode: 'none' } }, undefined],
      [undefined, undefined],
    ];
    for (const [bodyStyle, hint] of hints) {
      element = { id: 't', type: 'text', bodyStyle };
      const row = render() as React.ReactElement<P>;
      expect(row.type).toBe(MenuRow);
      expect(row.props.label).toBe('Text box options');
      expect(row.props.hint).toBe(hint);
    }
    (render() as React.ReactElement<P>).props.onClick();
    const [title, content] = menu.openModal.mock.calls[0];
    expect(title).toBe('Text box options');
    expect(find(content as React.ReactNode, (e) => e.type === LiveTextBodyFields)).toHaveLength(1);
    expect(find(content as React.ReactNode, (e) => e.type === LiveParagraphFields)[0].props.id).toBe('t');
    expect(find(content as React.ReactNode, (e) => e.type === LiveRunFields)[0].props.id).toBe('t');
  });
});

describe('LiveTextBodyFields', () => {
  const fields = (id = 't') => LiveTextBodyFields({ id }) as React.ReactElement<P>;
  it('renders nothing for a missing or non-text element', () => {
    element = undefined;
    expect(fields()).toBeNull();
    element = { id: 't', type: 'image' };
    expect(fields()).toBeNull();
  });
  it('allows resizing for a text box', () => {
    element = { id: 't', type: 'text' };
    expect(fields().props.allowResize).toBe(true);
  });
  it('applies a patch merged into the stored body and refits when layout changes', () => {
    element = { id: 't', type: 'shape', bodyStyle: { padding: 3 } as BodyStyle };
    const tree = fields();
    expect(tree.type).toBe(TextBodyFields);
    expect(tree.props.allowResize).toBe(false); // a shape does not resize to its text
    tree.props.onChange({ autofit: { mode: 'shrink' } });
    expect(updateElement).toHaveBeenLastCalledWith('t', { bodyStyle: { padding: 3, autofit: { mode: 'shrink' } } });
    expect(refit).toHaveBeenCalledTimes(1);
    for (const patch of [{ insets: { l: 1, t: 1, r: 1, b: 1 } }, { vert: 'vert' as const }, { vert: undefined }, { columns: { count: 2 } }, { anchorCtr: true }, { anchorCtr: undefined }]) {
      tree.props.onChange(patch);
    }
    expect(refit).toHaveBeenCalledTimes(7);
    tree.props.onChange({ padding: 9 });
    expect(refit).toHaveBeenCalledTimes(7);
  });
});

describe('LiveParagraphFields', () => {
  const fields = (id = 't') => LiveParagraphFields({ id }) as React.ReactElement<P>;
  afterEach(() => {
    editingView = null;
    applyFormatMock.mockClear();
  });
  it('renders nothing for a missing or non-text element', () => {
    element = undefined;
    expect(fields()).toBeNull();
    element = { id: 't', type: 'image' };
    expect(fields()).toBeNull();
  });
  it('idle: shows the first paragraph and applies to the element', () => {
    element = { id: 't', type: 'text', richText: { type: 'doc', content: [{ type: 'paragraph', attrs: { marginLeft: 9 } }] } };
    const tree = fields();
    expect(tree.type).toBe(ParagraphFields);
    expect(tree.props.attrs).toEqual({ marginLeft: 9 });
    tree.props.onChange({ marginLeft: 12 });
    expect(applyFormatMock).toHaveBeenCalledWith({ editingView: null, deck: expect.anything(), elementId: 't' }, { paragraph: { marginLeft: 12 } });
    element = { id: 't', type: 'shape' };
    expect(fields().props.attrs).toBeUndefined();
  });
  it('editing: shows the paragraph holding the caret and applies through the view', () => {
    element = { id: 't', type: 'text', richText: { type: 'doc', content: [{ type: 'paragraph', attrs: { marginLeft: 9 } }] } };
    editingView = { state: { selection: { $from: { parent: { type: { name: 'paragraph' }, attrs: { marginLeft: 33 } } } } } };
    const tree = fields();
    expect(tree.props.attrs).toEqual({ marginLeft: 33 });
    tree.props.onChange({ marginLeft: 1 });
    expect(applyFormatMock.mock.calls[0][0].editingView).toBe(editingView);
    editingView = { state: { selection: { $from: { parent: { type: { name: 'doc' }, attrs: {} } } } } };
    expect(fields().props.attrs).toEqual({ marginLeft: 9 });
  });
});

describe('LiveRunFields', () => {
  const fields = (id = 't') => LiveRunFields({ id }) as React.ReactElement<P>;
  const bold = { type: 'text', text: 'a', marks: [{ type: 'baseline', attrs: { pct: 30 } }] };
  afterEach(() => {
    editingView = null;
    applyFormatMock.mockClear();
    insertFieldMock.mockClear();
  });
  it('renders nothing for a missing or non-text element', () => {
    element = undefined;
    expect(fields()).toBeNull();
    element = { id: 't', type: 'image' };
    expect(fields()).toBeNull();
  });
  it('idle: shows the first run, applies marks to the element, cannot insert a field, and refits', () => {
    element = { id: 't', type: 'text', richText: { type: 'doc', content: [{ type: 'paragraph', content: [bold] }] } };
    const tree = fields();
    expect(tree.type).toBe(RunFields);
    expect(tree.props.run.baselinePct).toBe(30);
    expect(tree.props.onInsertField).toBeUndefined();
    tree.props.onMark({ type: 'caps', attrs: { kind: 'all' } });
    tree.props.onMark({ type: 'caps', attrs: null });
    expect(applyFormatMock).toHaveBeenNthCalledWith(1, { editingView: null, deck: expect.anything(), elementId: 't' }, { setMark: { type: 'caps', attrs: { kind: 'all' } } });
    expect(applyFormatMock).toHaveBeenNthCalledWith(2, expect.anything(), { removeMark: 'caps' });
    expect(refit).toHaveBeenCalledTimes(2);
    element = { id: 't', type: 'shape' };
    expect(fields().props.run.baselinePct).toBeNull();
  });
  it('editing: shows the marks at the caret (stored marks first) and inserts fields through the view', () => {
    element = { id: 't', type: 'text' };
    const mk = (pct: number) => ({ type: { name: 'baseline' }, attrs: { pct } });
    editingView = { state: { storedMarks: null, selection: { $from: { marks: () => [mk(30)] } } } };
    expect(fields().props.run.baselinePct).toBe(30);
    editingView = { state: { storedMarks: [mk(-25)], selection: { $from: { marks: () => [mk(30)] } } } };
    const tree = fields();
    expect(tree.props.run.baselinePct).toBe(-25);
    tree.props.onInsertField({ kind: 'slidenum' });
    expect(insertFieldMock).toHaveBeenCalledWith(editingView, { kind: 'slidenum' });
    expect(refit).toHaveBeenCalledTimes(1);
  });
});
