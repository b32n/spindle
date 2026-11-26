// Content controls (Word structured document tags): the model, the ProseMirror nodes, the Y document, collaboration
// blockers and the DOCX mapping. The io package tests the file format itself.

import * as Y from 'yjs';
import { EditorState, TextSelection } from 'prosemirror-state';
import { updateYFragment, yXmlFragmentToProseMirrorRootNode } from 'y-prosemirror';
import { InMemoryProvider, __resetInMemoryRooms } from '@b32nio/spindle-shared';
import { docsSchema } from './prosemirror/schema';
import { blocksToPmDoc, proseMirrorToDocument } from './prosemirror/sync';
import { DocumentImpl } from './document';
import { getCollabBlockers } from './collab/blockers';
import { getYDocFields, hydrateYDocFromData, serializeYDocToData } from './collab/y-schema';
import { exportToDocx, importFromDocx } from './export/docx';
import { CHECKBOX_CHECKED_GLYPH, CHECKBOX_UNCHECKED_GLYPH, findCheckboxControl, toggleCheckboxControl } from './prosemirror/commands';
import { getPlainText, isBlockEmpty } from './blocks/utils';
import { getBlockText, getBlockTextLength } from './selection';
import { DEFAULT_PAGE_CONFIG } from './types';
import type { Block, ContentControl, ContentControlBlock, DocumentData, InlineContent, ParagraphBlock } from './types';

const text = (t: string): InlineContent => ({ type: 'text', text: t });
const para = (...content: InlineContent[]): ParagraphBlock => ({ id: 'p', type: 'paragraph', content });

const FULL_CONTROL: ContentControl = {
  type: 'dropDown',
  tag: 'berth',
  alias: 'Berth',
  lock: 'sdtLocked',
  placeholder: 'DefaultPlaceholder_1',
  showingPlaceholder: true,
  items: [
    { text: 'North', value: 'N' },
    { text: 'East', value: 'E' },
  ],
  dateFormat: 'd MMMM yyyy',
  dateValue: '2026-03-02T00:00:00Z',
  multiline: true,
  gallery: 'Cover Pages',
};

const CHECKBOX: ContentControl = { type: 'checkbox', tag: 'ok', alias: 'OK', checked: true };

/** The blocks of a document without ids, which are generated on every conversion. */
function strip(blocks: Block[]): unknown {
  return JSON.parse(JSON.stringify(blocks, (k, v) => (k === 'id' ? undefined : v)));
}

function dataWith(blocks: Block[]): DocumentData {
  return {
    id: 'doc',
    title: 'Doc',
    sections: [{ id: 's1', pageConfig: DEFAULT_PAGE_CONFIG, blocks }],
    defaultPageConfig: DEFAULT_PAGE_CONFIG,
    textStylePool: {},
    paragraphStylePool: {},
  };
}

/** blocks -> PM doc -> blocks. */
function viaPm(blocks: Block[]): Block[] {
  const existing = dataWith([]) as unknown as Parameters<typeof proseMirrorToDocument>[1];
  return proseMirrorToDocument(blocksToPmDoc(blocks), existing).sections[0].blocks;
}

/** blocks -> Y.Doc -> blocks (the collab hydrate and serialize path). */
function viaYjs(blocks: Block[]): Block[] {
  const ydoc = new Y.Doc();
  hydrateYDocFromData(ydoc, dataWith(blocks));
  return serializeYDocToData(ydoc).sections[0].blocks;
}

const SAMPLE: Block[] = [
  para(text('Name: '), { type: 'contentControl', control: { type: 'text', tag: 'name', alias: 'Name' }, content: [text('Ada '), { type: 'text', text: 'Marlowe', bold: true }] }, text('.')),
  {
    id: 'cc',
    type: 'content-control',
    control: { type: 'richText', tag: 'outer' },
    blocks: [
      para(text('inside')),
      { id: 'inner', type: 'content-control', control: { type: 'richText', tag: 'inner' }, blocks: [para(text('deep'))] },
      { id: 'li', type: 'list-item', listType: 'bullet', level: 0, content: [text('item')] },
    ],
  },
  para(
    { type: 'contentControl', control: CHECKBOX, content: [text('☒')] },
    { type: 'contentControl', control: { type: 'date', dateFormat: 'dd/MM/yyyy', dateValue: '2026-03-02T00:00:00Z' }, content: [] }
  ),
];

describe('content controls in the ProseMirror document', () => {
  it('round-trips block and inline controls, nested ones and every property', () => {
    const blocks: Block[] = [
      para({ type: 'contentControl', control: FULL_CONTROL, content: [text('x')] }),
      { id: 'cc', type: 'content-control', control: FULL_CONTROL, blocks: [para(text('y'))] },
    ];
    expect(strip(viaPm(blocks))).toEqual(strip(blocks));
    expect(strip(viaPm(SAMPLE))).toEqual(strip(SAMPLE));
  });

  it('keeps a checkbox state, including unchecked, and drops properties that do not apply', () => {
    const unchecked: Block[] = [para({ type: 'contentControl', control: { type: 'checkbox', checked: false }, content: [text('☐')] })];
    expect(strip(viaPm(unchecked))).toEqual(strip(unchecked));
    const out = viaPm([para({ type: 'contentControl', control: { type: 'richText', checked: true }, content: [] })]);
    expect(((out[0] as ParagraphBlock).content[0] as { control: ContentControl }).control).toEqual({ type: 'richText' });
  });

  it('puts a paragraph in a block control that has no blocks, and reads an unknown type as rich text', () => {
    const pm = blocksToPmDoc([{ id: 'cc', type: 'content-control', control: { type: 'richText' }, blocks: [] }]);
    expect(pm.firstChild?.type.name).toBe('content_control');
    expect(pm.firstChild?.childCount).toBe(1);
    const odd = docsSchema.nodes.content_control.create({ controlType: 'mystery', lock: 'weird', items: ['x'] }, [docsSchema.nodes.paragraph.create()]);
    const doc = docsSchema.nodes.doc.create(null, [odd]);
    const blocks = proseMirrorToDocument(doc, dataWith([]) as never).sections[0].blocks;
    expect((blocks[0] as ContentControlBlock).control).toEqual({ type: 'richText' });
  });

  it('holds controls at the top level, in table cells and in other controls, but not in a list item', () => {
    const { content_control: cc, paragraph, bullet_list, list_item, table, table_row, table_cell } = docsSchema.nodes;
    const p = () => paragraph.create();
    expect(() => docsSchema.nodes.doc.create(null, [cc.create(null, [p()])]).check()).not.toThrow();
    expect(() => table.create(null, [table_row.create(null, [table_cell.create(null, [cc.create(null, [p()])])])]).check()).not.toThrow();
    expect(() => cc.create(null, [cc.create(null, [p()])]).check()).not.toThrow();
    expect(() => bullet_list.create(null, [list_item.create(null, [p(), cc.create(null, [p()])])]).check()).toThrow();
    // An inline control sits in a paragraph, and in another inline control.
    const inline = docsSchema.nodes.content_control_inline;
    expect(() => paragraph.create(null, [inline.create(null, [inline.create(null, [docsSchema.text('a')])])]).check()).not.toThrow();
  });

  it('describes itself in the DOM with every property and the title, and reads that description back', () => {
    const toDom = (name: string, attrs: Record<string, unknown>) => {
      const node = docsSchema.nodes[name].create(attrs, name === 'content_control' ? [docsSchema.nodes.paragraph.create()] : undefined);
      return docsSchema.nodes[name].spec.toDOM!(node) as [string, Record<string, string>, number];
    };
    const [tag, domAttrs] = toDom('content_control_inline', { controlType: 'text', alias: 'Name', tag: 'name' });
    expect(tag).toBe('span');
    expect(domAttrs.title).toBe('Name');
    expect(JSON.parse(domAttrs['data-content-control'])).toMatchObject({ controlType: 'text', alias: 'Name', tag: 'name', checked: false });
    expect(toDom('content_control', { controlType: 'richText' })[0]).toBe('div');
    expect(toDom('content_control', { controlType: 'richText' })[1].title).toBeUndefined();

    const rule = docsSchema.nodes.content_control_inline.spec.parseDOM![0];
    const read = (json: string) => (rule.getAttrs as (dom: unknown) => unknown)({ getAttribute: () => json });
    expect(read(domAttrs['data-content-control'])).toMatchObject({ controlType: 'text', alias: 'Name', tag: 'name' });
  });

  it('does not trust pasted HTML: bad JSON is not a control, unknown or mistyped fields are dropped', () => {
    const rule = docsSchema.nodes.content_control.spec.parseDOM![0];
    const read = (json: string) => (rule.getAttrs as (dom: unknown) => unknown)({ getAttribute: () => json });
    expect(read('{nope')).toBe(false);
    expect(read('5')).toBe(false);
    expect(read('')).toBe(false);
    const odd = read('{"controlType":"weird","tag":7,"lock":"x","items":[{"text":"a","value":"b"},{"text":1}],"checked":"yes","evil":"<script>"}') as Record<string, unknown>;
    expect(odd).toMatchObject({ controlType: 'richText', tag: null, lock: null, items: [{ text: 'a', value: 'b' }], checked: false });
    expect(odd).not.toHaveProperty('evil');
  });
});

describe('content controls in the Y document', () => {
  it('survive hydrate and serialize', () => {
    expect(strip(viaYjs(SAMPLE))).toEqual(strip(SAMPLE));
  });

  it('writes the schema version that introduced the control nodes', async () => {
    const { COLLAB_SCHEMA_VERSION } = await import('./collab/y-schema');
    expect(COLLAB_SCHEMA_VERSION).toBeGreaterThanOrEqual(3);
  });

  it('two clients converge: one toggles a checkbox while the other types inside another control', () => {
    const blocks = SAMPLE.slice();
    const a = new Y.Doc();
    hydrateYDocFromData(a, dataWith(blocks));
    const b = new Y.Doc();
    Y.applyUpdate(b, Y.encodeStateAsUpdate(a));

    const edit = (ydoc: Y.Doc, fn: (state: EditorState) => ReturnType<EditorState['tr']['replaceWith']>) => {
      const fragment = getYDocFields(ydoc).content;
      const state = EditorState.create({ doc: yXmlFragmentToProseMirrorRootNode(fragment, docsSchema) });
      updateYFragment(ydoc, fragment, fn(state).doc, { mapping: new Map(), isOMark: new Map() });
    };
    const findInline = (state: EditorState, controlType: string): number => {
      let found = -1;
      state.doc.descendants((node, pos) => {
        if (node.type.name === 'content_control_inline' && node.attrs.controlType === controlType) found = pos;
      });
      return found;
    };
    // A: unchecks the checkbox and swaps its glyph. B: types in the text control.
    edit(a, (state) => {
      const pos = findInline(state, 'checkbox');
      const node = state.doc.nodeAt(pos)!;
      return state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, checked: false }).replaceWith(pos + 1, pos + 2, docsSchema.text('☐'));
    });
    edit(b, (state) => state.tr.insertText('X', findInline(state, 'text') + 1));
    Y.applyUpdate(b, Y.encodeStateAsUpdate(a, Y.encodeStateVector(b)));
    Y.applyUpdate(a, Y.encodeStateAsUpdate(b, Y.encodeStateVector(a)));

    const dataA = serializeYDocToData(a).sections[0].blocks;
    const dataB = serializeYDocToData(b).sections[0].blocks;
    expect(strip(dataA)).toEqual(strip(dataB));
    const first = (dataA[0] as ParagraphBlock).content[1] as { content: InlineContent[] };
    expect(getPlainText(first.content)).toBe('XAda Marlowe');
    const last = (dataA[2] as ParagraphBlock).content[0] as { control: ContentControl; content: InlineContent[] };
    expect(last.control.checked).toBe(false);
    expect(getPlainText(last.content)).toBe('☐');
  });

  it('a room attached through the collab binding carries the controls to a joiner', async () => {
    __resetInMemoryRooms();
    const creator = new DocumentImpl('room', 'Doc');
    creator.setData({ ...dataWith(SAMPLE), id: 'room' });
    const identity = (n: string) => ({ userId: n, displayName: n, color: '#111' });
    const ha = await creator.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const joiner = new DocumentImpl('room', 'Doc');
    const hb = await joiner.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    expect(strip(serializeYDocToData(hb.ydoc).sections[0].blocks)).toEqual(strip(SAMPLE));
    ha.detach();
    hb.detach();
  });
});

describe('collab blockers for content controls', () => {
  it('knows the control block and inline types, so a document with controls is not blocked', () => {
    expect(getCollabBlockers(dataWith(SAMPLE))).toEqual([]);
  });

  it('looks inside controls: an unknown type in a block control or in an inline control is reported', () => {
    const data = dataWith([
      { id: 'cc', type: 'content-control', control: { type: 'richText' }, blocks: [{ id: 'x', type: 'chart' } as never, para(text('a'))] },
      para({ type: 'contentControl', control: { type: 'text' }, content: [{ type: 'smartArt' } as never] }),
    ]);
    expect(getCollabBlockers(data).map((b) => [b.id, b.count])).toEqual([
      ['collab.unsupportedBlock', 1],
      ['collab.unsupportedInline', 1],
    ]);
  });

  it('looks inside controls that sit in table cells', () => {
    const data = dataWith([
      {
        id: 't',
        type: 'table',
        rows: [{ id: 'r', cells: [{ id: 'c', content: [], blocks: [{ id: 'cc', type: 'content-control', control: { type: 'richText' }, blocks: [{ id: 'y', type: 'equation' } as never] }] }] }],
      },
    ]);
    expect(getCollabBlockers(data).map((b) => [b.id, b.count])).toEqual([['collab.unsupportedBlock', 1]]);
  });
});

describe('content controls in the block helpers', () => {
  it('counts the text of an inline control and treats a block control as a container', () => {
    const inline: InlineContent = { type: 'contentControl', control: { type: 'text' }, content: [text('ab'), { type: 'contentControl', control: { type: 'text' }, content: [text('cd')] }] };
    expect(getPlainText([text('x'), inline])).toBe('xabcd');
    const empty: Block = { id: 'c', type: 'content-control', control: { type: 'richText' }, blocks: [para()] };
    expect(isBlockEmpty(empty)).toBe(true);
    expect(isBlockEmpty({ ...empty, blocks: [para(text('a'))] })).toBe(false);
    expect(getBlockTextLength(empty)).toBe(0);
    expect(getBlockText(empty)).toBe('');
  });
});

describe('content controls through DOCX', () => {
  it('round-trip through export and import, with block and inline controls', async () => {
    const doc = new DocumentImpl('doc', 'Doc');
    doc.setData({ ...dataWith(SAMPLE.map((b, i) => ({ ...b, id: `b${i}` }))) });
    const { bytes } = await exportToDocx(doc);
    const reopened = new DocumentImpl('doc2', 'Doc2');
    await importFromDocx(bytes, reopened);
    expect(strip(reopened.getSections()[0].blocks)).toEqual(strip(SAMPLE));
  });

  it('a header or footer holds no controls: an imported one is read as its content', async () => {
    const doc = new DocumentImpl('doc', 'Doc');
    doc.setData({
      ...dataWith([para(text('body'))]),
      sections: [
        {
          id: 's1',
          pageConfig: DEFAULT_PAGE_CONFIG,
          blocks: [para(text('body'))],
          header: { blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'Head' }] }] },
        },
      ],
    });
    const { bytes } = await exportToDocx(doc);
    const reopened = new DocumentImpl('doc2', 'Doc2');
    await importFromDocx(bytes, reopened);
    expect(reopened.getSections()[0].header?.blocks[0].content).toEqual([expect.objectContaining({ type: 'text', text: 'Head' })]);
  });
});

describe('toggling a check box content control', () => {
  const { paragraph, content_control_inline: inlineControl } = docsSchema.nodes;
  const box = (checked: boolean, glyph: string, marks = docsSchema.marks.bold.create()) =>
    inlineControl.create({ controlType: 'checkbox', checked, alias: 'Box' }, [docsSchema.text(glyph, [marks])]);
  const stateWith = (...inline: Parameters<typeof paragraph.create>[1][]) =>
    EditorState.create({ doc: docsSchema.nodes.doc.create(null, [paragraph.create(null, inline as never)]) });
  /** Position of the control in a paragraph that starts with the text "a " (the paragraph starts at 0). */
  const AT = 3;

  it('flips the state and the glyph from a position inside the control, keeping the formatting', () => {
    const state = stateWith(docsSchema.text('a '), box(false, CHECKBOX_UNCHECKED_GLYPH));
    let next = state;
    expect(toggleCheckboxControl(AT + 1)(state, (tr) => (next = state.apply(tr)))).toBe(true);
    const control = next.doc.nodeAt(AT)!;
    expect(control.attrs.checked).toBe(true);
    expect(control.textContent).toBe(CHECKBOX_CHECKED_GLYPH);
    expect(control.firstChild!.marks.map((m) => m.type.name)).toEqual(['bold']);
    // And back.
    let again = next;
    toggleCheckboxControl(AT + 1)(next, (tr) => (again = next.apply(tr)));
    expect(again.doc.nodeAt(AT)!.attrs.checked).toBe(false);
    expect(again.doc.nodeAt(AT)!.textContent).toBe(CHECKBOX_UNCHECKED_GLYPH);
  });

  it('also works from the position just before the control, and from the selection by default', () => {
    const state = stateWith(docsSchema.text('a '), box(true, CHECKBOX_CHECKED_GLYPH));
    let next = state;
    expect(toggleCheckboxControl(AT)(state, (tr) => (next = state.apply(tr)))).toBe(true);
    expect(next.doc.nodeAt(AT)!.attrs.checked).toBe(false);
    const inside = state.apply(state.tr.setSelection(TextSelection.create(state.doc, AT + 1)));
    expect(toggleCheckboxControl()(inside)).toBe(true);
  });

  it('is only applicable when a check box is there, and does nothing without a dispatch', () => {
    const text = stateWith(docsSchema.text('plain'));
    expect(toggleCheckboxControl(2)(text)).toBe(false);
    expect(toggleCheckboxControl(-1)(text)).toBe(false);
    expect(toggleCheckboxControl(999)(text)).toBe(false);
    const other = stateWith(inlineControl.create({ controlType: 'text' }, [docsSchema.text('x')]));
    expect(toggleCheckboxControl(2)(other)).toBe(false);
    expect(findCheckboxControl(other.doc, 2)).toBeNull();
  });

  it('toggles the innermost check box of nested controls, and clears the placeholder flag', () => {
    const inner = inlineControl.create({ controlType: 'checkbox', checked: false, showingPlaceholder: true }, [docsSchema.text(CHECKBOX_UNCHECKED_GLYPH)]);
    const outer = inlineControl.create({ controlType: 'checkbox', checked: false }, [inner]);
    const state = stateWith(outer);
    expect(findCheckboxControl(state.doc, 3)!.node.attrs.showingPlaceholder).toBe(true);
    let next = state;
    toggleCheckboxControl(3)(state, (tr) => (next = state.apply(tr)));
    const toggled = next.doc.nodeAt(2)!;
    expect(toggled.attrs).toMatchObject({ checked: true, showingPlaceholder: false });
    expect(next.doc.nodeAt(1)!.attrs.checked).toBe(false);
  });
});
