// Two-client convergence tests for the docs collab pipeline.
//
// Each client owns a Y.Doc and edits it the way ySyncPlugin does: the Y.XmlFragment
// is read into a ProseMirror doc (yXmlFragmentToProseMirrorRootNode), edited with
// an ordinary transaction, and diffed back with updateYFragment. Peers exchange
// Y updates; "concurrent" edits are made while the clients are disconnected.
// The model path (hydrateYDocFromData / serializeYDocToData) is the one used for
// seeding the room and for getData() under collab.

import * as Y from 'yjs';
import { EditorState, type Transaction } from 'prosemirror-state';
import { Node as PmNode } from 'prosemirror-model';
import { updateYFragment, yXmlFragmentToProseMirrorRootNode } from 'y-prosemirror';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { docsSchema } from '../prosemirror/schema';
import { sectionsToPmDoc } from '../prosemirror/sync';
import { DocumentImpl } from '../document';
import { createParagraphFromText } from '../blocks/paragraph';
import { getYDocFields, hydrateYDocFromData, serializeYDocToData } from './y-schema';
import { DEFAULT_PAGE_CONFIG } from '../types';
import type { Block, DocumentData, InlineContent, ParagraphBlock } from '../types';

class Client {
  readonly ydoc = new Y.Doc();
  constructor(readonly name: string) {}

  get fragment(): Y.XmlFragment {
    return getYDocFields(this.ydoc).content;
  }
  pm(): PmNode {
    return yXmlFragmentToProseMirrorRootNode(this.fragment, docsSchema);
  }
  /** Apply a ProseMirror transaction and push the result into the Y.Doc. */
  edit(fn: (tr: Transaction) => void): void {
    const state = EditorState.create({ doc: this.pm() });
    const tr = state.tr;
    fn(tr);
    updateYFragment(this.ydoc, this.fragment, tr.doc, {
      mapping: new Map(),
      isOMark: new Map(),
    });
  }
  data(): DocumentData {
    return serializeYDocToData(this.ydoc);
  }
  blocks(): Block[] {
    return this.data().sections[0].blocks;
  }
}

function sync(a: Client, b: Client): void {
  const ua = Y.encodeStateAsUpdate(a.ydoc, Y.encodeStateVector(b.ydoc));
  const ub = Y.encodeStateAsUpdate(b.ydoc, Y.encodeStateVector(a.ydoc));
  Y.applyUpdate(b.ydoc, ua);
  Y.applyUpdate(a.ydoc, ub);
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

/** Two clients sharing a seeded room. */
function pair(blocks: Block[]): [Client, Client] {
  const a = new Client('A');
  hydrateYDocFromData(a.ydoc, dataWith(blocks));
  const b = new Client('B');
  sync(a, b);
  return [a, b];
}

const para = (id: string, content: InlineContent[]): ParagraphBlock => ({
  id,
  type: 'paragraph',
  content,
});
const text = (t: string): InlineContent => ({ type: 'text', text: t });

const IMAGE: InlineContent = {
  type: 'image',
  src: 'data:image/png;base64,iVBORw0KGgo=',
  width: 120,
  height: 80,
  alt: 'logo',
};
const PAGE_NUMBER: InlineContent = { type: 'dynamicField', fieldType: 'pageNumber' };
const DATE_FIELD: InlineContent = { type: 'dynamicField', fieldType: 'date', format: 'MM/DD/YYYY' };

const inlineOf = (blocks: Block[], id = 0): InlineContent[] =>
  (blocks[id] as ParagraphBlock).content;
const kinds = (c: InlineContent[]): string[] =>
  c.map((i) => (i.type === 'dynamicField' ? `field:${i.fieldType}` : i.type));

describe('inline image / page-number field survive collab', () => {
  it('seeded inline nodes reach a joining peer with attributes', () => {
    const [a, b] = pair([para('p', [text('a '), IMAGE, text(' b '), PAGE_NUMBER, DATE_FIELD])]);
    for (const c of [a, b]) {
      const content = inlineOf(c.blocks());
      expect(kinds(content)).toEqual(['text', 'image', 'text', 'field:pageNumber', 'field:date']);
      expect(content[1]).toMatchObject({ src: IMAGE.type === 'image' ? IMAGE.src : '', width: 120, height: 80, alt: 'logo' });
      expect(content[4]).toMatchObject({ fieldType: 'date', format: 'MM/DD/YYYY' });
    }
  });

  it('inserting an inline image and field on A reaches B and survives later edits', () => {
    const [a, b] = pair([para('p', [text('hello world')])]);
    a.edit((tr) => {
      tr.insert(
        6,
        docsSchema.nodes.inline_image.create({ src: 'data:image/png;base64,AAAA', width: 50, height: 40, alt: 'x' }),
      );
      tr.insert(1, docsSchema.nodes.dynamic_field.create({ fieldType: 'pageNumber', format: null }));
    });
    sync(a, b);
    // sequential edit on B, then on A
    b.edit((tr) => tr.insertText('!', b.pm().content.size - 1));
    sync(a, b);
    a.edit((tr) => tr.insertText('>', 1));
    sync(a, b);

    expect(b.pm().toJSON()).toEqual(a.pm().toJSON());
    for (const c of [a, b]) {
      const content = inlineOf(c.blocks());
      expect(kinds(content).filter((k) => k !== 'text')).toEqual(['field:pageNumber', 'image']);
      const img = content.find((i) => i.type === 'image');
      expect(img).toMatchObject({ src: 'data:image/png;base64,AAAA', width: 50, height: 40, alt: 'x' });
    }
  });

  it('concurrent edits on both clients converge and keep the nodes', () => {
    const [a, b] = pair([para('p', [text('one '), IMAGE, text(' two '), PAGE_NUMBER, text(' three')])]);
    a.edit((tr) => tr.insertText('A', 1));
    b.edit((tr) => tr.insertText('B', b.pm().content.size - 1));
    // B also inserts a new field concurrently
    b.edit((tr) => tr.insert(3, docsSchema.nodes.dynamic_field.create({ fieldType: 'totalPages' })));
    sync(a, b);

    expect(a.pm().toJSON()).toEqual(b.pm().toJSON());
    for (const c of [a, b]) {
      const content = inlineOf(c.blocks());
      const nonText = kinds(content).filter((k) => k !== 'text');
      expect(nonText.sort()).toEqual(['field:pageNumber', 'field:totalPages', 'image']);
      expect(content.find((i) => i.type === 'image')).toMatchObject({ width: 120, height: 80, alt: 'logo' });
    }
  });

  it('bookmarks hydrate, sync and serialize back unchanged, and an edit next to one leaves it alone', () => {
    const MARK: InlineContent = { type: 'bookmark', name: '_Toc7' };
    const [a, b] = pair([para('p', [MARK, text('Heading'), { type: 'bookmark', name: '_Ref8' }])]);
    for (const c of [a, b]) {
      expect(inlineOf(c.blocks())).toEqual([MARK, text('Heading'), { type: 'bookmark', name: '_Ref8' }]);
    }
    b.edit((tr) => tr.insertText('!', 2));
    sync(a, b);
    expect(inlineOf(a.blocks()).filter((i) => i.type === 'bookmark')).toEqual([MARK, { type: 'bookmark', name: '_Ref8' }]);
  });

  it('generic fields (plain and formatted results) hydrate, sync and serialize back unchanged', () => {
    const REF: InlineContent = { type: 'field', instruction: 'REF _Ref1 \\h', kind: 'REF', result: 'Figure 1' };
    const SEQ: InlineContent = {
      type: 'field',
      instruction: 'SEQ Table',
      kind: 'SEQ',
      result: [{ type: 'text', text: '2', bold: true }, { type: 'text', text: 'a' }],
      dirty: true,
    };
    const [a, b] = pair([para('p', [text('x '), REF, IMAGE, SEQ, PAGE_NUMBER, DATE_FIELD])]);
    for (const c of [a, b]) {
      expect(inlineOf(c.blocks())).toEqual([text('x '), REF, IMAGE, SEQ, PAGE_NUMBER, DATE_FIELD]);
    }
    // an edit next to the atoms leaves them alone
    b.edit((tr) => tr.insertText('!', 2));
    sync(a, b);
    expect(inlineOf(a.blocks()).filter((i) => i.type === 'field')).toEqual([REF, SEQ]);
  });

  it('getData() on a DocumentImpl under collab keeps inline nodes (InMemoryProvider)', async () => {
    __resetInMemoryRooms();
    const id = (n: string): CollabIdentity => ({ userId: n, displayName: n, color: '#111' });
    const mk = (): DocumentImpl => {
      const d = new DocumentImpl('room', 'Doc');
      d.setSectionBlocks(d.getSections()[0].id, [
        createParagraphFromText('x'),
        para('p', [text('see '), IMAGE, text(' page '), PAGE_NUMBER]),
      ]);
      return d;
    };
    const da = mk();
    const ha = await da.attachCollab(new InMemoryProvider(), id('A'), { roomId: 'room' });
    const db = mk();
    const hb = await db.attachCollab(new InMemoryProvider(), id('B'), { roomId: 'room' });
    for (const h of [ha, hb]) {
      const blocks = serializeYDocToData(h.ydoc).sections[0].blocks;
      expect(kinds(inlineOf(blocks, 1))).toEqual(['text', 'image', 'text', 'field:pageNumber']);
    }
    ha.detach();
    hb.detach();
    __resetInMemoryRooms();
  });
});

// ---------------------------------------------------------------------------
// Parametrised: every node and mark type in the docs schema.
// ---------------------------------------------------------------------------

type Json = Record<string, unknown>;
const p = (content: Json[] = [], attrs: Json = {}): Json => ({ type: 'paragraph', attrs, content });
const t = (s: string, marks?: Json[]): Json => (marks ? { type: 'text', text: s, marks } : { type: 'text', text: s });
const li = (level: number, s: string): Json => ({
  type: 'list_item',
  attrs: { level },
  content: [p([t(s)])],
});

/** One sample document per node type (non-default attrs), keyed by node name. */
const NODE_SAMPLES: Record<string, Json[]> = {
  doc: [p([t('plain')])],
  paragraph: [
    p([t('aligned')], { alignment: 'center', indent: 2, spaceBefore: 4, spaceAfter: 12 }),
    // DOCX paragraph formatting: line spacing, left/right/hanging indents, pagination flags, an explicit zero spacing.
    p([t('formatted')], {
      lineSpacing: { type: 'multiple', value: 1.15 },
      spaceAfter: 0,
      leftIndent: 36,
      rightIndent: 12,
      firstLineIndent: -18,
      keepWithNext: true,
      keepTogether: true,
      pageBreakBefore: true,
    }),
    p([t('exact')], { lineSpacing: { type: 'exactly', value: 20 } }),
    // A named paragraph style (DOCX pStyle).
    p([t('callout')], { namedStyleId: 'Callout' }),
  ],
  heading: [
    { type: 'heading', attrs: { level: 3, alignment: 'right' }, content: [t('Title')] },
    {
      type: 'heading',
      attrs: { level: 1, lineSpacing: { type: 'atLeast', value: 24 }, spaceBefore: 16, spaceAfter: 0, leftIndent: 8, keepWithNext: true, pageBreakBefore: true },
      content: [t('Styled')],
    },
    { type: 'heading', attrs: { level: 2, namedStyleId: 'Heading2' }, content: [t('Named')] },
  ],
  bullet_list: [{ type: 'bullet_list', content: [li(0, 'a'), li(2, 'b')] }],
  // A second numbered list right after the first restarts at 1 (it must not merge into the first).
  ordered_list: [
    { type: 'ordered_list', attrs: { start: 4 }, content: [li(0, 'a'), li(1, 'b')] },
    { type: 'ordered_list', content: [li(0, 'again')] },
  ],
  list_item: [
    { type: 'bullet_list', content: [li(3, 'deep')] },
    p([t('between')]), // adjacent bullet lists would merge into one
    {
      type: 'bullet_list',
      content: [
        {
          type: 'list_item',
          attrs: { level: 0 },
          content: [p([t('indented')], { spaceAfter: 0, lineSpacing: { type: 'multiple', value: 1.5 }, leftIndent: 20, keepWithNext: true, namedStyleId: 'ListParagraph' })],
        },
      ],
    },
  ],
  table: [
    {
      type: 'table',
      attrs: {
        colWidths: [100, 60],
        borders: { top: { style: 'double', widthPx: 2, color: '#ff0000' }, insideH: { style: 'none' } },
        look: { firstRow: true, noVBand: true },
        namedStyleId: 'Banded',
      },
      content: [
        {
          type: 'table_row',
          content: [
            { type: 'table_cell', attrs: { borders: { bottom: { style: 'dashed' } }, verticalAlign: 'center', colwidth: [100] }, content: [p([t('c')])] },
          ],
        },
      ],
    },
    {
      type: 'table',
      attrs: { colWidths: [100, 60], styleId: 'ts' },
      content: [
        {
          type: 'table_row',
          attrs: { height: 30 },
          content: [
            { type: 'table_cell', attrs: { colspan: 2, rowspan: 1, backgroundColor: '#ffeeaa', styleId: 'cs' }, content: [p([t('c')])] },
          ],
        },
      ],
    },
  ],
  table_row: [
    {
      type: 'table',
      content: [
        { type: 'table_row', attrs: { height: null, repeatHeader: true, cantSplit: true }, content: [{ type: 'table_cell', content: [p([t('r')])] }] },
      ],
    },
    {
      type: 'table',
      content: [
        { type: 'table_row', attrs: { height: 44 }, content: [{ type: 'table_cell', content: [p([t('r')])] }] },
      ],
    },
  ],
  table_cell: [
    {
      type: 'table',
      content: [
        {
          type: 'table_row',
          content: [{ type: 'table_cell', attrs: { colspan: 1, rowspan: 2, backgroundColor: '#112233', textDirection: 'btLr' }, content: [p([t('cell')])] }],
        },
      ],
    },
    {
      // several blocks, a list and a nested table in one cell
      type: 'table',
      content: [
        {
          type: 'table_row',
          content: [
            {
              type: 'table_cell',
              content: [
                p([t('first')]),
                { type: 'bullet_list', content: [{ type: 'list_item', attrs: { level: 0 }, content: [p([t('item')])] }] },
                { type: 'table', content: [{ type: 'table_row', content: [{ type: 'table_cell', content: [p([t('nested')])] }] }] },
                p([]),
              ],
            },
          ],
        },
      ],
    },
  ],
  table_header: [
    {
      type: 'table',
      content: [
        {
          type: 'table_row',
          content: [{ type: 'table_header', attrs: { colspan: 1, backgroundColor: '#445566' }, content: [p([t('head')])] }],
        },
      ],
    },
  ],
  image: [{ type: 'image', attrs: { src: 'data:image/png;base64,QQ==', alt: 'a', title: 'tt', width: 200, height: 100, alignment: 'right' } }],
  horizontal_rule: [{ type: 'horizontal_rule' }],
  page_break: [{ type: 'page_break' }],
  hard_break: [p([t('line1'), { type: 'hard_break' }, t('line2')])],
  text: [p([t('just text')])],
  inline_image: [p([t('x'), { type: 'inline_image', attrs: { src: 'data:image/png;base64,QQ==', width: 33, height: 22, alt: 'i' } }])],
  bookmark: [p([{ type: 'bookmark', attrs: { name: '_Toc1' } }, t('Heading'), { type: 'bookmark', attrs: { name: '_Ref2' } }])],
  field: [
    p([
      t('see '),
      { type: 'field', attrs: { instruction: 'REF _Ref1 \\h', kind: 'REF', result: 'Figure 1', dirty: false } },
      { type: 'field', attrs: { instruction: 'SEQ Table', kind: 'SEQ', result: [{ type: 'text', text: '2', bold: true }], dirty: true } },
    ]),
  ],
  section_break: [{ type: 'section_break', attrs: { sectionId: 'sec_a' } }],
  // A control around blocks, nested, with every property set; one around a table cell's content.
  content_control: [
    {
      type: 'content_control',
      attrs: { controlType: 'dropDown', tag: 'berth', alias: 'Berth', lock: 'sdtContentLocked', placeholder: 'Placeholder_1', showingPlaceholder: true, items: [{ text: 'North', value: 'N' }], dateFormat: 'dd/MM/yyyy', dateValue: '2026-03-02T00:00:00Z', multiline: true, gallery: 'Cover Pages' },
      content: [p([t('outer')]), { type: 'content_control', attrs: { controlType: 'richText', alias: 'Inner' }, content: [p([t('inner')])] }],
    },
    {
      type: 'table',
      content: [{ type: 'table_row', content: [{ type: 'table_cell', content: [{ type: 'content_control', attrs: { controlType: 'richText', tag: 'cell' }, content: [p([t('in a cell')])] }] }] }],
    },
  ],
  content_control_inline: [
    p([
      t('name '),
      { type: 'content_control_inline', attrs: { controlType: 'text', alias: 'Name', tag: 'name', multiline: true }, content: [t('Ada ', [{ type: 'bold' }]), { type: 'bookmark', attrs: { name: 'b' } }] },
      { type: 'content_control_inline', attrs: { controlType: 'checkbox', checked: true, alias: 'Inspected' }, content: [t('\u2612')] },
      { type: 'content_control_inline', attrs: { controlType: 'date', dateFormat: 'd MMMM yyyy', dateValue: '2026-03-02T00:00:00Z', showingPlaceholder: true }, content: [] },
      { type: 'content_control_inline', attrs: { controlType: 'comboBox', items: [{ text: 'a', value: 'b' }] }, content: [{ type: 'content_control_inline', attrs: { controlType: 'richText' }, content: [t('nested')] }] },
    ]),
  ],
  // Note references hold the text of the note, with formatting, a link, a line break, an image, a bookmark and a field.
  note_ref: [
    p([
      t('claim'),
      {
        type: 'note_ref',
        attrs: { kind: 'footnote', noteId: '7' },
        content: [t('See '), t('page 4', [{ type: 'bold' }]), t(' at '), t('link', [{ type: 'link', attrs: { href: 'https://example.com/n', title: null, target: '_blank' } }]), { type: 'hard_break' }, t('and more')],
      },
      t(' and'),
      { type: 'note_ref', attrs: { kind: 'endnote', noteId: '1', customMark: '*' }, content: [{ type: 'inline_image', attrs: { src: 'data:image/png;base64,QQ==', width: 10, height: 10, alt: 'i' } }, { type: 'bookmark', attrs: { name: 'nb' } }, { type: 'field', attrs: { instruction: 'REF x', kind: 'REF', result: 'x', dirty: false } }] },
      { type: 'note_ref', attrs: { kind: 'footnote', noteId: '8' }, content: [] },
    ]),
  ],
  dynamic_field: [p([{ type: 'dynamic_field', attrs: { fieldType: 'date', format: 'YYYY' } }, t(' / '), { type: 'dynamic_field', attrs: { fieldType: 'pageNumber' } }])],
};

const MARK_SAMPLES: Record<string, Json> = {
  bold: { type: 'bold' },
  italic: { type: 'italic' },
  underline: { type: 'underline' },
  strikethrough: { type: 'strikethrough' },
  superscript: { type: 'superscript' },
  subscript: { type: 'subscript' },
  smallCaps: { type: 'smallCaps' },
  allCaps: { type: 'allCaps' },
  hidden: { type: 'hidden' },
  link: { type: 'link', attrs: { href: 'https://example.com/a', title: null, target: '_blank' } },
  textStyle: { type: 'textStyle', attrs: { color: '#ff0000', backgroundColor: '#00ff00', fontSize: 18, fontFamily: 'Georgia', fontFamilyEastAsia: 'MS Mincho', fontFamilyComplex: 'Arial' } },
  comment: { type: 'comment', attrs: { threadId: 'thread_1' } },
  namedStyle: { type: 'namedStyle', attrs: { id: 'AccentWord' } },
};

const nodeNames = Object.keys(docsSchema.nodes);
const markNames = Object.keys(docsSchema.marks);

describe('schema coverage', () => {
  it('has a sample for every node and mark type', () => {
    expect(Object.keys(NODE_SAMPLES).sort()).toEqual([...nodeNames].sort());
    expect(Object.keys(MARK_SAMPLES).sort()).toEqual([...markNames].sort());
  });
});

function markedDoc(markName: string): Json[] {
  return [p([t('plain '), t('marked', [MARK_SAMPLES[markName]]), t(' tail')])];
}

function docOf(content: Json[]): PmNode {
  return docsSchema.nodeFromJSON({ type: 'doc', content });
}

describe.each(nodeNames.filter((n) => n !== 'doc'))('node type %s: two-client convergence', (name) => {
  it('PM-level: edits on both clients converge and keep the node + attrs', () => {
    const a = new Client('A');
    const b = new Client('B');
    const original = docOf([...NODE_SAMPLES[name], p([t('end')])]);
    a.edit((tr) => tr.replaceWith(0, tr.doc.content.size, original.content));
    sync(a, b);
    expect(b.pm().toJSON()).toEqual(original.toJSON());

    // concurrent text edits at both ends
    a.edit((tr) => tr.insertText('A', tr.doc.content.size - 1));
    b.edit((tr) => tr.insert(0, docsSchema.nodes.paragraph.create(null, docsSchema.text('top'))));
    sync(a, b);
    expect(a.pm().toJSON()).toEqual(b.pm().toJSON());
    // the sample content is still present, unchanged
    const survivors = a.pm().toJSON().content.slice(1, -1);
    expect(survivors).toEqual(original.toJSON().content.slice(0, -1));
  });

  it('model-level: hydrate/serialize (seeding + getData) does not drop the node', () => {
    // Build blocks via the PM doc -> sync converters used by the product, and
    // assert the Y.Doc round-trip keeps every PM node of this type.
    const original = docOf(NODE_SAMPLES[name]);
    const count = (d: PmNode): number => {
      let n = 0;
      d.descendants((x) => {
        if (x.type.name === name) n++;
      });
      return n;
    };
    const [a] = pair([]);
    a.edit((tr) => tr.replaceWith(0, tr.doc.content.size, original.content));
    const viaModel = sectionsToPmDoc(a.data().sections);
    expect(count(viaModel)).toBe(count(original));
    expect(viaModel.toJSON()).toEqual(original.toJSON());
  });
});

describe.each(markNames)('mark type %s: two-client convergence', (name) => {
  it('PM-level: mark and attrs survive concurrent edits', () => {
    const a = new Client('A');
    const b = new Client('B');
    const original = docOf(markedDoc(name));
    a.edit((tr) => tr.replaceWith(0, tr.doc.content.size, original.content));
    sync(a, b);
    a.edit((tr) => tr.insertText('A', 1));
    b.edit((tr) => tr.insertText('B', b.pm().content.size - 1));
    sync(a, b);
    expect(a.pm().toJSON()).toEqual(b.pm().toJSON());
    const marked = JSON.stringify(a.pm().toJSON());
    expect(marked).toContain(`"type":"${name}"`);
  });

  it('model-level: mark survives blocks round-trip', () => {
    const original = docOf(markedDoc(name));
    const [a] = pair([]);
    a.edit((tr) => tr.replaceWith(0, tr.doc.content.size, original.content));
    const viaModel = sectionsToPmDoc(a.data().sections);
    expect(viaModel.toJSON()).toEqual(original.toJSON());
  });
});

describe('header / footer content under collab', () => {
  it('page-number fields and logo images in header/footer survive hydrate -> serialize', () => {
    const data = dataWith([para('p', [text('body')])]);
    data.sections[0].header = {
      blocks: [{ type: 'paragraph', alignment: 'center', content: [{ type: 'image', src: 'data:image/png;base64,QQ==', width: 40, height: 20 }, { type: 'text', text: 'Hdr', bold: true }] }],
      differentFirstPage: true,
      firstPageBlocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'First' }] }],
    };
    data.sections[0].footer = {
      blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'Page ' }, { type: 'dynamicField', fieldType: 'pageNumber' }, { type: 'dynamicField', fieldType: 'totalPages' }] }],
    };
    const a = new Y.Doc();
    hydrateYDocFromData(a, data);
    const b = new Y.Doc();
    Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
    const out = serializeYDocToData(b).sections[0];
    expect(out.header).toEqual(data.sections[0].header);
    expect(out.footer).toEqual(data.sections[0].footer);
  });
});

describe('header/footer variants and section properties under collab', () => {
  it('even-page content, columns, page numbering and vertical alignment survive hydrate -> sync -> serialize', () => {
    const data = dataWith([para('p', [text('body')])]);
    data.sections[0].pageConfig = { ...DEFAULT_PAGE_CONFIG, columns: { count: 3, spacing: 24, separator: true, widths: [100, 200, 300] } };
    data.sections[0].header = {
      blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'odd' }] }],
      differentFirstPage: true,
      differentOddEven: true,
      evenPageBlocks: [{ type: 'paragraph', alignment: 'right', content: [{ type: 'text', text: 'even' }] }],
    };
    data.sections[0].pageNumbering = { start: 10, format: 'lowerRoman' };
    data.sections[0].verticalAlign = 'center';
    const a = new Y.Doc();
    hydrateYDocFromData(a, data);
    const b = new Y.Doc();
    Y.applyUpdate(b, Y.encodeStateAsUpdate(a));
    const out = serializeYDocToData(b).sections[0];
    expect(out.header).toEqual(data.sections[0].header);
    expect(out.pageConfig.columns).toEqual({ count: 3, spacing: 24, separator: true, widths: [100, 200, 300] });
    expect(out.pageNumbering).toEqual({ start: 10, format: 'lowerRoman' });
    expect(out.verticalAlign).toBe('center');
  });

  it('how a section starts survives a Y round trip', () => {
    const data = dataWith([para('p', [text('body')])]);
    data.sections[0].breakType = 'evenPage';
    const a = new Y.Doc();
    hydrateYDocFromData(a, data);
    expect(serializeYDocToData(a).sections[0].breakType).toBe('evenPage');
  });

  it('a section without them reads back without the fields', () => {
    const a = new Y.Doc();
    hydrateYDocFromData(a, dataWith([para('p', [text('body')])]));
    const out = serializeYDocToData(a).sections[0];
    expect(out).not.toHaveProperty('pageNumbering');
    expect(out).not.toHaveProperty('verticalAlign');
    expect(out).not.toHaveProperty('breakType');
  });
});
