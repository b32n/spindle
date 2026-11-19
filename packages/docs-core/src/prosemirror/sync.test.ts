import { blocksToPmDoc, proseMirrorToDocument } from './sync';
import type { Block, Document, ListItemBlock, TableBlock, PageConfig } from '../types';

const emptyPageConfig = {} as PageConfig;

/** blocks -> PM doc -> blocks (the collab hydrate/serialize round trip). */
function roundTrip(blocks: Block[]): Block[] {
  const pm = blocksToPmDoc(blocks);
  const existing: Document = {
    id: 'd',
    title: 'D',
    defaultPageConfig: emptyPageConfig,
    sections: [{ id: 's', pageConfig: emptyPageConfig, blocks: [] }],
  };
  return proseMirrorToDocument(pm, existing).sections[0].blocks;
}

describe('collab round-trip preserves list nesting level', () => {
  it('keeps distinct list-item levels', () => {
    const li = (text: string, level: number): ListItemBlock => ({
      id: `li_${text}`,
      type: 'list-item',
      listType: 'bullet',
      level,
      content: [{ type: 'text', text }],
    });
    const out = roundTrip([li('a', 0), li('b', 1), li('c', 2)]) as ListItemBlock[];
    const items = out.filter((b) => b.type === 'list-item') as ListItemBlock[];
    expect(items.map((b) => b.level)).toEqual([0, 1, 2]);
  });
});

describe('collab round-trip preserves table metadata', () => {
  it('keeps colWidths, row height, and cell/table styleId', () => {
    const table: TableBlock = {
      id: 't',
      type: 'table',
      colWidths: [120, 80],
      styleId: 'style_7',
      rows: [
        {
          id: 'r0',
          height: 44,
          cells: [
            { id: 'c0', content: [{ type: 'text', text: 'A' }], styleId: 'style_3' },
            { id: 'c1', content: [{ type: 'text', text: 'B' }] },
          ],
        },
      ],
    };
    const out = roundTrip([table]);
    const t = out.find((b) => b.type === 'table') as TableBlock;
    expect(t.colWidths).toEqual([120, 80]);
    expect(t.styleId).toBe('style_7');
    expect(t.rows[0].height).toBe(44);
    expect(t.rows[0].cells[0].styleId).toBe('style_3');
  });
});

describe('collab round-trip preserves inline content', () => {
  it('keeps inline images, dynamic fields, line breaks, and formatted links', () => {
    const para: Block = {
      id: 'p',
      type: 'paragraph',
      content: [
        { type: 'text', text: 'a\nb' },
        { type: 'image', src: 'data:image/png;base64,QQ==', width: 10, height: 20, alt: 'x' },
        { type: 'dynamicField', fieldType: 'date', format: 'YYYY' },
        { type: 'link', text: 'l', href: 'https://example.com', bold: true },
      ],
    };
    const out = roundTrip([para]);
    expect(out[0]).toMatchObject({ content: para.content });
  });

  it('keeps bookmarks and internal links', () => {
    const para: Block = {
      id: 'p',
      type: 'paragraph',
      content: [
        { type: 'bookmark', name: '_Toc1' },
        { type: 'text', text: 'see ' },
        { type: 'link', text: 'there', href: '#_Toc1' },
      ],
    };
    expect(roundTrip([para])[0]).toMatchObject({ content: para.content });
  });

  it('keeps generic fields with their instruction, cached result and dirty flag', () => {
    const para: Block = {
      id: 'p',
      type: 'paragraph',
      content: [
        { type: 'text', text: 'see ' },
        { type: 'field', instruction: 'REF _Ref1 \\h', kind: 'REF', result: 'Figure 1' },
        { type: 'field', instruction: 'SEQ Table', kind: 'SEQ', result: [{ type: 'text', text: '2', bold: true }], dirty: true },
      ],
    };
    expect(roundTrip([para])[0]).toMatchObject({ content: para.content });
  });

  it('keeps tab characters inside a run, with or without formatting, and in a plain run with line breaks', () => {
    const para: Block = {
      id: 'p',
      type: 'paragraph',
      content: [
        { type: 'text', text: 'Chapter one\t1', bold: true },
        { type: 'text', text: ' plain\ttab\nnext\tline' },
      ],
    };
    expect(roundTrip([para])[0]).toMatchObject({ content: para.content });
  });

  it('resolves styleId runs from the pool when no pool is passed explicitly', () => {
    const pm = blocksToPmDoc([
      { id: 'p', type: 'paragraph', content: [{ type: 'text', text: 'hi', styleId: 's1' }] },
    ], {
      styles: new Map(),
      getOrCreate: () => 's1',
      get: (id: string) => (id === 's1' ? { bold: true } : undefined),
      getAllStyles: () => new Map(),
    });
    expect(pm.firstChild!.firstChild!.marks.map((m) => m.type.name)).toEqual(['bold']);
  });
});

describe('collab round-trip preserves DOCX paragraph formatting and run flags', () => {
  const text = (t: string, extra: object = {}) => ({ type: 'text' as const, text: t, ...extra });
  const fmt = {
    lineSpacing: { type: 'multiple' as const, value: 1.15 },
    leftIndent: 36,
    rightIndent: 12,
    firstLineIndent: -18,
    keepWithNext: true,
    keepTogether: true,
    pageBreakBefore: true,
  };

  it('keeps indents, line spacing, flags and an explicit zero spacing on paragraphs, headings and list items', () => {
    const blocks: Block[] = [
      { id: 'p', type: 'paragraph', content: [text('p')], spaceAfter: 0, spaceBefore: 6, ...fmt },
      { id: 'h', type: 'heading', level: 2, content: [text('h')], spaceAfter: 0, spaceBefore: 12, ...fmt },
      { id: 'l', type: 'list-item', listType: 'bullet', level: 0, content: [text('l')], spaceAfter: 0, ...fmt },
    ];
    const out = roundTrip(blocks);
    expect(out[0]).toMatchObject({ type: 'paragraph', spaceAfter: 0, spaceBefore: 6, ...fmt });
    expect(out[1]).toMatchObject({ type: 'heading', level: 2, spaceAfter: 0, spaceBefore: 12, ...fmt });
    expect(out[2]).toMatchObject({ type: 'list-item', spaceAfter: 0, ...fmt });
  });

  it('leaves unset formatting unset (defaults are not written back as values)', () => {
    const out = roundTrip([
      { id: 'p', type: 'paragraph', content: [text('p')] },
      { id: 'h', type: 'heading', level: 1, content: [text('h')] },
    ]);
    for (const b of out) {
      expect(b).not.toHaveProperty('spaceAfter');
      expect(b).not.toHaveProperty('spaceBefore');
      expect(b).not.toHaveProperty('leftIndent');
      expect(b).not.toHaveProperty('lineSpacing');
      expect(b).not.toHaveProperty('keepWithNext');
    }
  });

  it('does not store single spacing (it is the default)', () => {
    const out = roundTrip([{ id: 'p', type: 'paragraph', content: [text('p')], lineSpacing: { type: 'single' } }]);
    expect(out[0]).not.toHaveProperty('lineSpacing');
  });

  it('keeps all caps, hidden text and per-script fonts on text and on links', () => {
    const out = roundTrip([
      {
        id: 'p',
        type: 'paragraph',
        content: [
          text('a', { allCaps: true, hidden: true, fontFamily: 'Georgia', fontFamilyEastAsia: 'MS Mincho', fontFamilyComplex: 'Arial' }),
          { type: 'link', text: 'l', href: 'https://example.com/', allCaps: true, hidden: true },
        ],
      },
    ]);
    const content = (out[0] as { content: object[] }).content;
    expect(content[0]).toMatchObject({ allCaps: true, hidden: true, fontFamily: 'Georgia', fontFamilyEastAsia: 'MS Mincho', fontFamilyComplex: 'Arial' });
    expect(content[1]).toMatchObject({ type: 'link', allCaps: true, hidden: true });
  });

  describe('numbered lists', () => {
    const li = (id: string, start?: number): ListItemBlock => ({
      id,
      type: 'list-item',
      listType: 'numbered',
      level: 0,
      content: [text(id)],
      ...(start !== undefined && { start }),
    });
    const starts = (blocks: Block[]) => blocks.map((b) => (b.type === 'list-item' ? b.start : 'x'));

    it('keeps a first number other than 1', () => {
      expect(starts(roundTrip([li('a', 7), li('b')]))).toEqual([7, undefined]);
    });

    it('keeps a restart at 1 right after another numbered list as a separate list', () => {
      const out = roundTrip([li('a'), li('b'), li('c', 1), li('d')]);
      expect(starts(out)).toEqual([undefined, undefined, 1, undefined]);
      expect(blocksToPmDoc([li('a'), li('b'), li('c', 1), li('d')]).childCount).toBe(2);
    });

    it('does not turn a start of 1 on the very first item, or a start on a bullet, into a restart', () => {
      expect(starts(roundTrip([li('a', 1), li('b')]))).toEqual([undefined, undefined]);
      expect(blocksToPmDoc([li('a', 1), li('b')]).childCount).toBe(1);
    });
  });
});

describe('collab round-trip preserves named style references', () => {
  const text = (t: string, extra: object = {}) => ({ type: 'text' as const, text: t, ...extra });

  it('keeps namedStyleId on paragraphs, headings and list items', () => {
    const out = roundTrip([
      { id: 'p', type: 'paragraph', content: [text('p')], namedStyleId: 'Callout' },
      { id: 'h', type: 'heading', level: 2, content: [text('h')], namedStyleId: 'Heading2' },
      { id: 'l', type: 'list-item', listType: 'numbered', level: 0, content: [text('l')], namedStyleId: 'ListParagraph' },
    ]);
    expect(out[0]).toMatchObject({ type: 'paragraph', namedStyleId: 'Callout' });
    expect(out[1]).toMatchObject({ type: 'heading', namedStyleId: 'Heading2' });
    expect(out[2]).toMatchObject({ type: 'list-item', namedStyleId: 'ListParagraph' });
  });

  it('keeps namedStyleId on text runs and links next to their formatting, and leaves unstyled runs unstyled', () => {
    const out = roundTrip([
      {
        id: 'p',
        type: 'paragraph',
        content: [
          text('plain '),
          text('accent', { namedStyleId: 'AccentWord', bold: true }),
          { type: 'link' as const, text: 'site', href: 'https://example.com/', namedStyleId: 'Hyperlink', underline: true },
        ],
      },
    ]);
    const content = (out[0] as { content: object[] }).content;
    expect(content[0]).not.toHaveProperty('namedStyleId');
    expect(content[1]).toMatchObject({ text: 'accent', namedStyleId: 'AccentWord' });
    expect(content[2]).toMatchObject({ type: 'link', href: 'https://example.com/', namedStyleId: 'Hyperlink' });
  });

  it('does not invent a namedStyleId', () => {
    const out = roundTrip([{ id: 'p', type: 'paragraph', content: [text('p')] }]);
    expect(out[0]).not.toHaveProperty('namedStyleId');
  });
});
