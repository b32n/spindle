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
