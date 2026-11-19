import { DocumentImpl } from './document';
import { createParagraphFromText } from './blocks/paragraph';

describe('DocumentImpl — construction', () => {
  it('starts with id, title, and one section holding one paragraph', () => {
    const d = new DocumentImpl('doc1', 'My Doc');
    expect(d.getId()).toBe('doc1');
    expect(d.getTitle()).toBe('My Doc');
    expect(d.getSections()).toHaveLength(1);
    expect(d.getSections()[0].blocks).toHaveLength(1);
  });

  it('defaults the title when none is given', () => {
    expect(new DocumentImpl().getTitle()).toBe('Untitled Document');
  });
});

describe('DocumentImpl — sections and blocks', () => {
  it('adds, inserts, and deletes sections', () => {
    const d = new DocumentImpl();
    const s2 = d.addSection();
    expect(d.getSections()).toHaveLength(2);
    const s0 = d.insertSection(0);
    expect(d.getSections()[0].id).toBe(s0.id);
    expect(d.deleteSection(s2.id)).toBe(true);
    expect(d.getSections()).toHaveLength(2);
  });

  it('refuses to delete the last remaining section', () => {
    const d = new DocumentImpl();
    const only = d.getSections()[0].id;
    expect(d.deleteSection(only)).toBe(false);
    expect(d.getSections()).toHaveLength(1);
  });

  it('sets, adds, inserts, updates, and deletes blocks', () => {
    const d = new DocumentImpl();
    const sectionId = d.getSections()[0].id;
    const a = createParagraphFromText('A');
    const b = createParagraphFromText('B');
    d.setSectionBlocks(sectionId, [a, b]);
    expect(d.getSection(sectionId)!.blocks.map((x) => x.id)).toEqual([a.id, b.id]);

    const c = createParagraphFromText('C');
    d.insertBlock(sectionId, 1, c);
    expect(d.getSection(sectionId)!.blocks.map((x) => x.id)).toEqual([a.id, c.id, b.id]);

    const dd = createParagraphFromText('D');
    d.addBlock(sectionId, dd);
    expect(d.getSection(sectionId)!.blocks).toHaveLength(4);

    expect(d.deleteBlock(c.id)).toBe(true);
    expect(d.getBlock(c.id)).toBeUndefined();
    expect(d.getSection(sectionId)!.blocks.map((x) => x.id)).toEqual([a.id, b.id, dd.id]);
  });

  it('locates a block by id across sections', () => {
    const d = new DocumentImpl();
    const s1 = d.getSections()[0].id;
    const s2 = d.addSection().id;
    const target = createParagraphFromText('here');
    d.setSectionBlocks(s2, [createParagraphFromText('x'), target]);
    expect(d.getBlockIndex(target.id)).toEqual({ sectionIndex: 1, blockIndex: 1 });
    expect(d.getBlockSection(target.id)!.id).toBe(s2);
    expect(d.getBlockIndex('nope')).toBeNull();
    expect(s1).not.toBe(s2);
  });
});

describe('DocumentImpl — getData / setData round-trip', () => {
  it('round-trips the full document through serialization', () => {
    const d = new DocumentImpl('doc-x', 'Round Trip');
    const sectionId = d.getSections()[0].id;
    d.setSectionBlocks(sectionId, [createParagraphFromText('one'), createParagraphFromText('two')]);
    const data = d.getData();

    const d2 = new DocumentImpl();
    d2.setData(data);
    expect(d2.getId()).toBe('doc-x');
    expect(d2.getTitle()).toBe('Round Trip');
    expect(d2.getData()).toEqual(data);
  });

  it('keeps the page numbering, vertical alignment, columns and header/footer variants of a section', () => {
    const d = new DocumentImpl('doc-y', 'Sections');
    const data = d.getData();
    data.sections[0] = {
      ...data.sections[0],
      pageConfig: { ...data.sections[0].pageConfig, columns: { count: 2, spacing: 24 } },
      header: { blocks: [], differentOddEven: true, evenPageBlocks: [] },
      pageNumbering: { start: 3, format: 'lowerRoman' },
      verticalAlign: 'bottom',
    };
    const d2 = new DocumentImpl();
    d2.setData(data);
    expect(d2.getData()).toEqual(data);
    expect(d2.getData().sections[0]).toMatchObject({ pageNumbering: { start: 3, format: 'lowerRoman' }, verticalAlign: 'bottom', header: { differentOddEven: true, evenPageBlocks: [] } });
    expect(d2.getData().sections[0].pageConfig.columns).toEqual({ count: 2, spacing: 24 });
  });
});
