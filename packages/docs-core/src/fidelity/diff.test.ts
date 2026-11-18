import { MAX_DIFF_ENTRIES, diffModels } from '../../../../scripts/docx-fidelity/diff';
import type { Json } from '../../../../scripts/docx-fidelity/types';

const para = (text: string, extra: Record<string, Json> = {}): Json => ({ type: 'paragraph', content: [{ type: 'text', text }], ...extra });
const doc = (blocks: Json[], pools: Record<string, Json> = {}): Json => ({ sections: [{ blocks }], ...pools });

describe('diffModels', () => {
  it('reports no differences for equal models', () => {
    const d = diffModels(doc([para('a'), para('b')]), doc([para('a'), para('b')]));
    expect(d).toMatchObject({ total: 0, added: 0, removed: 0, changed: 0, entries: [], truncated: false, byField: {} });
    expect(d.before).toEqual({ sections: 1, blocks: { paragraph: 2 }, textChars: 2 });
  });

  it('reports a changed field with its path', () => {
    const d = diffModels(doc([para('a', { alignment: 'left' })]), doc([para('a', { alignment: 'center' })]));
    expect(d.entries).toEqual([{ path: 'sections[0].blocks[0].alignment', kind: 'changed', before: '"left"', after: '"center"' }]);
    expect(d.byField).toEqual({ 'changed sections[].blocks[].alignment': 1 });
  });

  it('reports added and removed fields', () => {
    const d = diffModels(doc([para('a', { alignment: 'left' })]), doc([para('a', { indent: 5 })]));
    expect(d.removed).toBe(1);
    expect(d.added).toBe(1);
    expect(d.entries.map((e) => `${e.kind}:${e.path}`)).toEqual(['removed:sections[0].blocks[0].alignment', 'added:sections[0].blocks[0].indent']);
    expect(d.entries[0]).toEqual({ path: 'sections[0].blocks[0].alignment', kind: 'removed', before: '"left"' });
    expect(d.entries[1]).toEqual({ path: 'sections[0].blocks[0].indent', kind: 'added', after: '5' });
  });

  it('aligns blocks by content so a dropped paragraph is one entry, not a shift cascade', () => {
    const d = diffModels(doc([para('one'), para('two'), para('three'), para('four')]), doc([para('one'), para('three'), para('four')]));
    expect(d.total).toBe(1);
    expect(d.entries[0]).toMatchObject({ kind: 'removed', path: 'sections[0].blocks[1]' });
    expect(d.before.blocks.paragraph).toBe(4);
    expect(d.after.blocks.paragraph).toBe(3);
  });

  it('places an added block among changes in order', () => {
    const d = diffModels(doc([para('a'), para('c', { alignment: 'left' })]), doc([para('a'), para('b'), para('c', { alignment: 'right' })]));
    expect(d.entries.map((e) => `${e.kind}:${e.path}`)).toEqual(['added:sections[0].blocks[1]', 'changed:sections[0].blocks[1].alignment']);
    // added block first (after matched block 0), then the change in matched block 1
  });

  it('adds a block before the first matched block', () => {
    const d = diffModels(doc([para('b')]), doc([para('a'), para('b')]));
    expect(d.entries).toEqual([{ path: 'sections[0].blocks[0]', kind: 'added', after: expect.stringContaining('"a"') }]);
  });

  it('compares table text and non-text inline content when aligning', () => {
    const table = (cell: string): Json => ({ type: 'table', rows: [{ cells: [{ content: [{ type: 'text', text: cell }] }, { content: [{ type: 'image' }] }] }, { notCells: true }] });
    const d = diffModels(doc([table('x')]), doc([table('y')]));
    // different cell text: block signatures differ, so it is one removal and one addition
    expect(d.removed).toBe(1);
    expect(d.added).toBe(1);
    expect(d.before.textChars).toBeGreaterThan(0);
    const same = diffModels(doc([table('x')]), doc([table('x')]));
    expect(same.total).toBe(0);
  });

  it('ignores pool key naming and compares resolved styles', () => {
    const a = doc([para('a', { styleId: 'text-style-1' })], { paragraphStylePool: { 'text-style-1': { alignment: 'center' } } });
    const b = doc([para('a', { styleId: 'paragraph-style-9' })], { paragraphStylePool: { 'paragraph-style-9': { alignment: 'center' } } });
    expect(diffModels(a, b).total).toBe(0);
    const c = doc([para('a', { styleId: 'paragraph-style-9' })], { paragraphStylePool: { 'paragraph-style-9': { alignment: 'right' } } });
    expect(diffModels(a, c).entries[0].path).toBe('sections[0].blocks[0].style.alignment');
  });

  it('caps listed entries but keeps exact totals', () => {
    const many = Array.from({ length: MAX_DIFF_ENTRIES + 10 }, (_, i) => para(`p${i}`, { indent: 1 }));
    const changed = Array.from({ length: MAX_DIFF_ENTRIES + 10 }, (_, i) => para(`p${i}`, { indent: 2 }));
    const d = diffModels(doc(many), doc(changed));
    expect(d.total).toBe(MAX_DIFF_ENTRIES + 10);
    expect(d.entries).toHaveLength(MAX_DIFF_ENTRIES);
    expect(d.truncated).toBe(true);
    expect(d.byField['changed sections[].blocks[].indent']).toBe(MAX_DIFF_ENTRIES + 10);
  });

  it('handles scalars, arrays of unequal length and missing sections', () => {
    expect(diffModels(1, 2).entries).toEqual([{ path: '', kind: 'changed', before: '1', after: '2' }]);
    expect(diffModels({ list: [1, 2] }, { list: [1] }).entries[0]).toMatchObject({ kind: 'removed', path: 'list[1]' });
    expect(diffModels({ list: [1] }, { list: [1, 2] }).entries[0]).toMatchObject({ kind: 'added', path: 'list[1]' });
    expect(diffModels({}, {}).before).toEqual({ sections: 0, blocks: {}, textChars: 0 });
    expect(diffModels({ sections: [null, {}, { blocks: [null] }] }, {}).before.blocks).toEqual({ unknown: 1 });
    expect(diffModels(null, null).total).toBe(0);
  });

  it('truncates long previews', () => {
    const d = diffModels({ a: 'x'.repeat(200) }, { a: 'y' });
    expect(d.entries[0].before!.length).toBe(80);
    expect(d.entries[0].before!.endsWith('...')).toBe(true);
  });
});
