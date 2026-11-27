import {
  countMatches,
  createLossReport,
  describeImportLoss,
  groupImportLoss,
  hasImportLoss,
  type ImportLossItem,
} from './loss';

const item = (id: string, state?: ImportLossItem['state'], count = 1): ImportLossItem => ({
  id,
  label: id,
  count,
  effect: 'e',
  ...(state && { state }),
});

describe('loss report helpers', () => {
  it('drops zero-count items and describes the rest', () => {
    const r = createLossReport('xlsx', [
      item('a'),
      item('b', 'kept', 0),
      { ...item('c'), where: ['S1', 'S2'] },
    ]);
    expect(r.items.map((i) => i.id)).toEqual(['a', 'c']);
    expect(hasImportLoss(r)).toBe(true);
    expect(hasImportLoss(createLossReport('csv', []))).toBe(false);
    expect(describeImportLoss(r)).toEqual(['1 × a — e', '1 × c (S1, S2) — e']);
  });

  it('keeps `state` out of the item unless an importer sets it', () => {
    const r = createLossReport('xlsx', [item('a')]);
    expect(Object.keys(r.items[0])).not.toContain('state');
    expect(JSON.stringify(r)).not.toContain('state');
  });

  it('groups by state, absent meaning lost, preserving order', () => {
    const g = groupImportLoss({
      format: 'pptx',
      items: [
        item('a'),
        item('b', 'kept'),
        item('c', 'lost'),
        item('d', 'approximated'),
        item('e', 'kept'),
      ],
    });
    expect(g.lost.map((i) => i.id)).toEqual(['a', 'c']);
    expect(g.approximated.map((i) => i.id)).toEqual(['d']);
    expect(g.kept.map((i) => i.id)).toEqual(['b', 'e']);
  });

  it('counts regex matches', () => {
    expect(countMatches('a1b22', /\d/g)).toBe(3);
    expect(countMatches('abc', /\d/g)).toBe(0);
  });
});
