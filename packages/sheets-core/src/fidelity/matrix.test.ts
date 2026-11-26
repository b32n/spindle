import {
  buildMatrix,
  injectMatrix,
  uncoveredConstructs,
} from '../../../../scripts/xlsx-fidelity/matrix';

const registry = [
  { id: 'a.one', label: 'One' },
  { id: 'a.two', label: 'Two' },
];
const start = '<!-- start -->';
const end = '<!-- end -->';

describe('matrix', () => {
  it('builds a file list and a construct x file table', () => {
    const md = buildMatrix(registry, [
      { id: 'f1', constructs: ['a.one'], title: 'First', source: { tool: 'tool' } },
      { id: 'f2', constructs: ['a.one', 'a.two'] },
    ]);
    expect(md.split('\n')).toEqual([
      '| # | Fixture | Title | Produced by |',
      '|--:|---|---|---|',
      '| 1 | `f1` | First | tool |',
      '| 2 | `f2` |  |  |',
      '',
      '| Construct | 1 | 2 |',
      '|---|:-:|:-:|',
      '| One (`a.one`) | yes | yes |',
      '| Two (`a.two`) |  | yes |',
    ]);
  });

  it('lists constructs no fixture claims', () => {
    expect(uncoveredConstructs(registry, [{ constructs: ['a.one'] }])).toEqual(['a.two']);
    expect(
      uncoveredConstructs(registry, [{ constructs: ['a.one'] }, { constructs: ['a.two'] }])
    ).toEqual([]);
    expect(uncoveredConstructs(registry, [])).toEqual(['a.one', 'a.two']);
  });

  it('replaces the generated block and leaves the rest alone', () => {
    const readme = `intro\n${start}\n\nold\n\n${end}\noutro\n`;
    const next = injectMatrix(readme, 'NEW', start, end);
    expect(next).toBe(`intro\n${start}\n\nNEW\n\n${end}\noutro\n`);
    expect(injectMatrix(next, 'NEW', start, end)).toBe(next);
  });

  it('refuses a README without the markers, or with them the wrong way round', () => {
    expect(() => injectMatrix('nothing', 'x', start, end)).toThrow(/markers/);
    expect(() => injectMatrix(`${end}\n${start}`, 'x', start, end)).toThrow(/markers/);
    expect(() => injectMatrix(start, 'x', start, end)).toThrow(/markers/);
  });
});
