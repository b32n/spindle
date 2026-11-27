import { parseTree, TreeBuilder } from './tree';

describe('parseTree', () => {
  it('builds the tree with local names, attributes, text (entities decoded) and children', () => {
    const { root, overLimit } = parseTree('<?xml version="1.0"?><x:a xmlns:x="urn:x" k="1&lt;2"><b>t&amp;u</b><c/>tail</x:a>', 100);
    expect(overLimit).toBe(false);
    expect(root).toMatchObject({ name: 'a', ns: '', attrs: { k: '1<2' }, text: 'tail' });
    expect(root?.children.map((c) => [c.name, c.text, c.children.length])).toEqual([['b', 't&u', 0], ['c', '', 0]]);
  });

  it('keeps __proto__ and constructor as attribute names like any other', () => {
    const root = parseTree('<a __proto__="1" constructor="2"/>', 10).root;
    expect(root?.attrs.__proto__).toBe('1');
    expect(root?.attrs.constructor).toBe('2');
    expect(({} as Record<string, unknown>).constructor).toBe(Object);
  });

  it('is not over at exactly the element budget, and over with one more', () => {
    expect(parseTree('<a><b/><b/></a>', 3)).toMatchObject({ overLimit: false, root: { name: 'a' } });
    expect(parseTree('<a><b/><b/></a>', 2)).toEqual({ root: undefined, overLimit: true });
  });

  it('stops scanning at the first element over the budget (a hostile part costs the budget, not its size)', () => {
    const t0 = Date.now();
    const out = parseTree(`<a>${'<b><c/></b>'.repeat(3_000_000)}</a>`, 1_000);
    expect(out).toEqual({ root: undefined, overLimit: true });
    expect(Date.now() - t0).toBeLessThan(5_000);
  });

  it('gives no root for text that is not XML, an empty part, damage and nesting past the depth cap', () => {
    for (const text of ['', 'just text', '<?xml version="1.0"?>', '<a><b></a>', '<a>', `${'<a>'.repeat(300)}${'</a>'.repeat(300)}`]) expect(parseTree(text, 1_000)).toEqual({ root: undefined, overLimit: false });
  });

  it('keeps only the first of two root elements', () => {
    const { root } = parseTree('<a><x/></a><b><y/></b>', 100);
    expect(root?.name).toBe('a');
    expect(root?.children.map((c) => c.name)).toEqual(['x']);
  });
});

describe('TreeBuilder', () => {
  it('counts the elements it builds against the shared budget and reports a dropped one', () => {
    const budget = { left: 2 };
    const b = new TreeBuilder(budget);
    b.start('a', Object.create(null));
    b.start('b', Object.create(null));
    expect(b.dropped).toBe(false);
    b.start('c', Object.create(null));
    b.text('lost');
    b.end();
    expect(b.dropped).toBe(true);
    expect(b.open).toBe(true);
    b.text('kept');
    b.end();
    b.end();
    expect(b.open).toBe(false);
    expect(budget.left).toBe(0);
    expect(b.root?.children[0]).toMatchObject({ name: 'b', text: 'kept', children: [] });
  });

  it('ignores text outside any element', () => {
    const b = new TreeBuilder({ left: 1 });
    b.text('x');
    expect(b.root).toBeUndefined();
    expect(b.open).toBe(false);
  });
});
