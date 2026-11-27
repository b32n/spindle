import { groupChain, groupParents, rootGroupOf } from './group-tree';

describe('group tree', () => {
  it('reads each group parent from its members, smallest member id winning a disagreement', () => {
    const parents = groupParents([
      { id: 'a', groupId: 'inner', parentGroupId: 'outer' },
      { id: 'b', groupId: 'inner', parentGroupId: 'outer' },
      { id: 'c', groupId: 'inner', parentGroupId: 'other' },
      { id: 'd', groupId: 'outer' },
      { id: 'e' },
      { id: 'f', parentGroupId: 'orphan' },
      { id: 'g', groupId: 'self', parentGroupId: 'self' },
    ]);
    expect([...parents]).toEqual([['inner', 'outer']]);
    // Order independence: the same members in another order give the same tree.
    expect(groupParents([{ id: 'c', groupId: 'inner', parentGroupId: 'other' }, { id: 'a', groupId: 'inner', parentGroupId: 'outer' }]).get('inner')).toBe('outer');
    expect(groupParents([{ id: 'a', groupId: 'inner', parentGroupId: 'outer' }, { id: 'c', groupId: 'inner', parentGroupId: 'other' }]).get('inner')).toBe('outer');
  });

  it('walks the chain innermost first and cuts cycles', () => {
    const parents = new Map([['a', 'b'], ['b', 'c']]);
    expect(groupChain('a', parents)).toEqual(['a', 'b', 'c']);
    expect(groupChain('c', parents)).toEqual(['c']);
    expect(groupChain(undefined, parents)).toEqual([]);
    expect(groupChain('x', new Map([['x', 'y'], ['y', 'x']]))).toEqual(['x', 'y']);
  });

  it('finds the outermost group of an element', () => {
    const parents = new Map([['a', 'b']]);
    expect(rootGroupOf({ id: 'e', groupId: 'a' }, parents)).toBe('b');
    expect(rootGroupOf({ id: 'e', groupId: 'b' }, parents)).toBe('b');
    expect(rootGroupOf({ id: 'e' }, parents)).toBeUndefined();
  });
});
