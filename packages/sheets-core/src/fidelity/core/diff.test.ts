import {
  MAX_DIFF_ENTRIES,
  alignBySignature,
  diffValues,
  isObj,
} from '../../../../../scripts/fidelity-core/diff';
import type { Json } from '../../../../../scripts/fidelity-core/types';

const counts = (m: Json) => ({ keys: isObj(m) ? Object.keys(m).length : 0 });

describe('isObj', () => {
  it('is true only for plain objects', () => {
    expect([isObj({}), isObj([]), isObj(null), isObj(1), isObj(undefined)]).toEqual([
      true,
      false,
      false,
      false,
      false,
    ]);
  });
});

describe('alignBySignature', () => {
  it('matches the longest common subsequence', () => {
    expect(alignBySignature(['a', 'b', 'c', 'd'], ['a', 'c', 'x', 'd'])).toEqual([
      [0, 0],
      [2, 1],
      [3, 3],
    ]);
  });
  it('handles empty sides and no overlap, preferring to skip the earlier item on ties', () => {
    expect(alignBySignature([], ['a'])).toEqual([]);
    expect(alignBySignature(['a'], [])).toEqual([]);
    expect(alignBySignature(['a', 'b'], ['b', 'a'])).toEqual([[1, 0]]);
    expect(alignBySignature(['x'], ['y'])).toEqual([]);
  });
});

describe('diffValues', () => {
  it('finds nothing in equal values', () => {
    const d = diffValues({ a: [1, { b: 2 }] }, { a: [1, { b: 2 }] }, { counts });
    expect(d).toMatchObject({
      total: 0,
      added: 0,
      removed: 0,
      changed: 0,
      entries: [],
      truncated: false,
      byField: {},
    });
    expect(d.before).toEqual({ keys: 1 });
  });

  it('records added, removed and changed values with paths in key order', () => {
    const d = diffValues({ a: 1, b: 2, c: [1, 2] }, { a: 9, c: [1, 2, 3], d: 'new' }, { counts });
    expect(d.total).toBe(4);
    expect(d.entries).toEqual([
      { path: 'a', kind: 'changed', before: '1', after: '9' },
      { path: 'b', kind: 'removed', before: '2' },
      { path: 'c[2]', kind: 'added', after: '3' },
      { path: 'd', kind: 'added', after: '"new"' },
    ]);
    expect(d.byField).toEqual({ 'added c[]': 1, 'added d': 1, 'changed a': 1, 'removed b': 1 });
  });

  it('treats arrays as changed values when a side is not an array and compares type changes', () => {
    const d = diffValues(
      { a: [1], b: { x: 1 }, c: null },
      { a: { x: 1 }, b: [1], c: 0 },
      { counts }
    );
    expect(d.changed).toBe(3);
  });

  it('previews long values with an ellipsis and survives values JSON cannot print', () => {
    const d = diffValues({ s: 'x'.repeat(200) }, { s: 'y' }, { counts });
    expect(d.entries[0].before).toHaveLength(80);
    expect(d.entries[0].before?.endsWith('...')).toBe(true);
    const e = diffValues({}, { u: undefined } as unknown as Json, { counts });
    expect(e.total).toBe(0);
    const f = diffValues({ a: 1 }, { a: undefined } as unknown as Json, { counts });
    expect(f.entries[0]).toMatchObject({ kind: 'removed', path: 'a' });
    const g = diffValues({ a: undefined } as unknown as Json, { a: 1 }, { counts });
    expect(g.entries[0]).toMatchObject({ kind: 'added', path: 'a' });
    const h = diffValues({ a: 1 }, { a: undefined } as unknown as Json, { counts, cap: 0 });
    expect(h.truncated).toBe(true);
  });

  it('caps the entries but counts every difference', () => {
    const before: Json = { list: Array.from({ length: 100 }, (_, i) => i) };
    const after: Json = { list: Array.from({ length: 100 }, (_, i) => i + 1000) };
    const d = diffValues(before, after, { counts });
    expect(d.total).toBe(100);
    expect(d.entries).toHaveLength(MAX_DIFF_ENTRIES);
    expect(d.truncated).toBe(true);
    expect(d.byField).toEqual({ 'changed list[]': 100 });
    expect(diffValues(before, after, { counts, cap: 3 }).entries).toHaveLength(3);
  });

  it('applies prepare to both sides but counts the originals', () => {
    const d = diffValues(
      { a: 1, ignored: 1 },
      { a: 1, ignored: 2 },
      {
        counts,
        prepare: (m) => (isObj(m) ? { a: m.a } : m),
      }
    );
    expect(d.total).toBe(0);
    expect(d.before).toEqual({ keys: 2 });
  });

  it('folds fields with the adapter function', () => {
    const d = diffValues(
      { cells: { A1: 1, B2: 1 } },
      { cells: { A1: 2, B2: 2 } },
      { counts, field: (p) => p.replace(/\.[A-Z]\d/, '.*') }
    );
    expect(d.byField).toEqual({ 'changed cells.*': 2 });
  });

  it('aligns arrays by signature so one removed item is one difference', () => {
    const before: Json = { blocks: ['a', 'b', 'c', 'd'].map((t) => ({ t })) };
    const after: Json = { blocks: ['a', 'c', 'd'].map((t) => ({ t })) };
    const options = {
      counts,
      signatures: (path: string, b: Json[], a: Json[]) =>
        path === 'blocks'
          ? {
              before: b.map((x) => String(isObj(x) ? x.t : '')),
              after: a.map((x) => String(isObj(x) ? x.t : '')),
            }
          : null,
    };
    const aligned = diffValues(before, after, options);
    expect(aligned.entries).toEqual([{ path: 'blocks[1]', kind: 'removed', before: '{"t":"b"}' }]);
    expect(diffValues(before, after, { counts }).total).toBe(3);
  });

  it('places an added item after the matched item before it, in document order', () => {
    const options = {
      counts,
      signatures: (_p: string, b: Json[], a: Json[]) => ({
        before: b.map(String),
        after: a.map(String),
      }),
    };
    const d = diffValues([1, 2, 3] as Json, [0, 1, 9, 2, 3, 4] as Json, options);
    expect(d.entries.map((e) => `${e.kind} ${e.path}`)).toEqual([
      'added [0]',
      'added [2]',
      'added [5]',
    ]);
    const changed = diffValues([{ k: 'a', v: 1 }] as Json, [{ k: 'a', v: 2 }] as Json, {
      counts,
      signatures: (_p, b, a) => ({
        before: b.map((x) => String(isObj(x) && x.k)),
        after: a.map((x) => String(isObj(x) && x.k)),
      }),
    });
    expect(changed.entries[0]).toMatchObject({ path: '[0].v', kind: 'changed' });
  });
});
