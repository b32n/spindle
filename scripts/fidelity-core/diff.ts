import type { DiffEntry, DiffReport, GroupedCounts, Json } from './types';

type Obj = { [key: string]: Json };
export const isObj = (v: Json | undefined): v is Obj =>
  typeof v === 'object' && v !== null && !Array.isArray(v);

export const MAX_DIFF_ENTRIES = 40;

const preview = (v: Json | undefined): string => {
  const s = String(JSON.stringify(v));
  return s.length > 80 ? `${s.slice(0, 77)}...` : s;
};

/** Longest-common-subsequence alignment of two signature lists; returns matched index pairs. */
export function alignBySignature(a: string[], b: string[]): Array<[number, number]> {
  const n = a.length;
  const m = b.length;
  const table: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      table[i][j] =
        a[i] === b[j] ? table[i + 1][j + 1] + 1 : Math.max(table[i + 1][j], table[i][j + 1]);
    }
  }
  const pairs: Array<[number, number]> = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) pairs.push([i++, j++]);
    else if (table[i + 1][j] >= table[i][j + 1]) i++;
    else j++;
  }
  return pairs;
}

export interface DiffOptions {
  /** Size counters of each side, shown in the report next to the differences. */
  counts: (model: Json) => GroupedCounts;
  /** Applied to both sides before comparing (e.g. resolve style pools so pool keys do not matter). */
  prepare?: (model: Json) => Json;
  /**
   * Arrays whose items should be matched by content rather than by index (so one dropped item is one
   * entry, not a cascade of index shifts): return the signature of each side's items, or null to compare by index.
   */
  signatures?: (
    path: string,
    before: Json[],
    after: Json[]
  ) => { before: string[]; after: string[] } | null;
  /** Folds a path into the key used for `byField` (default: array indexes stripped). */
  field?: (path: string) => string;
  /** Entries kept in the report (default 40); totals always count everything. */
  cap?: number;
}

const stripIndexes = (path: string) => path.replace(/\[\d+\]/g, '[]');

/**
 * Structural diff of two normalised models. Objects are compared key by key, arrays by index (or by
 * content alignment when the adapter asks for it); `total` counts every difference while `entries`
 * keeps only the first `cap`, in document order.
 */
export function diffValues(before: Json, after: Json, options: DiffOptions): DiffReport {
  const prepare = options.prepare ?? ((m: Json) => m);
  const field = options.field ?? stripIndexes;
  const cap = options.cap ?? MAX_DIFF_ENTRIES;
  const entries: DiffEntry[] = [];
  const byField: Record<string, number> = {};
  const tally = { added: 0, removed: 0, changed: 0 };

  const record = (path: string, kind: DiffEntry['kind'], b?: Json, a?: Json) => {
    tally[kind]++;
    const key = `${kind} ${field(path)}`;
    byField[key] = (byField[key] ?? 0) + 1;
    if (entries.length >= cap) return;
    const entry: DiffEntry = { path, kind };
    if (kind !== 'added') entry.before = preview(b);
    if (kind !== 'removed') entry.after = preview(a);
    entries.push(entry);
  };

  const walkAligned = (
    b: Json[],
    a: Json[],
    path: string,
    sig: { before: string[]; after: string[] }
  ) => {
    const pairs = alignBySignature(sig.before, sig.after);
    const matchedB = new Set(pairs.map(([i]) => i));
    const matchedA = new Set(pairs.map(([, j]) => j));
    const events: Array<{ order: number; run: () => void }> = [];
    pairs.forEach(([i, j]) =>
      events.push({ order: i * 2, run: () => walk(b[i], a[j], `${path}[${i}]`) })
    );
    b.forEach(
      (v, i) =>
        !matchedB.has(i) &&
        events.push({ order: i * 2, run: () => record(`${path}[${i}]`, 'removed', v) })
    );
    a.forEach((v, j) => {
      if (matchedA.has(j)) return;
      // position an added item after the last matched item that precedes it
      const prev = pairs.filter(([, pj]) => pj < j).pop();
      events.push({
        order: (prev ? prev[0] : -1) * 2 + 1,
        run: () => record(`${path}[${j}]`, 'added', undefined, v),
      });
    });
    events.sort((x, y) => x.order - y.order).forEach((e) => e.run());
  };

  const walk = (b: Json | undefined, a: Json | undefined, path: string): void => {
    if (b === undefined && a === undefined) return;
    if (b === undefined) return record(path, 'added', undefined, a);
    if (a === undefined) return record(path, 'removed', b, undefined);
    if (Array.isArray(b) && Array.isArray(a)) {
      const sig = options.signatures?.(path, b, a) ?? null;
      if (sig) return walkAligned(b, a, path, sig);
      const len = Math.max(b.length, a.length);
      for (let i = 0; i < len; i++) walk(b[i], a[i], `${path}[${i}]`);
      return;
    }
    if (isObj(b) && isObj(a)) {
      const keys = [...new Set([...Object.keys(b), ...Object.keys(a)])].sort();
      for (const key of keys) walk(b[key], a[key], path ? `${path}.${key}` : key);
      return;
    }
    if (JSON.stringify(b) !== JSON.stringify(a)) record(path, 'changed', b, a);
  };

  walk(prepare(before), prepare(after), '');
  const total = tally.added + tally.removed + tally.changed;
  return {
    total,
    ...tally,
    byField: Object.fromEntries(Object.entries(byField).sort(([x], [y]) => x.localeCompare(y))),
    before: options.counts(before),
    after: options.counts(after),
    entries,
    truncated: total > entries.length,
  };
}
