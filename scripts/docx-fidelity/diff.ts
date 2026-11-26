import { resolveStyles } from './normalize';
import type { DiffEntry, DiffReport, Json, ModelCounts } from './types';

type Obj = { [key: string]: Json };
const isObj = (v: Json | undefined): v is Obj => typeof v === 'object' && v !== null && !Array.isArray(v);

export const MAX_DIFF_ENTRIES = 40;

const preview = (v: Json | undefined): string => {
  const s = JSON.stringify(v) ?? 'undefined';
  return s.length > 80 ? `${s.slice(0, 77)}...` : s;
};

/** Plain text of a block (inline text, link text, table cells and the blocks inside them, content controls and what they wrap), used to align blocks. */
function blockText(block: Json): string {
  if (!isObj(block)) return '';
  const inline = (items: Json | undefined): string =>
    Array.isArray(items)
      ? items.map((i) => (isObj(i) && typeof i.text === 'string' ? i.text : isObj(i) && Array.isArray(i.content) ? inline(i.content) : isObj(i) && i.type ? `[${String(i.type)}]` : '')).join('')
      : '';
  if (Array.isArray(block.blocks)) return block.blocks.map(blockText).join('');
  if (Array.isArray(block.rows)) {
    return block.rows
      .map((row) =>
        isObj(row) && Array.isArray(row.cells)
          ? row.cells.map((c) => (isObj(c) ? (Array.isArray(c.blocks) ? c.blocks.map(blockText).join('\n') : inline(c.content)) : '')).join('\t')
          : ''
      )
      .join('\n');
  }
  return inline(block.content);
}

const signature = (block: Json): string => `${isObj(block) ? String(block.type) : '?'}|${blockText(block)}`;

/** Longest-common-subsequence alignment of two block lists; returns matched index pairs. */
function alignBlocks(before: Json[], after: Json[]): Array<[number, number]> {
  const a = before.map(signature);
  const b = after.map(signature);
  const n = a.length;
  const m = b.length;
  const table: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      table[i][j] = a[i] === b[j] ? table[i + 1][j + 1] + 1 : Math.max(table[i + 1][j], table[i][j + 1]);
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

function countModel(model: Json): ModelCounts {
  const counts: ModelCounts = { sections: 0, blocks: {}, textChars: 0 };
  if (!isObj(model) || !Array.isArray(model.sections)) return counts;
  counts.sections = model.sections.length;
  for (const section of model.sections) {
    if (!isObj(section) || !Array.isArray(section.blocks)) continue;
    for (const block of section.blocks) {
      const type = isObj(block) ? String(block.type) : 'unknown';
      counts.blocks[type] = (counts.blocks[type] ?? 0) + 1;
      counts.textChars += blockText(block).length;
    }
  }
  counts.blocks = Object.fromEntries(Object.entries(counts.blocks).sort(([x], [y]) => x.localeCompare(y)));
  return counts;
}

/**
 * Structural diff of two normalized docs-core models. Styles are resolved first (so pool keys do not
 * matter) and section block lists are aligned by content, so one dropped paragraph is one entry, not
 * a cascade of index shifts.
 */
export function diffModels(before: Json, after: Json): DiffReport {
  const entries: DiffEntry[] = [];
  const byField: Record<string, number> = {};
  const tally = { added: 0, removed: 0, changed: 0 };

  const record = (path: string, kind: DiffEntry['kind'], b?: Json, a?: Json) => {
    tally[kind]++;
    const field = `${kind} ${path.replace(/\[\d+\]/g, '[]')}`;
    byField[field] = (byField[field] ?? 0) + 1;
    const entry: DiffEntry = { path, kind };
    if (kind !== 'added') entry.before = preview(b);
    if (kind !== 'removed') entry.after = preview(a);
    entries.push(entry);
  };

  const walk = (b: Json | undefined, a: Json | undefined, path: string): void => {
    if (b === undefined && a === undefined) return;
    if (b === undefined) return record(path, 'added', undefined, a);
    if (a === undefined) return record(path, 'removed', b, undefined);
    if (Array.isArray(b) && Array.isArray(a)) {
      if (path.endsWith('.blocks')) {
        const pairs = alignBlocks(b, a);
        const matchedB = new Set(pairs.map(([i]) => i));
        const matchedA = new Set(pairs.map(([, j]) => j));
        const events: Array<{ order: number; run: () => void }> = [];
        pairs.forEach(([i, j]) => events.push({ order: i * 2, run: () => walk(b[i], a[j], `${path}[${i}]`) }));
        b.forEach((v, i) => !matchedB.has(i) && events.push({ order: i * 2, run: () => record(`${path}[${i}]`, 'removed', v) }));
        a.forEach((v, j) => {
          if (matchedA.has(j)) return;
          // position an added block after the last matched block that precedes it
          const prev = pairs.filter(([, pj]) => pj < j).pop();
          events.push({ order: (prev ? prev[0] : -1) * 2 + 1, run: () => record(`${path}[${j}]`, 'added', undefined, v) });
        });
        events.sort((x, y) => x.order - y.order).forEach((e) => e.run());
        return;
      }
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

  walk(resolveStyles(before), resolveStyles(after), '');
  const total = tally.added + tally.removed + tally.changed;
  return {
    total,
    ...tally,
    byField: Object.fromEntries(Object.entries(byField).sort(([x], [y]) => x.localeCompare(y))),
    before: countModel(before),
    after: countModel(after),
    entries: entries.slice(0, MAX_DIFF_ENTRIES),
    truncated: total > MAX_DIFF_ENTRIES,
  };
}
