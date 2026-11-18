import { createHash } from 'node:crypto';
import type { Json, JsonObject } from './types';

type Obj = Record<string, unknown>;

const isObj = (v: unknown): v is Obj => typeof v === 'object' && v !== null && !Array.isArray(v);

/** Container key -> the kind of id its children carry (`blocks` -> `block-1`, `block-2`, ...). */
const ID_KIND: Record<string, string> = {
  sections: 'section',
  blocks: 'block',
  rows: 'row',
  cells: 'cell',
  threads: 'thread',
  comments: 'comment',
};

/** A data URL's bytes are in the fixture; keep a short digest so the golden stays small but still pins them. */
export function digestDataUrl(src: string): string {
  const comma = src.indexOf(',');
  const head = src.slice(5, comma < 0 ? src.length : comma); // "image/png;base64"
  const mime = head.split(';')[0] || 'unknown';
  const body = comma < 0 ? '' : src.slice(comma + 1);
  const sha = createHash('sha256').update(src).digest('hex').slice(0, 16);
  return `data:${mime};sha256=${sha};chars=${body.length}`;
}

/**
 * Turn a docs-core `DocumentData` into a snapshot-stable value: ids become sequential placeholders
 * in document order (`section-1`, `block-1`, `row-1`, `cell-1`, ...), style-pool keys become
 * `text-style-N` / `paragraph-style-N` in first-reference order (unreferenced entries follow in
 * content order), volatile fields (document id, timestamps) are dropped, and data-URL image sources
 * are replaced by a digest. Key order is not touched here; `stableStringify` sorts keys.
 */
export function normalizeModel(data: unknown): Json {
  const root = JSON.parse(JSON.stringify(data ?? null)) as unknown;
  if (!isObj(root)) return root as Json;

  delete root.id;
  delete root.createdAt;
  delete root.updatedAt;

  const counters: Record<string, number> = {};
  const nextId = (kind: string) => `${kind}-${(counters[kind] = (counters[kind] ?? 0) + 1)}`;
  const textPool = new PoolRenamer(root.textStylePool, 'text-style');
  const paragraphPool = new PoolRenamer(root.paragraphStylePool, 'paragraph-style');

  const walk = (value: unknown, idKind: string | undefined): void => {
    if (Array.isArray(value)) {
      for (const item of value) walk(item, idKind);
      return;
    }
    if (!isObj(value)) return;
    if (typeof value.id === 'string') value.id = nextId(idKind ?? 'id');
    if (typeof value.styleId === 'string') {
      const inline = value.type === 'text' || value.type === 'link' || value.type === 'dynamicField';
      value.styleId = (inline ? textPool : paragraphPool).rename(value.styleId);
    }
    if (typeof value.src === 'string' && value.src.startsWith('data:')) value.src = digestDataUrl(value.src);
    for (const [key, child] of Object.entries(value)) {
      if (key === 'textStylePool' || key === 'paragraphStylePool') continue;
      walk(child, ID_KIND[key]);
    }
  };
  for (const [key, child] of Object.entries(root)) {
    if (key === 'textStylePool' || key === 'paragraphStylePool') continue;
    walk(child, ID_KIND[key]);
  }

  if (isObj(root.textStylePool)) root.textStylePool = textPool.finish();
  if (isObj(root.paragraphStylePool)) root.paragraphStylePool = paragraphPool.finish();
  return root as Json;
}

class PoolRenamer {
  private readonly map = new Map<string, string>();
  constructor(
    private readonly pool: unknown,
    private readonly prefix: string
  ) {}

  rename(id: string): string {
    let next = this.map.get(id);
    if (!next) {
      next = `${this.prefix}-${this.map.size + 1}`;
      this.map.set(id, next);
    }
    return next;
  }

  finish(): Obj {
    const pool = isObj(this.pool) ? this.pool : {};
    const out: Obj = {};
    // Entries referenced by the document keep document order; the rest are ordered by content.
    for (const [oldId, newId] of this.map) if (oldId in pool) out[newId] = pool[oldId];
    const rest = Object.keys(pool)
      .filter((k) => !this.map.has(k))
      .sort((a, b) => JSON.stringify(pool[a]).localeCompare(JSON.stringify(pool[b])));
    for (const oldId of rest) out[this.rename(oldId)] = pool[oldId];
    return out;
  }
}

/**
 * Inline each `styleId` as the pool entry it points at (`style`) and drop the pools, so two models
 * that use the same formatting through different pool keys compare equal.
 */
export function resolveStyles(model: Json): Json {
  if (!isObj(model)) return model;
  const textPool = isObj(model.textStylePool) ? model.textStylePool : {};
  const paragraphPool = isObj(model.paragraphStylePool) ? model.paragraphStylePool : {};
  const resolve = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(resolve);
    if (!isObj(value)) return value;
    const out: Obj = {};
    for (const [key, child] of Object.entries(value)) {
      if (key === 'styleId' && typeof child === 'string') {
        const inline = value.type === 'text' || value.type === 'link' || value.type === 'dynamicField';
        out.style = (inline ? textPool : paragraphPool)[child] ?? null;
      } else out[key] = resolve(child);
    }
    return out;
  };
  const out = resolve(model) as JsonObject;
  delete out.textStylePool;
  delete out.paragraphStylePool;
  return out;
}
