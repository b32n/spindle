import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs';
import path from 'node:path';
import type { Json, SnapshotStatus } from './types';

/** Pretty-printed JSON with recursively sorted keys and a trailing newline: byte-stable across runs. */
export function stableStringify(value: unknown): string {
  return JSON.stringify(sortKeys(value), null, 2) + '\n';
}

function sortKeys(value: unknown): Json {
  if (Array.isArray(value)) return value.map(sortKeys);
  if (value && typeof value === 'object') {
    const out: { [key: string]: Json } = {};
    for (const key of Object.keys(value).sort()) {
      const v = (value as Record<string, unknown>)[key];
      if (v !== undefined) out[key] = sortKeys(v);
    }
    return out;
  }
  return value === undefined ? null : (value as Json);
}

export function snapshotPath(goldensDir: string, id: string, kind: string): string {
  return path.join(goldensDir, `${id}.${kind}.json`);
}

/**
 * Compare `value` to the golden on disk. In update mode the golden is (re)written when it is
 * missing or different, and the returned status says what happened; otherwise a missing golden is
 * `missing` and a different one `mismatch`.
 */
export function compareSnapshot(file: string, value: unknown, update: boolean): SnapshotStatus {
  const next = stableStringify(value);
  const exists = existsSync(file);
  const current = exists ? readFileSync(file, 'utf8') : null;
  if (current === next) return 'match';
  if (!update) return exists ? 'mismatch' : 'missing';
  mkdirSync(path.dirname(file), { recursive: true });
  writeFileSync(file, next);
  return exists ? 'updated' : 'created';
}

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Golden files (`<id>.<kind>.json`) whose fixture id is not in `ids` (a fixture was removed or renamed). */
export function findOrphanGoldens(
  goldensDir: string,
  ids: Iterable<string>,
  kinds: readonly string[]
): string[] {
  if (!existsSync(goldensDir)) return [];
  const known = new Set(ids);
  const re = new RegExp(`\\.(?:${kinds.map(escapeRe).join('|')})\\.json$`);
  return readdirSync(goldensDir)
    .filter((f) => re.test(f))
    .filter((f) => !known.has(f.replace(re, '')))
    .sort();
}

export function removeGoldens(goldensDir: string, files: string[]): void {
  for (const f of files) unlinkSync(path.join(goldensDir, f));
}
