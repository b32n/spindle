import { existsSync, readFileSync, readdirSync } from 'node:fs';
import path from 'node:path';
import { CONSTRUCTS } from './constructs';
import type { Manifest, ManifestEntry } from './types';

export class ManifestError extends Error {
  constructor(public readonly problems: string[]) {
    super(`Invalid DOCX fixture manifest:\n- ${problems.join('\n- ')}`);
    this.name = 'ManifestError';
  }
}

const ID_RE = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
const REQUIRED_STRINGS = ['id', 'filename', 'title', 'license', 'notes'] as const;

/** Problems with a parsed manifest (empty when valid). File presence is checked against `fixturesDir`. */
export function validateManifest(manifest: unknown, fixturesDir: string): string[] {
  const problems: string[] = [];
  const m = manifest as Partial<Manifest> | null;
  if (!m || typeof m !== 'object' || !Array.isArray(m.fixtures)) return ['manifest must be an object with a "fixtures" array'];
  if (m.version !== 1) problems.push(`unsupported manifest version ${String(m.version)} (expected 1)`);
  const ids = new Set<string>();
  const files = new Set<string>();
  m.fixtures.forEach((entry: Partial<ManifestEntry>, i) => {
    const at = `fixtures[${i}]${entry && entry.id ? ` (${entry.id})` : ''}`;
    for (const key of REQUIRED_STRINGS) {
      if (typeof entry?.[key] !== 'string' || !entry[key]) problems.push(`${at}: "${key}" must be a non-empty string`);
    }
    if (typeof entry?.id === 'string') {
      if (!ID_RE.test(entry.id)) problems.push(`${at}: id must be kebab-case`);
      if (ids.has(entry.id)) problems.push(`${at}: duplicate id`);
      ids.add(entry.id);
    }
    if (typeof entry?.filename === 'string') {
      if (!/^[A-Za-z0-9._-]+\.docx$/.test(entry.filename)) problems.push(`${at}: filename must be a plain *.docx name`);
      else if (files.has(entry.filename)) problems.push(`${at}: duplicate filename`);
      else if (!existsSync(path.join(fixturesDir, entry.filename))) problems.push(`${at}: file ${entry.filename} does not exist`);
      files.add(entry.filename);
    }
    if (!entry?.source || typeof entry.source.tool !== 'string' || !entry.source.tool || typeof entry.source.os !== 'string' || !entry.source.os) {
      problems.push(`${at}: "source" needs non-empty "tool" and "os"`);
    }
    if (!Array.isArray(entry?.constructs) || entry.constructs.length === 0) problems.push(`${at}: "constructs" must be a non-empty array`);
    else {
      for (const c of entry.constructs) if (!CONSTRUCTS.some((k) => k.id === c)) problems.push(`${at}: unknown construct "${c}"`);
    }
    if (entry?.expectImportError !== undefined && typeof entry.expectImportError !== 'boolean') problems.push(`${at}: "expectImportError" must be a boolean`);
  });
  if (existsSync(fixturesDir)) {
    for (const f of readdirSync(fixturesDir).filter((n) => n.toLowerCase().endsWith('.docx')).sort()) {
      if (!files.has(f)) problems.push(`${f} is in the fixtures directory but not in the manifest`);
    }
  }
  return problems;
}

/** Read and validate `<fixturesDir>/manifest.json`; throws `ManifestError` listing every problem. */
export function loadManifest(fixturesDir: string): Manifest {
  const file = path.join(fixturesDir, 'manifest.json');
  let parsed: unknown;
  try {
    parsed = JSON.parse(readFileSync(file, 'utf8'));
  } catch (e) {
    throw new ManifestError([`cannot read ${file}: ${e instanceof Error ? e.message : String(e)}`]);
  }
  const problems = validateManifest(parsed, fixturesDir);
  if (problems.length) throw new ManifestError(problems);
  return parsed as Manifest;
}

/** The fixtures to run: everything in the manifest, in id order, optionally narrowed by id. */
export function discoverFixtures(fixturesDir: string, only?: string[]): ManifestEntry[] {
  const all = [...loadManifest(fixturesDir).fixtures].sort((a, b) => a.id.localeCompare(b.id));
  if (!only || only.length === 0) return all;
  const missing = only.filter((id) => !all.some((f) => f.id === id));
  if (missing.length) throw new ManifestError([`unknown fixture id(s): ${missing.join(', ')}`]);
  return all.filter((f) => only.includes(f.id));
}
