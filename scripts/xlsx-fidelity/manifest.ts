import { existsSync, readFileSync, readdirSync } from 'node:fs';
import path from 'node:path';
import type { Manifest, ManifestEntry, ManifestOptions } from './types';

export class ManifestError extends Error {
  constructor(
    public readonly problems: string[],
    label = 'fixture'
  ) {
    super(`Invalid ${label} manifest:\n- ${problems.join('\n- ')}`);
    this.name = 'ManifestError';
  }
}

const ID_RE = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
const FILE_RE = /^[A-Za-z0-9._-]+$/;
const REQUIRED_STRINGS = ['id', 'filename', 'title', 'license', 'notes'] as const;
const OPTIONAL_BOOLEANS = ['expectImportError', 'generatedAtRuntime'] as const;

const ext = (name: string) => path.extname(name).toLowerCase();

/**
 * Problems with a parsed manifest (empty when valid). Committed files must exist and every file in the
 * directory must be listed; `generatedAtRuntime` entries are exempt from the file checks but must name a generator.
 */
export function validateManifest(
  manifest: unknown,
  fixturesDir: string,
  opts: ManifestOptions
): string[] {
  const problems: string[] = [];
  const m = manifest as Partial<Manifest> | null;
  if (!m || typeof m !== 'object' || !Array.isArray(m.fixtures))
    return ['manifest must be an object with a "fixtures" array'];
  if (m.version !== 1)
    problems.push(`unsupported manifest version ${String(m.version)} (expected 1)`);
  const ids = new Set<string>();
  const files = new Set<string>();
  m.fixtures.forEach((entry: Partial<ManifestEntry>, i) => {
    const at = `fixtures[${i}]${entry && entry.id ? ` (${entry.id})` : ''}`;
    for (const key of REQUIRED_STRINGS) {
      if (typeof entry?.[key] !== 'string' || !entry[key])
        problems.push(`${at}: "${key}" must be a non-empty string`);
    }
    for (const key of OPTIONAL_BOOLEANS) {
      if (entry?.[key] !== undefined && typeof entry[key] !== 'boolean')
        problems.push(`${at}: "${key}" must be a boolean`);
    }
    const runtime = entry?.generatedAtRuntime === true;
    if (typeof entry?.id === 'string') {
      if (!ID_RE.test(entry.id)) problems.push(`${at}: id must be kebab-case`);
      if (ids.has(entry.id)) problems.push(`${at}: duplicate id`);
      ids.add(entry.id);
    }
    if (typeof entry?.filename === 'string') {
      if (!FILE_RE.test(entry.filename) || !opts.extensions.includes(ext(entry.filename))) {
        problems.push(
          `${at}: filename must be a plain name ending in ${opts.extensions.join(' / ')}`
        );
      } else if (files.has(entry.filename)) problems.push(`${at}: duplicate filename`);
      else if (!runtime && !existsSync(path.join(fixturesDir, entry.filename)))
        problems.push(`${at}: file ${entry.filename} does not exist`);
      else if (runtime && existsSync(path.join(fixturesDir, entry.filename)))
        problems.push(`${at}: generatedAtRuntime file ${entry.filename} must not be committed`);
      files.add(entry.filename);
    }
    if (runtime && (typeof entry.generator !== 'string' || !entry.generator))
      problems.push(`${at}: a generatedAtRuntime fixture needs a "generator"`);
    const src = entry?.source;
    if (
      !src ||
      typeof src.tool !== 'string' ||
      !src.tool ||
      typeof src.os !== 'string' ||
      !src.os
    ) {
      problems.push(`${at}: "source" needs non-empty "tool" and "os"`);
    } else if (src.version !== undefined && (typeof src.version !== 'string' || !src.version))
      problems.push(`${at}: "source.version" must be a non-empty string`);
    if (!Array.isArray(entry?.constructs) || entry.constructs.length === 0)
      problems.push(`${at}: "constructs" must be a non-empty array`);
    else {
      for (const c of entry.constructs)
        if (!opts.constructIds.includes(c)) problems.push(`${at}: unknown construct "${c}"`);
    }
  });
  if (existsSync(fixturesDir)) {
    for (const f of readdirSync(fixturesDir)
      .filter((n) => opts.extensions.includes(ext(n)))
      .sort()) {
      if (!files.has(f)) problems.push(`${f} is in the fixtures directory but not in the manifest`);
    }
  }
  return problems;
}

/** Read and validate `<fixturesDir>/manifest.json`; throws `ManifestError` listing every problem. */
export function loadManifest(fixturesDir: string, opts: ManifestOptions): Manifest {
  const file = path.join(fixturesDir, 'manifest.json');
  let parsed: unknown;
  try {
    parsed = JSON.parse(readFileSync(file, 'utf8'));
  } catch (e) {
    throw new ManifestError(
      [`cannot read ${file}: ${e instanceof Error ? e.message : String(e)}`],
      opts.label
    );
  }
  const problems = validateManifest(parsed, fixturesDir, opts);
  if (problems.length) throw new ManifestError(problems, opts.label);
  return parsed as Manifest;
}
