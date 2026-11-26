import { readFileSync } from 'node:fs';
import path from 'node:path';
import type { OoxmlParts } from './ooxml-package';
import { readParts } from './ooxml-package';
import type { Construct, Manifest } from './types';

/** Construct ids a fixture claims whose probe finds nothing in the package. */
export function unverifiedClaims(
  claimed: string[],
  registry: Array<Construct<OoxmlParts>>,
  parts: OoxmlParts
): string[] {
  return claimed.filter((id) => (registry.find((c) => c.id === id)?.count(parts) ?? 0) === 0);
}

/** Constructs actually found in a package (for reporting; claims are what the manifest lists). */
export function detectConstructs(
  registry: Array<Construct<OoxmlParts>>,
  parts: OoxmlParts
): string[] {
  return registry.filter((c) => c.count(parts) > 0).map((c) => c.id);
}

/** Every construct a committed fixture claims must be present in its package; returns the false claims. */
export async function verifyClaims(
  fixturesDir: string,
  manifest: Manifest,
  registry: Array<Construct<OoxmlParts>>
): Promise<string[]> {
  const problems: string[] = [];
  for (const f of manifest.fixtures) {
    if (f.generatedAtRuntime) continue;
    const parts = await readParts(new Uint8Array(readFileSync(path.join(fixturesDir, f.filename))));
    for (const c of unverifiedClaims(f.constructs, registry, parts))
      problems.push(`${f.id}: claims "${c}" but the file does not contain it`);
  }
  return problems;
}
