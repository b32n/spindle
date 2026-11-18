import { readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { loadManifest } from '../../../../scripts/docx-fidelity/manifest';
import { describeOutcome, runHarness } from '../../../../scripts/docx-fidelity/harness';
import { buildMatrix, injectMatrix, verifyClaims } from '../../../../scripts/docx-fidelity/matrix';
import { SNAPSHOT_KINDS } from '../../../../scripts/docx-fidelity/types';
import { sourceApi } from './support';

// Run with DOCX_FIDELITY_UPDATE=1 to rewrite the goldens (same as `npm run docs:docx-fidelity:update`).
const update = process.env.DOCX_FIDELITY_UPDATE === '1';
const fixturesDir = path.resolve(__dirname, '../../../../fixtures/docx');

describe('DOCX fidelity corpus', () => {
  const manifest = loadManifest(fixturesDir);

  it('lists at least one fixture', () => {
    expect(manifest.fixtures.length).toBeGreaterThan(0);
  });

  it('only claims constructs that are really in the files', async () => {
    expect(await verifyClaims(fixturesDir)).toEqual([]);
  });

  it('keeps the README coverage matrix in sync with the manifest', () => {
    const readmePath = path.join(fixturesDir, 'README.md');
    const readme = readFileSync(readmePath, 'utf8');
    const expected = injectMatrix(readme, buildMatrix([...manifest.fixtures].sort((a, b) => a.id.localeCompare(b.id))));
    if (update) writeFileSync(readmePath, expected);
    else expect(readme).toBe(expected);
  });

  describe('snapshots', () => {
    let outcomes: Awaited<ReturnType<typeof runHarness>>;
    beforeAll(async () => {
      outcomes = await runHarness({ api: sourceApi, fixturesDir, update });
    }, 120_000);

    it('has no orphaned goldens', () => {
      expect(outcomes.orphanGoldens).toEqual([]);
    });

    it.each(manifest.fixtures.map((f) => f.id).sort())('%s matches its goldens', (id) => {
      const outcome = outcomes.fixtures.find((f) => f.entry.id === id)!;
      expect(describeOutcome(outcome)).toEqual([]);
      for (const kind of SNAPSHOT_KINDS) expect(['match', 'created', 'updated']).toContain(outcome.status[kind]);
    });
  });
});
