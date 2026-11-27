import { readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { verifyClaims } from '../../../../scripts/fidelity-core/claims';
import { describeOutcome, runHarness } from '../../../../scripts/fidelity-core/harness';
import { loadManifest } from '../../../../scripts/fidelity-core/manifest';
import { buildMatrix, injectMatrix } from '../../../../scripts/fidelity-core/matrix';
import { CONSTRUCTS } from '../../../../scripts/docx-fidelity/constructs';
import { DOCX_FIXTURES_DIR, DOCX_MANIFEST_OPTIONS, MATRIX_END, MATRIX_START } from '../../../../scripts/docx-fidelity';
import { docxAdapter } from '../../../../scripts/docx-fidelity/run';
import { DOCX_KINDS } from '../../../../scripts/docx-fidelity/types';
import { sourceApi } from './support';

// Run with DOCX_FIDELITY_UPDATE=1 to rewrite the goldens (same as `npm run docs:docx-fidelity:update`).
const update = process.env.DOCX_FIDELITY_UPDATE === '1';
const fixturesDir = DOCX_FIXTURES_DIR;

describe('DOCX fidelity corpus', () => {
  const manifest = loadManifest(fixturesDir, DOCX_MANIFEST_OPTIONS);

  it('lists at least one fixture', () => {
    expect(manifest.fixtures.length).toBeGreaterThan(0);
  });

  it('only claims constructs that are really in the files', async () => {
    expect(await verifyClaims(fixturesDir, manifest, CONSTRUCTS)).toEqual([]);
  });

  it('keeps the README coverage matrix in sync with the manifest', () => {
    const readmePath = path.join(fixturesDir, 'README.md');
    const readme = readFileSync(readmePath, 'utf8');
    const expected = injectMatrix(readme, buildMatrix(CONSTRUCTS, manifest.fixtures), MATRIX_START, MATRIX_END);
    if (update) writeFileSync(readmePath, expected);
    else expect(readme).toBe(expected);
  });

  describe('snapshots', () => {
    let outcomes: Awaited<ReturnType<typeof runHarness>>;
    beforeAll(async () => {
      outcomes = await runHarness({ adapter: docxAdapter(sourceApi), fixturesDir, manifestOptions: DOCX_MANIFEST_OPTIONS, strictImport: true, update });
    }, 120_000);

    it('has no orphaned goldens', () => {
      expect(outcomes.orphanGoldens).toEqual([]);
    });

    it.each(manifest.fixtures.map((f) => f.id).sort())('%s matches its goldens', (id) => {
      const outcome = outcomes.fixtures.find((f) => f.entry.id === id)!;
      expect(describeOutcome(outcome, DOCX_KINDS)).toEqual([]);
      for (const kind of DOCX_KINDS) expect(['match', 'created', 'updated']).toContain(outcome.status[kind]);
    });
  });
});
