import path from 'node:path';
import { describeOutcome, runHarness } from '../../../../scripts/fidelity-core/harness';
import { XLSX_KINDS, runXlsxFixture } from '../../../../scripts/xlsx-fidelity/run';
import { toXlsxReport } from '../../../../scripts/xlsx-fidelity/report';
import {
  XLSX_FIXTURES_DIR,
  XLSX_MANIFEST_OPTIONS,
  loadXlsxManifest,
} from '../../../../scripts/xlsx-fidelity';
import { xlsxAdapter } from '../../../../scripts/xlsx-fidelity/run';
import { readFileSync } from 'node:fs';
import { sourceApi } from './support';

// Run with XLSX_FIDELITY_UPDATE=1 to rewrite the goldens (same as `npm run sheets:xlsx-fidelity:update`).
const update = process.env.XLSX_FIDELITY_UPDATE === '1';

describe('XLSX fidelity corpus (sheets-core source)', () => {
  const manifest = loadXlsxManifest();
  let run: Awaited<ReturnType<typeof runHarness<Awaited<ReturnType<typeof runXlsxFixture>>>>>;

  beforeAll(async () => {
    run = await runHarness({
      adapter: xlsxAdapter(sourceApi),
      fixturesDir: XLSX_FIXTURES_DIR,
      manifestOptions: XLSX_MANIFEST_OPTIONS,
      update,
    });
  }, 300_000);

  it('has no orphaned goldens', () => {
    expect(run.orphanGoldens).toEqual([]);
  });

  const committed = manifest.fixtures
    .filter((f) => !f.generatedAtRuntime)
    .map((f) => f.id)
    .sort();

  it.each(committed)('%s matches its goldens', (id) => {
    const outcome = run.fixtures.find((f) => f.entry.id === id);
    expect(outcome).toBeDefined();
    if (!outcome) return;
    expect(describeOutcome(outcome, XLSX_KINDS)).toEqual([]);
    for (const kind of XLSX_KINDS)
      expect(['match', 'created', 'updated']).toContain(outcome.status[kind]);
  });

  it('records the importer behaviour the manifest expects for the damaged files', () => {
    for (const o of run.fixtures.filter((f) => f.entry.expectImportError))
      expect({ id: o.entry.id, threw: o.result.model.importError !== null }).toEqual({
        id: o.entry.id,
        threw: true,
      });
  });

  it('builds a report whose scoreboard rows are the plan rows this run measures', () => {
    const report = toXlsxReport(run);
    expect(report.scoreboard.map((r) => r.id)).toEqual(['L4', 'L5', 'L6', 'L7', 'L8']);
    expect(report.files).toHaveLength(committed.length);
  });

  it('is deterministic: a second run of a file gives identical results', async () => {
    for (const id of [
      'cell-types',
      'formulas-modern',
      'openpyxl-image',
      'edge-duplicate-sheet-names',
    ]) {
      const entry = manifest.fixtures.find((f) => f.id === id);
      const bytes = new Uint8Array(
        readFileSync(path.join(XLSX_FIXTURES_DIR, entry?.filename ?? ''))
      );
      const a = await runXlsxFixture(sourceApi, bytes);
      const b = await runXlsxFixture(sourceApi, bytes);
      expect(b).toEqual(a);
    }
  });
});
