import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describeOutcome, runHarness } from '../../../../scripts/fidelity-core/harness';
import { PPTX_KINDS, pptxAdapter, runPptxFixture } from '../../../../scripts/pptx-fidelity/run';
import { toPptxReport } from '../../../../scripts/pptx-fidelity/report';
import {
  PPTX_FIXTURES_DIR,
  PPTX_MANIFEST_OPTIONS,
  loadPptxManifest,
} from '../../../../scripts/pptx-fidelity';
import { sourceApi } from './support';

// Run with PPTX_FIDELITY_UPDATE=1 to rewrite the goldens (same as `npm run slides:pptx-fidelity:update`).
const update = process.env.PPTX_FIDELITY_UPDATE === '1';

describe('PPTX fidelity corpus (io source)', () => {
  const manifest = loadPptxManifest();
  let run: Awaited<ReturnType<typeof runHarness<Awaited<ReturnType<typeof runPptxFixture>>>>>;

  beforeAll(async () => {
    run = await runHarness({
      adapter: pptxAdapter(sourceApi),
      fixturesDir: PPTX_FIXTURES_DIR,
      manifestOptions: PPTX_MANIFEST_OPTIONS,
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

  it('writes four goldens per committed file', () => {
    expect(PPTX_KINDS).toEqual(['model', 'loss', 'roundtrip', 'ledger']);
    expect(committed.length).toBeGreaterThanOrEqual(39);
  });

  it.each(committed)('%s matches its goldens', (id) => {
    const outcome = run.fixtures.find((f) => f.entry.id === id);
    expect(outcome).toBeDefined();
    if (!outcome) return;
    expect(describeOutcome(outcome, PPTX_KINDS)).toEqual([]);
    for (const kind of PPTX_KINDS)
      expect(['match', 'created', 'updated']).toContain(outcome.status[kind]);
  });

  it('records "no importer" for every file, skips every round trip and never throws in the detector', () => {
    for (const o of run.fixtures) {
      expect({ id: o.entry.id, error: o.result.model.importError }).toEqual({
        id: o.entry.id,
        error: 'no importer',
      });
      expect(o.result.model.presentation).toBeNull();
      expect(o.result.roundtrip.status).toBe('skipped-import-error');
      expect(o.result.loss.items.map((i) => i.id).filter((i) => i.startsWith('harness.'))).toEqual(
        []
      );
      expect(o.result.ledger.importError).toBe(true);
      // nothing is mapped without an importer, and nothing is kept opaque without the part store
      expect(o.result.ledger.totals).toMatchObject({ mapped: 0, opaque: 0 });
    }
  });

  it('names the encrypted and the legacy binary stubs, and nothing else about them', () => {
    const lossIds = (id: string) =>
      run.fixtures.find((f) => f.entry.id === id)?.result.loss.items.map((i) => i.id);
    expect(lossIds('edge-encrypted')).toEqual(['pptx.encrypted']);
    expect(lossIds('edge-legacy-ppt')).toEqual(['pptx.legacyBinary']);
    expect(lossIds('edge-truncated-zip')).toEqual([]);
  });

  it('accounts for every part that a loss id covers, so the unaccounted ones are exactly the unimported ones', () => {
    for (const o of run.fixtures)
      for (const p of o.result.ledger.parts)
        if (p.status === 'dropped') expect(p.loss?.length).toBe(1);
        else expect(p.status).toBe('unaccounted');
  });

  it('builds a report whose scoreboard rows are the plan rows this run can speak to', () => {
    const report = toPptxReport(run);
    expect(report.scoreboard.map((r) => r.id)).toEqual(['L11', 'L12', 'L13', 'L14']);
    expect(report.scoreboard.map((r) => r.status)).toEqual([
      'not-measured',
      'not-measured',
      'not-met',
      'not-measured',
    ]);
    expect(report.files).toHaveLength(committed.length);
  });

  it('is deterministic: a second run of a file gives identical results', async () => {
    for (const id of ['kitchen-sink', 'python-pptx-deck', 'edge-invalid-xml', 'edge-encrypted']) {
      const entry = manifest.fixtures.find((f) => f.id === id);
      const bytes = new Uint8Array(
        readFileSync(path.join(PPTX_FIXTURES_DIR, entry?.filename ?? ''))
      );
      const a = await runPptxFixture(sourceApi, bytes);
      const b = await runPptxFixture(sourceApi, bytes);
      expect(b).toEqual(a);
    }
  });
});
