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

  it('imports every presentation, rejects the four files that are not one with a named message, and never throws in the detector', () => {
    const rejected = new Map([
      ['edge-truncated-zip', /damaged/],
      ['edge-no-presentation', /not a PowerPoint presentation/],
      ['edge-encrypted', /password-protected/],
      ['edge-legacy-ppt', /legacy binary/],
    ]);
    for (const o of run.fixtures) {
      const pattern = rejected.get(o.entry.id);
      if (pattern) {
        expect(o.result.model.importError).toMatch(pattern);
        expect(o.entry.expectImportError).toBe(true);
        expect(o.result.model.presentation).toBeNull();
        expect(o.result.roundtrip.status).toBe('skipped-import-error');
      } else {
        expect({ id: o.entry.id, error: o.result.model.importError }).toEqual({ id: o.entry.id, error: null });
        expect(o.result.model.presentation).not.toBeNull();
        // the writer (K12) round-trips every readable deck
        expect(o.result.roundtrip.status).toBe('ok');
      }
      expect(o.importExpectation).toBe('as-expected');
      expect(o.result.loss.items.map((i) => i.id).filter((i) => i.startsWith('harness.'))).toEqual([]);
      // the parts of the kept content (charts, SmartArt, OLE, clips, ink) are opaque and nothing else is
      const opaqueParts = o.result.ledger.parts.filter((p) => p.status === 'opaque');
      expect(opaqueParts).toHaveLength(o.result.ledger.totals.opaque);
      expect(opaqueParts.every((p) => p.rule === 'opaque')).toBe(true);
    }
  });

  it('reads the slide list, size, theme, sections and hidden flags of the corpus decks', () => {
    const model = (id: string) =>
      run.fixtures.find((f) => f.entry.id === id)?.result.model.presentation as unknown as {
        slideSize: { w: number; h: number };
        slides: { pptxId: string; hidden?: boolean; error?: string; layoutPart?: string }[];
        sections?: { name: string; slideIds: string[] }[];
        theme: { name: string; colors: Record<string, string>; fonts: { major: string; minor: string } };
      };
    expect(model('sizes-4x3').slideSize).toEqual({ w: 960, h: 720 });
    expect(model('sizes-a4').slideSize).toEqual({ w: 1123, h: 794 });
    expect(model('sizes-portrait').slideSize.h).toBeGreaterThan(model('sizes-portrait').slideSize.w);
    expect(model('hidden-sections').slides.map((s) => s.hidden === true)).toEqual([false, true, false, false, false]);
    expect(model('hidden-sections').sections?.map((s) => s.name)).toEqual(['Intro', 'Body', 'Wrap-up']);
    expect(model('edge-slide-missing').slides.map((s) => s.error)).toEqual([undefined, 'the part is missing', undefined]);
    // the XML parser is lenient: a slide with an unclosed element is still read as far as it goes
    expect(model('edge-invalid-xml').slides.length).toBeGreaterThan(0);
    expect(model('themes').theme.colors.accent1).toMatch(/^#[0-9A-F]{6}$/);
    for (const o of run.fixtures) {
      const m = o.result.model.presentation as unknown as { slides?: { layoutPart?: string; error?: string }[] } | null;
      if (m) expect(m.slides?.filter((s) => !s.error && !s.layoutPart && o.entry.id !== 'edge-bad-rels')).toEqual([]);
    }
  });

  it('keeps every source text run in the model (the automated parity check) and reads the text decks\' shapes', () => {
    let checked = 0;
    for (const o of run.fixtures) {
      const parity = o.result.model.textParity;
      if (o.result.model.presentation === null) {
        expect(parity).toBeUndefined();
        continue;
      }
      expect({ id: o.entry.id, missing: parity?.missing, extra: parity?.extra }).toEqual({ id: o.entry.id, missing: [], extra: [] });
      expect(parity?.modelRuns).toBe(parity?.sourceRuns);
      checked += parity?.sourceRuns ?? 0;
    }
    expect(checked).toBeGreaterThan(100);
    const shapes = (id: string) => (run.fixtures.find((f) => f.entry.id === id)?.result.model.presentation as unknown as { slides: { elements?: { text?: unknown; geometry?: { prst?: string; known?: boolean } }[] }[] }).slides.flatMap((s) => s.elements ?? []);
    expect(shapes('text-basic').filter((e) => e.text).length).toBeGreaterThan(0);
    expect(shapes('shapes-presets').some((e) => e.geometry?.known === false)).toBe(true);
    expect(shapes('shapes-presets').some((e) => e.geometry?.known === true)).toBe(true);
  });

  it('reads the pictures and tables of the image and table decks (plan K10a)', () => {
    type Element = { kind: string; dataUrl?: { dataUrl: string; bytes: number }; unsupported?: string; crop?: unknown; rows?: { cells: { gridSpan?: number; fill?: unknown }[] }[] };
    const elements = (id: string) => (run.fixtures.find((f) => f.entry.id === id)?.result.model.presentation as unknown as { slides: { elements?: Element[] }[] }).slides.flatMap((s) => s.elements ?? []);
    const pictures = elements('images').filter((e) => e.kind === 'picture');
    // embedded images are digested in the golden; the metafiles and the linked picture are placeholders
    expect(pictures.filter((p) => p.dataUrl).map((p) => p.dataUrl?.dataUrl)).toEqual(expect.arrayContaining(['image/png', 'image/jpeg', 'image/gif']));
    expect(pictures.filter((p) => p.unsupported).map((p) => p.unsupported).sort()).toEqual(['emf', 'linked', 'wmf']);
    expect(pictures.some((p) => p.crop)).toBe(true);
    const tables = elements('tables').filter((e) => e.kind === 'table');
    expect(tables.length).toBeGreaterThan(2);
    expect(tables.some((t) => t.rows?.some((r) => r.cells.some((c) => c.gridSpan === 2)))).toBe(true);
    expect(tables.some((t) => t.rows?.some((r) => r.cells.some((c) => c.fill)))).toBe(true);
  });

  it('names the encrypted and the legacy binary stubs, and nothing else about them', () => {
    const lossIds = (id: string) =>
      run.fixtures.find((f) => f.entry.id === id)?.result.loss.items.map((i) => i.id);
    expect(lossIds('edge-encrypted')).toEqual(['pptx.encrypted']);
    expect(lossIds('edge-legacy-ppt')).toEqual(['pptx.legacyBinary']);
    expect(lossIds('edge-truncated-zip')).toEqual([]);
  });

  it('leaves no part unaccounted in any file the reader opens, and covers every dropped part by exactly one loss id', () => {
    for (const o of run.fixtures) {
      if (o.result.model.importError !== null) continue;
      for (const p of o.result.ledger.parts) {
        if (p.status === 'dropped') expect(p.loss?.length).toBe(1);
        else expect({ id: o.entry.id, part: p.name, status: p.status === 'opaque' ? 'mapped' : p.status }).toEqual({ id: o.entry.id, part: p.name, status: 'mapped' });
      }
      expect(o.result.ledger.totals.unaccounted).toBe(0);
    }
  });

  it('builds a report whose scoreboard rows are the plan rows this run can speak to', () => {
    const report = toPptxReport(run);
    expect(report.scoreboard.map((r) => r.id)).toEqual(['L11', 'L12', 'L13', 'L14']);
    // L12 holds (every file ends as its manifest says); L13 does not: the rejected no-presentation package has 26 parts
    // that nothing accounts for, which is correct (nothing was imported) and recorded
    expect(report.scoreboard.map((r) => r.status)).toEqual(['not-measured', 'met', 'not-met', 'not-met']);
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
