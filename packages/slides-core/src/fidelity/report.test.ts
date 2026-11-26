import type { FixtureOutcome, HarnessReport } from '../../../../scripts/fidelity-core/harness';
import type { ReportFile } from '../../../../scripts/fidelity-core/report';
import { pptxScoreboard, toPptxReport, toReportFile } from '../../../../scripts/pptx-fidelity/report';
import { crashedResult, type PptxResult } from '../../../../scripts/pptx-fidelity/run';

const result = (over: Partial<PptxResult> = {}): PptxResult => ({
  loss: { format: 'pptx', items: [{ id: 'pptx.charts', label: 'Charts', count: 1, effect: 'e' }] },
  model: { importError: 'no importer', warnings: [], presentation: null },
  roundtrip: { status: 'skipped-import-error', exportWarnings: [], reimportWarnings: [], diff: null },
  ledger: {
    package: 'zip',
    importError: true,
    totals: { mapped: 0, opaque: 0, dropped: 1, unaccounted: 2, total: 3 },
    parts: [
      { name: 'b', status: 'dropped', rule: 'r', loss: ['x'] },
      ...Array.from({ length: 14 }, (_, i) => ({ name: `u${String(i).padStart(2, '0')}`, status: 'unaccounted' as const, rule: 'none' })),
    ],
  },
  ...over,
});
const outcome = (id: string, r: PptxResult, over: Partial<FixtureOutcome<PptxResult>> = {}): FixtureOutcome<PptxResult> => ({
  entry: { id, filename: `${id}.pptx`, title: `T ${id}`, source: { tool: 'tool', os: 'o' }, license: 'l', constructs: [], notes: '' },
  result: r,
  status: { model: 'match', loss: 'match', roundtrip: 'match', ledger: 'match' },
  problems: [],
  importExpectation: 'unexpected-error',
  ...over,
});

describe('toReportFile', () => {
  it('maps a harness outcome onto the report file, listing at most twelve unaccounted parts, with no values section', () => {
    const f = toReportFile(outcome('a', result()));
    expect(f).toMatchObject({ id: 'a', title: 'T a', producer: 'tool', importError: 'no importer', importExpectation: 'unexpected-error' });
    expect(f.ledger).toMatchObject({ dropped: 1, unaccounted: 2, total: 3 });
    expect(f.ledger?.unaccountedParts).toHaveLength(12);
    expect(f.values).toBeUndefined();
  });
});

describe('pptxScoreboard', () => {
  const files = (list: Array<[string, Partial<ReportFile>]>): ReportFile[] => list.map(([id, over]) => ({ ...toReportFile(outcome(id, result())), ...over }));

  it('says L11, L12 and L14 are not measurable without an importer, and measures the ledger for L13', () => {
    const rows = pptxScoreboard(files([['a', {}], ['b', {}]]));
    expect(rows.map((r) => r.id)).toEqual(['L11', 'L12', 'L13', 'L14']);
    expect(rows.map((r) => r.status)).toEqual(['not-measured', 'not-measured', 'not-met', 'not-measured']);
    expect(rows[2].measured).toBe('4 of 6 parts unaccounted in 2 of 2 files (0 mapped, 0 opaque, 2 dropped); nothing is mapped until an importer exists');
    for (const i of [0, 1, 3]) expect(rows[i].measured).toMatch(/^not measurable yet: no importer yet/);
  });

  it('shows whether the detector names the encrypted and legacy stubs', () => {
    const named = pptxScoreboard(
      files([
        ['edge-encrypted', { loss: [{ id: 'pptx.encrypted', label: 'e', count: 1, effect: 'x' }] }],
        ['edge-legacy-ppt', { loss: [] }],
      ])
    );
    expect(named[0].measured).toContain('encrypted: reported, legacy .ppt: not reported');
    expect(pptxScoreboard(files([['a', {}]]))[0].measured).toContain('encrypted: not reported, legacy .ppt: not reported');
  });

  it('meets L13 when nothing is unaccounted', () => {
    const clean = { mapped: 3, opaque: 0, dropped: 0, unaccounted: 0, total: 3, unaccountedParts: [] };
    expect(pptxScoreboard(files([['a', { ledger: clean }]]))[2].status).toBe('met');
  });
});

describe('toPptxReport', () => {
  it('assembles the format-neutral report', () => {
    const run: HarnessReport<PptxResult> = {
      fixtures: [outcome('a', result()), outcome('b', crashedResult('x'))],
      orphanGoldens: ['old.model.json'],
      ok: false,
    };
    const report = toPptxReport(run);
    expect(report).toMatchObject({
      format: 'PPTX',
      updateCommand: 'npm run slides:pptx-fidelity:update',
      ok: false,
      orphanGoldens: ['old.model.json'],
      kinds: ['model', 'loss', 'roundtrip', 'ledger'],
    });
    expect(report.files.map((f) => f.id)).toEqual(['a', 'b']);
    expect(report.scoreboard).toHaveLength(4);
  });
});
