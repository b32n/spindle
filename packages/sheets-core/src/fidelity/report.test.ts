import type { FixtureOutcome, HarnessReport } from '../../../../scripts/fidelity-core/harness';
import type { ReportFile } from '../../../../scripts/fidelity-core/report';
import {
  FORMULA_PROBE_IDS,
  toReportFile,
  toXlsxReport,
  xlsxScoreboard,
} from '../../../../scripts/xlsx-fidelity/report';
import { crashedResult, type XlsxResult } from '../../../../scripts/xlsx-fidelity/run';

const result = (over: Partial<XlsxResult> = {}): XlsxResult => ({
  loss: { format: 'xlsx', items: [{ id: 'xlsx.charts', label: 'Charts', count: 1, effect: 'e' }] },
  model: { importError: null, warnings: ['w'], workbook: {} },
  roundtrip: { status: 'ok', exportWarnings: [], reimportWarnings: [], diff: null },
  ledger: {
    package: 'zip',
    importError: false,
    totals: { mapped: 1, opaque: 0, dropped: 1, unaccounted: 2, total: 4 },
    parts: [
      { name: 'a', status: 'mapped', rule: 'r' },
      { name: 'b', status: 'dropped', rule: 'r', loss: ['x'] },
      ...Array.from({ length: 14 }, (_, i) => ({
        name: `u${String(i).padStart(2, '0')}`,
        status: 'unaccounted' as const,
        rule: 'none',
      })),
    ],
  },
  values: {
    status: 'ok',
    formulaCells: 3,
    compared: 3,
    matched: 2,
    mismatched: 1,
    uncached: 0,
    volatile: 0,
    unreadableSheets: 0,
    mismatchClasses: { 'number -> text': 1 },
    firstMismatches: [{ sheet: 'S', cell: 'A1', expected: '1', actual: '2' }],
  },
  ...over,
});
const outcome = (
  id: string,
  r: XlsxResult,
  over: Partial<FixtureOutcome<XlsxResult>> = {}
): FixtureOutcome<XlsxResult> => ({
  entry: {
    id,
    filename: `${id}.xlsx`,
    title: `T ${id}`,
    source: { tool: 'tool', os: 'o' },
    license: 'l',
    constructs: [],
    notes: '',
  },
  result: r,
  status: { model: 'match', loss: 'match', roundtrip: 'match', ledger: 'match', values: 'match' },
  problems: [],
  importExpectation: 'as-expected',
  ...over,
});

describe('toReportFile', () => {
  it('maps a harness outcome onto the report file, listing at most twelve unaccounted parts', () => {
    const f = toReportFile(outcome('a', result()));
    expect(f).toMatchObject({
      id: 'a',
      title: 'T a',
      producer: 'tool',
      importError: null,
      warnings: ['w'],
      importExpectation: 'as-expected',
    });
    expect(f.loss).toHaveLength(1);
    expect(f.ledger).toMatchObject({ mapped: 1, unaccounted: 2, total: 4 });
    expect(f.ledger?.unaccountedParts).toHaveLength(12);
    expect(f.ledger?.unaccountedParts[0]).toBe('u00');
    expect(f.values).toEqual({
      status: 'ok',
      compared: 3,
      matched: 2,
      mismatched: 1,
      firstMismatches: [{ sheet: 'S', cell: 'A1', expected: '1', actual: '2' }],
    });
  });
});

describe('xlsxScoreboard', () => {
  const files = (list: Array<[string, Partial<ReportFile>]>): ReportFile[] =>
    list.map(([id, over]) => ({ ...toReportFile(outcome(id, result())), ...over }));
  const row = (rows: ReturnType<typeof xlsxScoreboard>, id: string) =>
    rows.find((r) => r.id === id);

  it('measures L4 (exceptions), L5 (values), L7 (round trip) and L8 (ledger), and says L6 is not measurable', () => {
    const rows = xlsxScoreboard(
      files([
        [
          'formulas-core',
          {
            values: {
              status: 'ok',
              compared: 159,
              matched: 109,
              mismatched: 50,
              firstMismatches: [],
            },
          },
        ],
        ['openpyxl-image', { importError: 'boom', importExpectation: 'unexpected-error' }],
        ['edge-x', { importExpectation: 'unexpected-success' }],
      ])
    );
    expect(rows.map((r) => r.id)).toEqual(['L4', 'L5', 'L6', 'L7', 'L8']);
    expect(row(rows, 'L4')).toMatchObject({
      status: 'not-met',
      measured: '2 of 3 files: openpyxl-image threw, edge-x imported but should fail',
    });
    expect(row(rows, 'L5')).toMatchObject({ status: 'not-met' });
    expect(row(rows, 'L5')?.measured).toBe(
      'probe 109 of 159 (68.6%); corpus 113 of 165 over 3 files (68.5%)'
    );
    expect(row(rows, 'L6')?.status).toBe('not-measured');
    expect(row(rows, 'L7')).toMatchObject({
      status: 'met',
      measured: '0 differences in 0 of 3 round-tripped files; 0 failed, 0 skipped (import error)',
    });
    expect(row(rows, 'L8')).toMatchObject({ status: 'not-met' });
    expect(row(rows, 'L8')?.measured).toBe(
      '6 of 12 parts unaccounted in 3 of 3 files (3 mapped, 0 opaque, 3 dropped)'
    );
  });

  it('marks rows met when the numbers reach the target', () => {
    const perfect = {
      status: 'ok',
      compared: 100,
      matched: 100,
      mismatched: 0,
      firstMismatches: [],
    };
    const clean = {
      mapped: 3,
      opaque: 0,
      dropped: 0,
      unaccounted: 0,
      total: 3,
      unaccountedParts: [],
    };
    const rows = xlsxScoreboard(
      files([
        ['formulas-core', { values: perfect, ledger: clean }],
        ['b', { values: perfect, ledger: clean }],
      ])
    );
    expect(rows.map((r) => r.status)).toEqual(['met', 'met', 'not-measured', 'met', 'met']);
    expect(row(rows, 'L4')?.measured).toBe('0 of 2 files');
  });

  it('fails L7 on a difference, a skipped file or a failed round trip', () => {
    const rt = (status: 'ok' | 'skipped-import-error' | 'export-error', total = 0) => ({
      roundtrip: {
        status,
        exportWarnings: [],
        reimportWarnings: [],
        diff:
          status === 'ok'
            ? {
                total,
                added: 0,
                removed: 0,
                changed: total,
                byField: {},
                before: {},
                after: {},
                entries: [],
                truncated: false,
              }
            : null,
      },
    });
    for (const over of [rt('ok', 1), rt('skipped-import-error'), rt('export-error')])
      expect(row(xlsxScoreboard(files([['a', over]])), 'L7')?.status).toBe('not-met');
  });

  it('says L5 cannot be measured without the probe values, and passes it when both ratios are high', () => {
    expect(FORMULA_PROBE_IDS).toEqual(['formulas-core']);
    expect(row(xlsxScoreboard(files([['other', {}]])), 'L5')).toMatchObject({
      status: 'not-measured',
      measured: 'no probe values compared',
    });
    const probe = { status: 'ok', compared: 100, matched: 99, mismatched: 1, firstMismatches: [] };
    expect(row(xlsxScoreboard(files([['formulas-core', { values: probe }]])), 'L5')?.status).toBe(
      'met'
    );
    const lowCorpus = files([
      ['formulas-core', { values: probe }],
      ['b', { values: { ...probe, matched: 10, mismatched: 90 } }],
    ]);
    expect(row(xlsxScoreboard(lowCorpus), 'L5')?.status).toBe('not-met');
  });
});

describe('toXlsxReport', () => {
  it('assembles the format-neutral report', () => {
    const run: HarnessReport<XlsxResult> = {
      fixtures: [outcome('a', result()), outcome('b', crashedResult('x'))],
      orphanGoldens: ['old.model.json'],
      ok: false,
    };
    const report = toXlsxReport(run);
    expect(report).toMatchObject({
      format: 'XLSX',
      updateCommand: 'npm run sheets:xlsx-fidelity:update',
      ok: false,
      orphanGoldens: ['old.model.json'],
      kinds: ['model', 'loss', 'roundtrip', 'ledger', 'values'],
    });
    expect(report.files.map((f) => f.id)).toEqual(['a', 'b']);
    expect(report.scoreboard).toHaveLength(5);
  });
});
