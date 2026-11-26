import type { ReportFile } from '../../../../../scripts/fidelity-core/report';
import {
  exceptionStats,
  ledgerStats,
  percent,
  roundtripStats,
  scoreRow,
  unmeasuredRow,
  valuesStats,
} from '../../../../../scripts/fidelity-core/scoreboard';

const file = (id: string, extra: Partial<ReportFile> = {}): ReportFile => ({
  id,
  title: id,
  producer: 'p',
  snapshots: {},
  problems: [],
  importError: null,
  importExpectation: 'as-expected',
  warnings: [],
  loss: [],
  roundtrip: { status: 'ok', exportWarnings: [], reimportWarnings: [], diff: null },
  ...extra,
});
const diff = (total: number) => ({
  total,
  added: 0,
  removed: 0,
  changed: total,
  byField: {},
  before: {},
  after: {},
  entries: [],
  truncated: false,
});

describe('scoreboard helpers', () => {
  it('formats percentages and guards the empty case', () => {
    expect(percent(1, 3)).toBe('33.3%');
    expect(percent(0, 0)).toBe('n/a');
  });

  it('builds met, not-met and not-measured rows', () => {
    expect(scoreRow('L4', 'A', 'c', 'm', 't', true)).toEqual({
      id: 'L4',
      area: 'A',
      criterion: 'c',
      measured: 'm',
      target: 't',
      status: 'met',
    });
    expect(scoreRow('L4', 'A', 'c', 'm', 't', false).status).toBe('not-met');
    expect(unmeasuredRow('L6', 'A', 'c', 'why', 't')).toMatchObject({
      measured: 'why',
      status: 'not-measured',
    });
  });

  it('counts imports that did not end as the manifest says', () => {
    const stats = exceptionStats([
      file('a'),
      file('b', { importExpectation: 'unexpected-error' }),
      file('c', { importExpectation: 'unexpected-success' }),
    ]);
    expect(stats).toEqual({ total: 3, unexpectedErrors: ['b'], unexpectedSuccesses: ['c'] });
  });

  it('tallies round trips', () => {
    const stats = roundtripStats([
      file('clean', {
        roundtrip: { status: 'ok', exportWarnings: [], reimportWarnings: [], diff: diff(0) },
      }),
      file('dirty', {
        roundtrip: { status: 'ok', exportWarnings: [], reimportWarnings: [], diff: diff(7) },
      }),
      file('nodiff'),
      file('skipped', {
        roundtrip: {
          status: 'skipped-import-error',
          exportWarnings: [],
          reimportWarnings: [],
          diff: null,
        },
      }),
      file('exp', {
        roundtrip: { status: 'export-error', exportWarnings: [], reimportWarnings: [], diff: null },
      }),
      file('re', {
        roundtrip: {
          status: 'reimport-error',
          exportWarnings: [],
          reimportWarnings: [],
          diff: null,
        },
      }),
    ]);
    expect(stats).toEqual({
      total: 6,
      roundTripped: 3,
      filesWithDiffs: 1,
      diffs: 7,
      skipped: 1,
      failed: 2,
    });
  });

  it('tallies ledgers over the files that have one', () => {
    const ledger = (u: number) => ({
      mapped: 5,
      opaque: 1,
      dropped: 2,
      unaccounted: u,
      total: 8 + u,
      unaccountedParts: [],
    });
    expect(
      ledgerStats([file('a', { ledger: ledger(0) }), file('b', { ledger: ledger(3) }), file('c')])
    ).toEqual({
      files: 2,
      filesWithUnaccounted: 1,
      mapped: 10,
      opaque: 2,
      dropped: 4,
      unaccounted: 3,
      total: 19,
    });
    expect(ledgerStats([])).toMatchObject({ files: 0, unaccounted: 0 });
  });

  it('tallies values over all files or a chosen few', () => {
    const v = (compared: number, matched: number) => ({
      status: 'ok',
      compared,
      matched,
      mismatched: compared - matched,
      firstMismatches: [],
    });
    const files = [file('a', { values: v(10, 7) }), file('b', { values: v(4, 4) }), file('c')];
    expect(valuesStats(files)).toEqual({ files: 2, compared: 14, matched: 11, mismatched: 3 });
    expect(valuesStats(files, ['a'])).toEqual({
      files: 1,
      compared: 10,
      matched: 7,
      mismatched: 3,
    });
    expect(valuesStats(files, ['zzz'])).toEqual({
      files: 0,
      compared: 0,
      matched: 0,
      mismatched: 0,
    });
  });
});
