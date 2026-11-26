import type { ReportFile } from './report';

/** One row of the launch-bar scoreboard (plan section 7.3). */
export interface ScoreRow {
  /** The plan's row id, e.g. `L4`. */
  id: string;
  area: string;
  criterion: string;
  /** What this run measured, as text ("3 of 46 files", "104 of 159 (65.4%)") or why it cannot be measured here. */
  measured: string;
  target: string;
  status: 'met' | 'not-met' | 'not-measured';
}

export const percent = (part: number, whole: number): string =>
  whole === 0 ? 'n/a' : `${((part / whole) * 100).toFixed(1)}%`;

export function scoreRow(
  id: string,
  area: string,
  criterion: string,
  measured: string,
  target: string,
  met: boolean
): ScoreRow {
  return { id, area, criterion, measured, target, status: met ? 'met' : 'not-met' };
}

export function unmeasuredRow(
  id: string,
  area: string,
  criterion: string,
  why: string,
  target: string
): ScoreRow {
  return { id, area, criterion, measured: why, target, status: 'not-measured' };
}

/** Files whose import did not end as the manifest says (an unexpected throw is an uncaught exception in the importer). */
export function exceptionStats(files: ReportFile[]) {
  return {
    total: files.length,
    unexpectedErrors: files
      .filter((f) => f.importExpectation === 'unexpected-error')
      .map((f) => f.id),
    unexpectedSuccesses: files
      .filter((f) => f.importExpectation === 'unexpected-success')
      .map((f) => f.id),
  };
}

/** Round-trip tallies: files that round-tripped, how many differ, the difference total, and failures. */
export function roundtripStats(files: ReportFile[]) {
  const ok = files.filter((f) => f.roundtrip.status === 'ok');
  const diffTotals = ok.map((f) => f.roundtrip.diff?.total ?? 0);
  return {
    total: files.length,
    roundTripped: ok.length,
    filesWithDiffs: diffTotals.filter((n) => n > 0).length,
    diffs: diffTotals.reduce((a, b) => a + b, 0),
    skipped: files.filter((f) => f.roundtrip.status === 'skipped-import-error').length,
    failed: files.filter(
      (f) => f.roundtrip.status === 'export-error' || f.roundtrip.status === 'reimport-error'
    ).length,
  };
}

/** Ledger tallies over the files that have one. */
export function ledgerStats(files: ReportFile[]) {
  const ledgers = files.flatMap((f) => (f.ledger ? [f.ledger] : []));
  const sum = (key: 'mapped' | 'opaque' | 'dropped' | 'unaccounted' | 'total') =>
    ledgers.reduce((n, l) => n + l[key], 0);
  return {
    files: ledgers.length,
    filesWithUnaccounted: ledgers.filter((l) => l.unaccounted > 0).length,
    mapped: sum('mapped'),
    opaque: sum('opaque'),
    dropped: sum('dropped'),
    unaccounted: sum('unaccounted'),
    total: sum('total'),
  };
}

/** Values tallies, over every file or only the given fixture ids. */
export function valuesStats(files: ReportFile[], ids?: string[]) {
  const picked = files.flatMap((f) => (f.values && (!ids || ids.includes(f.id)) ? [f.values] : []));
  const sum = (key: 'compared' | 'matched' | 'mismatched') =>
    picked.reduce((n, v) => n + v[key], 0);
  return {
    files: picked.length,
    compared: sum('compared'),
    matched: sum('matched'),
    mismatched: sum('mismatched'),
  };
}
