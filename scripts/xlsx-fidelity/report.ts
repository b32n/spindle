import type { HarnessReport } from '../fidelity-core/harness';
import type { FidelityReport, ReportFile } from '../fidelity-core/report';
import {
  exceptionStats,
  ledgerStats,
  percent,
  roundtripStats,
  scoreRow,
  unmeasuredRow,
  valuesStats,
  type ScoreRow,
} from '../fidelity-core/scoreboard';
import { XLSX_KINDS, type XlsxResult } from './run';

/** The fixture that holds the plan's 159 probe formulas (row L5's baseline). */
export const FORMULA_PROBE_IDS = ['formulas-core'];

const MAX_UNACCOUNTED_LISTED = 12;

export function toReportFile(o: HarnessReport<XlsxResult>['fixtures'][number]): ReportFile {
  const r = o.result;
  return {
    id: o.entry.id,
    title: o.entry.title,
    producer: o.entry.source.tool,
    snapshots: o.status,
    problems: o.problems,
    importError: r.model.importError,
    importExpectation: o.importExpectation,
    warnings: r.model.warnings,
    loss: r.loss.items,
    roundtrip: r.roundtrip,
    ledger: {
      ...r.ledger.totals,
      unaccountedParts: r.ledger.parts
        .filter((p) => p.status === 'unaccounted')
        .map((p) => p.name)
        .slice(0, MAX_UNACCOUNTED_LISTED),
    },
    values: {
      status: r.values.status,
      compared: r.values.compared,
      matched: r.values.matched,
      mismatched: r.values.mismatched,
      passthrough: r.values.passthrough,
      firstMismatches: r.values.firstMismatches,
    },
  };
}

/** The launch-bar rows (plan 7.3) an XLSX run can measure. L9 to L11 need other tools and are not listed. */
export function xlsxScoreboard(files: ReportFile[]): ScoreRow[] {
  const ex = exceptionStats(files);
  const rt = roundtripStats(files);
  const lg = ledgerStats(files);
  const all = valuesStats(files);
  const probe = valuesStats(files, FORMULA_PROBE_IDS);
  const area = 'Sheets';
  const bad = ex.unexpectedErrors.length + ex.unexpectedSuccesses.length;
  return [
    scoreRow(
      'L4',
      area,
      'Corpus and edge: 0 uncaught exceptions; every edge file ends as its manifest says',
      `${bad} of ${ex.total} files${bad ? `: ${[...ex.unexpectedErrors.map((i) => `${i} threw`), ...ex.unexpectedSuccesses.map((i) => `${i} imported but should fail`)].join(', ')}` : ''}`,
      '0',
      bad === 0
    ),
    probe.compared > 0
      ? scoreRow(
          'L5',
          area,
          "Formulas match the file's cached values (probe: the 159 formulas of formulas-core; all files below)",
          `probe ${probe.matched} of ${probe.compared} (${percent(probe.matched, probe.compared)}); corpus ${all.matched} of ${all.compared} over ${all.files} files (${percent(all.matched, all.compared)})`,
          '>= 98%',
          probe.matched / probe.compared >= 0.98 && all.matched / all.compared >= 0.98
        )
      : unmeasuredRow(
          'L5',
          area,
          "Formulas match the file's cached values",
          'no probe values compared',
          '>= 98%'
        ),
    unmeasuredRow(
      'L6',
      area,
      'Number-format display match, codes survive a round trip',
      'needs display strings (the number-format suite); not measurable from the values kind',
      '>= 98%; >= 95%'
    ),
    scoreRow(
      'L7',
      area,
      'Round trip: no differences on the corpus',
      `${rt.diffs} differences in ${rt.filesWithDiffs} of ${rt.roundTripped} round-tripped files; ${rt.failed} failed, ${rt.skipped} skipped (import error)`,
      '0',
      rt.diffs === 0 && rt.failed === 0 && rt.skipped === 0
    ),
    scoreRow(
      'L8',
      area,
      'Ledger: 0 unaccounted package parts on every file',
      `${lg.unaccounted} of ${lg.total} parts unaccounted in ${lg.filesWithUnaccounted} of ${lg.files} files (${lg.mapped} mapped, ${lg.opaque} opaque, ${lg.dropped} dropped)`,
      '0',
      lg.unaccounted === 0
    ),
  ];
}

export function toXlsxReport(run: HarnessReport<XlsxResult>): FidelityReport {
  const files = run.fixtures.map(toReportFile);
  return {
    format: 'XLSX',
    updateCommand: 'npm run sheets:xlsx-fidelity:update',
    ok: run.ok,
    kinds: [...XLSX_KINDS],
    orphanGoldens: run.orphanGoldens,
    files,
    scoreboard: xlsxScoreboard(files),
  };
}
