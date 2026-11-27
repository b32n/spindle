import type { HarnessReport } from '../fidelity-core/harness';
import type { FidelityReport, ReportFile } from '../fidelity-core/report';
import { exceptionStats, ledgerStats, scoreRow, unmeasuredRow, type ScoreRow } from '../fidelity-core/scoreboard';
import { PPTX_KINDS, type PptxResult } from './run';

const MAX_UNACCOUNTED_LISTED = 12;

export function toReportFile(o: HarnessReport<PptxResult>['fixtures'][number]): ReportFile {
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
  };
}

const NO_EXPORTER_YET = 'no exporter yet (plan K12): every round trip stops at "no exporter"';

/**
 * The launch-bar rows (plan 7.3) a PPTX run can speak to. L12 and L13 are real now that the reader exists (K9); L11 only
 * covers the stub files until the limit probes are added, and L14 needs the exporter (K12). L15 to L17 need other tools (PowerPoint, Google Slides, LibreOffice) and are not listed.
 */
export function pptxScoreboard(files: ReportFile[]): ScoreRow[] {
  const area = 'Slides';
  const lg = ledgerStats(files);
  const withId = (id: string) => files.find((f) => f.id === id);
  const flagged = (id: string, lossId: string) =>
    withId(id)?.loss.some((i) => i.id === lossId) ? 'reported' : 'not reported';
  const ex = exceptionStats(files);
  const bad = ex.unexpectedErrors.length + ex.unexpectedSuccesses.length;
  return [
    unmeasuredRow(
      'L11',
      area,
      'Safety: bomb, entity, traversal, encrypted, legacy files rejected with a named message under the limits',
      `only the stub files are covered so far (the reader rejects them with a named message; the detector reports encrypted: ${flagged('edge-encrypted', 'pptx.encrypted')}, legacy .ppt: ${flagged('edge-legacy-ppt', 'pptx.legacyBinary')}); the bomb, entity and traversal probes are not in the corpus`,
      'all rejected'
    ),
    scoreRow(
      'L12',
      area,
      'Corpus and edge: 0 uncaught exceptions; every edge file ends as its manifest says',
      `${bad} of ${ex.total} files${bad ? `: ${[...ex.unexpectedErrors.map((i) => `${i} threw`), ...ex.unexpectedSuccesses.map((i) => `${i} imported but should fail`)].join(', ')}` : ''}`,
      '0',
      bad === 0
    ),
    scoreRow(
      'L13',
      area,
      'Ledger: 0 unaccounted package parts on every file (text parity and element mapping arrive with the slide content reader: not measurable yet)',
      `${lg.unaccounted} of ${lg.total} parts unaccounted in ${lg.filesWithUnaccounted} of ${lg.files} files (${lg.mapped} mapped, ${lg.opaque} opaque, ${lg.dropped} dropped)`,
      '0',
      lg.unaccounted === 0
    ),
    unmeasuredRow(
      'L14',
      area,
      'Round trip: counts, positions, text, theme equal; opaque parts hash-identical',
      `not measurable yet: ${NO_EXPORTER_YET}`,
      '0 diffs'
    ),
  ];
}

export function toPptxReport(run: HarnessReport<PptxResult>): FidelityReport {
  const files = run.fixtures.map(toReportFile);
  return {
    format: 'PPTX',
    updateCommand: 'npm run slides:pptx-fidelity:update',
    ok: run.ok,
    kinds: [...PPTX_KINDS],
    orphanGoldens: run.orphanGoldens,
    files,
    scoreboard: pptxScoreboard(files),
  };
}
