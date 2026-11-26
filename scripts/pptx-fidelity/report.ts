import type { HarnessReport } from '../fidelity-core/harness';
import type { FidelityReport, ReportFile } from '../fidelity-core/report';
import { ledgerStats, scoreRow, unmeasuredRow, type ScoreRow } from '../fidelity-core/scoreboard';
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

const NO_IMPORTER_YET = 'no importer yet (plan K9 to K15): every file records importError "no importer"';

/**
 * The launch-bar rows (plan 7.3) a PPTX run can speak to. L11 to L14 need the importer, so they say so; the ledger figure
 * under L13 is real today. L15 to L17 need other tools (PowerPoint, Google Slides, LibreOffice) and are not listed.
 */
export function pptxScoreboard(files: ReportFile[]): ScoreRow[] {
  const area = 'Slides';
  const lg = ledgerStats(files);
  const withId = (id: string) => files.find((f) => f.id === id);
  const flagged = (id: string, lossId: string) =>
    withId(id)?.loss.some((i) => i.id === lossId) ? 'reported' : 'not reported';
  return [
    unmeasuredRow(
      'L11',
      area,
      'Safety: bomb, entity, traversal, encrypted, legacy files rejected with a named message under the limits',
      `not measurable yet: ${NO_IMPORTER_YET}. The detector itself names the stubs (encrypted: ${flagged('edge-encrypted', 'pptx.encrypted')}, legacy .ppt: ${flagged('edge-legacy-ppt', 'pptx.legacyBinary')}) and is bounded by size, entry and depth limits`,
      'all rejected'
    ),
    unmeasuredRow(
      'L12',
      area,
      'Corpus and edge: 0 uncaught exceptions',
      `not measurable yet: ${NO_IMPORTER_YET}, so there is nothing that could throw`,
      '0'
    ),
    scoreRow(
      'L13',
      area,
      'Ledger: 0 unaccounted package parts on every file (text parity and element mapping need the importer: not measurable yet)',
      `${lg.unaccounted} of ${lg.total} parts unaccounted in ${lg.filesWithUnaccounted} of ${lg.files} files (${lg.mapped} mapped, ${lg.opaque} opaque, ${lg.dropped} dropped); nothing is mapped until an importer exists`,
      '0',
      lg.unaccounted === 0
    ),
    unmeasuredRow(
      'L14',
      area,
      'Round trip: counts, positions, text, theme equal; opaque parts hash-identical',
      `not measurable yet: ${NO_IMPORTER_YET}, so every round trip is skipped`,
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
