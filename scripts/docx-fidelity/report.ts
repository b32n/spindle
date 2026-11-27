import type { HarnessReport } from '../fidelity-core/harness';
import type { FidelityReport } from '../fidelity-core/report';
import type { FixtureResult } from './types';
import { DOCX_KINDS } from './types';

/** Map a DOCX harness run onto the format-neutral report (no ledger, values or launch-bar rows for DOCX). */
export function toDocxReport(run: HarnessReport<FixtureResult>): FidelityReport {
  return {
    format: 'DOCX',
    updateCommand: 'npm run docs:docx-fidelity:update',
    ok: run.ok,
    kinds: [...DOCX_KINDS],
    orphanGoldens: run.orphanGoldens,
    scoreboard: [],
    files: run.fixtures.map((o) => ({
      id: o.entry.id,
      title: o.entry.title,
      producer: o.entry.source.tool,
      snapshots: o.status,
      problems: o.problems,
      importError: o.result.model.importError,
      importExpectation: o.importExpectation,
      warnings: o.result.model.warnings,
      loss: o.result.loss.items,
      roundtrip: o.result.roundtrip,
    })),
  };
}
