/**
 * The XLSX adapter for the generic harness: one file in, five measurements out (model, loss, roundtrip, ledger,
 * values). Every step catches its own failures, so a throwing importer is a recorded result, never a crash.
 */
import path from 'node:path';
import { diffValues } from '../fidelity-core/diff';
import type { HarnessAdapter } from '../fidelity-core/harness';
import { readParts } from '../fidelity-core/ooxml-package';
import type { Json, LossSnapshot, RoundtripSnapshot } from '../fidelity-core/types';
import type { WorkbookLike, XlsxEditorApi } from './api';
import { buildLedger, type LedgerSnapshot } from './ledger';
import {
  countWorkbook,
  diffField,
  normalizeWorkbook,
  resolvePools,
  scrubMessage,
} from './normalize';
import { compareValues, readFileFormulas, unreadableValues, type ValuesSnapshot } from './values';

export const XLSX_KINDS = ['model', 'loss', 'roundtrip', 'ledger', 'values'] as const;
export type XlsxKind = (typeof XLSX_KINDS)[number];

export interface ModelSnapshot {
  importError: string | null;
  warnings: string[];
  workbook: Json | null;
}

export interface XlsxResult {
  loss: LossSnapshot;
  model: ModelSnapshot;
  roundtrip: RoundtripSnapshot;
  ledger: LedgerSnapshot;
  values: ValuesSnapshot;
}

/** Repository root and working directory, removed from error messages so the goldens do not depend on the checkout. */
const PATH_ROOTS = [path.resolve(__dirname, '..', '..'), process.cwd()];
const errorMessage = (e: unknown): string =>
  scrubMessage(e instanceof Error ? e.message : String(e), PATH_ROOTS);

const FORMAT = 'xlsx';

async function inspect(api: XlsxEditorApi, bytes: Uint8Array): Promise<LossSnapshot> {
  try {
    const report = await api.inspectXlsx(bytes);
    return {
      format: report.format,
      items: [...report.items].sort((a, b) => a.id.localeCompare(b.id)),
    };
  } catch (e) {
    // inspectXlsx is documented never to throw; if it does, that is a finding to record, not a crash.
    return {
      format: FORMAT,
      items: [
        {
          id: 'harness.inspectThrew',
          label: 'inspectXlsx threw',
          count: 1,
          effect: errorMessage(e),
        },
      ],
    };
  }
}

const roundtripOf = (
  status: RoundtripSnapshot['status'],
  extra: Partial<RoundtripSnapshot> = {}
): RoundtripSnapshot => ({
  status,
  exportWarnings: [],
  reimportWarnings: [],
  diff: null,
  ...extra,
});

/** Import -> export -> import again, and the structured difference between the two imported models. */
async function roundtrip(
  api: XlsxEditorApi,
  workbook: WorkbookLike,
  imported: Json
): Promise<RoundtripSnapshot> {
  let exported: Uint8Array;
  try {
    exported = await api.exportToXlsx(workbook);
  } catch (e) {
    return roundtripOf('export-error', { error: errorMessage(e) });
  }
  const second = api.createWorkbook();
  try {
    const { warnings } = await api.importFromXlsx(exported, second);
    return roundtripOf('ok', {
      reimportWarnings: warnings,
      diff: diffValues(imported, normalizeWorkbook(second.getData()), {
        counts: countWorkbook,
        prepare: resolvePools,
        field: diffField,
      }),
    });
  } catch (e) {
    return roundtripOf('reimport-error', { error: errorMessage(e) });
  }
}

/** The normalised model of an import that recalculates on load; null when that import fails (the values step then reports it as skipped). */
async function recalculatedModel(api: XlsxEditorApi, bytes: Uint8Array): Promise<Json | null> {
  try {
    const workbook = api.createWorkbook();
    await api.importFromXlsx(bytes, workbook, { recalculateOnLoad: true });
    return normalizeWorkbook(workbook.getData());
  } catch {
    return null;
  }
}

/** Run one file through inspect -> import -> export -> re-import, plus the ledger and the values comparison. Never throws. */
export async function runXlsxFixture(api: XlsxEditorApi, bytes: Uint8Array): Promise<XlsxResult> {
  const loss = await inspect(api, bytes);

  const workbook = api.createWorkbook();
  let model: ModelSnapshot;
  try {
    const { warnings } = await api.importFromXlsx(bytes, workbook);
    model = { importError: null, warnings, workbook: normalizeWorkbook(workbook.getData()) };
  } catch (e) {
    model = { importError: errorMessage(e), warnings: [], workbook: null };
  }

  const parts = await readParts(bytes);
  const importFailed = model.importError !== null;
  const roundtripResult =
    model.workbook === null ? roundtripOf('skipped-import-error') : await roundtrip(api, workbook, model.workbook);
  // The values comparison measures the engine against the file's saved results, so it imports with every formula the engine
  // can evaluate recalculated (the default import keeps the saved results, which would compare them with themselves).
  const recalculated = importFailed ? null : await recalculatedModel(api, bytes);
  return {
    loss,
    model,
    roundtrip: roundtripResult,
    ledger: buildLedger({ parts, lossIds: loss.items.map((i) => i.id), importFailed }),
    values:
      parts.kind === 'zip'
        ? compareValues(readFileFormulas(parts), recalculated, importFailed)
        : unreadableValues(),
  };
}

/** The result of a run that did not finish (timed out): every kind says so. */
export function crashedResult(message: string): XlsxResult {
  return {
    loss: {
      format: FORMAT,
      items: [
        {
          id: 'harness.crashed',
          label: 'the harness run did not finish',
          count: 1,
          effect: scrubMessage(message, PATH_ROOTS),
        },
      ],
    },
    model: { importError: scrubMessage(message, PATH_ROOTS), warnings: [], workbook: null },
    roundtrip: roundtripOf('skipped-import-error'),
    ledger: {
      package: 'unreadable',
      importError: true,
      totals: { mapped: 0, opaque: 0, dropped: 0, unaccounted: 0, total: 0 },
      parts: [],
    },
    values: unreadableValues(),
  };
}

export function xlsxAdapter(api: XlsxEditorApi): HarnessAdapter<XlsxResult> {
  return {
    kinds: XLSX_KINDS,
    run: (bytes) => runXlsxFixture(api, bytes),
    crashed: crashedResult,
    snapshot: (result, kind) => result[kind as XlsxKind],
    importError: (result) => result.model.importError,
  };
}
