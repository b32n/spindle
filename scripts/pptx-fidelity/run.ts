/**
 * The PPTX adapter for the generic harness: one file in, four measurements out (model, loss, roundtrip, ledger). Every
 * step catches its own failures, so a throwing importer is a recorded result, never a crash.
 *
 * The reader exists (plan K9) and the writer does not (K12): `model` holds the imported presentation (or the import
 * error, for the files that are not presentations) with the text parity check against the source, and `roundtrip` stops at `no exporter` until the writer lands. Against
 * an API without `importPptx` the model records `importError: "no importer"`.
 */
import path from 'node:path';
import { diffValues } from '../fidelity-core/diff';
import type { HarnessAdapter } from '../fidelity-core/harness';
import { readParts } from '../fidelity-core/ooxml-package';
import type { Json, LossSnapshot, RoundtripSnapshot } from '../fidelity-core/types';
import { NO_IMPORTER, type PptxEditorApi } from './api';
import { buildLedger, type LedgerSnapshot } from './ledger';
import { countPresentation, diffField, normalizePresentation, scrubMessage } from './normalize';
import { textParity, type TextParity } from './parity';

export const PPTX_KINDS = ['model', 'loss', 'roundtrip', 'ledger'] as const;
export type PptxKind = (typeof PPTX_KINDS)[number];

export interface ModelSnapshot {
  importError: string | null;
  warnings: string[];
  presentation: Json | null;
  /** Source text runs against the model's (plan K9): present when a model was produced. */
  textParity?: TextParity;
}

export interface PptxResult {
  loss: LossSnapshot;
  model: ModelSnapshot;
  roundtrip: RoundtripSnapshot;
  ledger: LedgerSnapshot;
}

/** Repository root and working directory, removed from error messages so the goldens do not depend on the checkout. */
const PATH_ROOTS = [path.resolve(__dirname, '..', '..'), process.cwd()];
const errorMessage = (e: unknown): string =>
  scrubMessage(e instanceof Error ? e.message : String(e), PATH_ROOTS);

const FORMAT = 'pptx';

async function inspect(api: PptxEditorApi, bytes: Uint8Array): Promise<LossSnapshot> {
  try {
    const report = await api.inspectPptx(bytes);
    return {
      format: report.format,
      items: [...report.items].sort((a, b) => a.id.localeCompare(b.id)),
    };
  } catch (e) {
    // detectPptxLoss is documented never to throw; if it does, that is a finding to record, not a crash.
    return {
      format: FORMAT,
      items: [
        {
          id: 'harness.inspectThrew',
          label: 'inspectPptx threw',
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
  api: PptxEditorApi,
  presentation: unknown,
  imported: Json
): Promise<RoundtripSnapshot> {
  if (!api.exportPptx || !api.importPptx) return roundtripOf('export-error', { error: 'no exporter' });
  let exported: Uint8Array;
  try {
    exported = await api.exportPptx(presentation);
  } catch (e) {
    return roundtripOf('export-error', { error: errorMessage(e) });
  }
  try {
    const second = await api.importPptx(exported);
    return roundtripOf('ok', {
      reimportWarnings: second.warnings,
      diff: diffValues(imported, normalizePresentation(second.presentation), {
        counts: countPresentation,
        field: diffField,
      }),
    });
  } catch (e) {
    return roundtripOf('reimport-error', { error: errorMessage(e) });
  }
}

async function importModel(
  api: PptxEditorApi,
  bytes: Uint8Array
): Promise<{ model: ModelSnapshot; raw: unknown }> {
  if (!api.importPptx)
    return { model: { importError: NO_IMPORTER, warnings: [], presentation: null }, raw: null };
  try {
    const { warnings, presentation } = await api.importPptx(bytes);
    return {
      model: { importError: null, warnings, presentation: normalizePresentation(presentation) },
      raw: presentation,
    };
  } catch (e) {
    return {
      model: { importError: errorMessage(e), warnings: [], presentation: null },
      raw: null,
    };
  }
}

/** The package part names the importer stored in the opaque part store: every part reference of every opaque element, groups included. */
export function opaquePartNames(presentation: unknown): Set<string> {
  const names = new Set<string>();
  const walk = (elements: unknown): void => {
    if (!Array.isArray(elements)) return;
    for (const e of elements as Array<Record<string, unknown>>) {
      if (e.kind === 'opaque') {
        for (const p of e.parts as Array<{ partName?: string }>) if (p.partName) names.add(p.partName);
      } else if (e.kind === 'group') walk(e.children);
    }
  };
  for (const slide of ((presentation as { slides?: Array<{ elements?: unknown }> } | null)?.slides ?? [])) walk(slide.elements);
  return names;
}

/** Run one file through inspect -> import -> export -> re-import, plus the ledger. Never throws. */
export async function runPptxFixture(api: PptxEditorApi, bytes: Uint8Array): Promise<PptxResult> {
  const loss = await inspect(api, bytes);
  const { model, raw } = await importModel(api, bytes);
  const parts = await readParts(bytes);
  const opaque = opaquePartNames(raw);
  if (model.presentation !== null) model.textParity = textParity(parts, model.presentation);
  return {
    loss,
    model,
    roundtrip:
      model.presentation === null
        ? roundtripOf('skipped-import-error')
        : await roundtrip(api, raw, model.presentation),
    ledger: buildLedger({
      parts,
      lossIds: loss.items.map((i) => i.id),
      importFailed: model.importError !== null,
      isOpaque: (name) => opaque.has(name),
    }),
  };
}

/** The result of a run that did not finish (timed out): every kind says so. */
export function crashedResult(message: string): PptxResult {
  const scrubbed = scrubMessage(message, PATH_ROOTS);
  return {
    loss: {
      format: FORMAT,
      items: [
        {
          id: 'harness.crashed',
          label: 'the harness run did not finish',
          count: 1,
          effect: scrubbed,
        },
      ],
    },
    model: { importError: scrubbed, warnings: [], presentation: null },
    roundtrip: roundtripOf('skipped-import-error'),
    ledger: {
      package: 'unreadable',
      importError: true,
      totals: { mapped: 0, opaque: 0, dropped: 0, unaccounted: 0, total: 0 },
      parts: [],
    },
  };
}

export function pptxAdapter(api: PptxEditorApi): HarnessAdapter<PptxResult> {
  return {
    kinds: PPTX_KINDS,
    run: (bytes) => runPptxFixture(api, bytes),
    crashed: crashedResult,
    snapshot: (result, kind) => result[kind as PptxKind],
    importError: (result) => result.model.importError,
  };
}
