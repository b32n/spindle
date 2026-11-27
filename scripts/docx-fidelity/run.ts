/**
 * The DOCX adapter for the generic harness: one file in, three measurements out (model, loss, roundtrip).
 * Every step catches its own failures, so a throwing importer is a recorded result, never a crash.
 */
import type { HarnessAdapter } from '../fidelity-core/harness';
import type { Json, LossSnapshot, RoundtripSnapshot } from '../fidelity-core/types';
import { diffModels } from './diff';
import { normalizeModel } from './normalize';
import { DOCX_KINDS, type DocxKind, type EditorApi, type FixtureResult, type ModelSnapshot } from './types';

const errorMessage = (e: unknown): string => (e instanceof Error ? e.message : String(e));

async function inspect(api: EditorApi, bytes: Uint8Array): Promise<LossSnapshot> {
  try {
    const report = await api.inspectDocx(bytes);
    return { format: report.format, items: [...report.items].sort((a, b) => a.id.localeCompare(b.id)) };
  } catch (e) {
    // inspectDocx is documented never to throw; if it does, that is a finding to record, not a crash.
    return { format: 'docx', items: [{ id: 'harness.inspectThrew', label: 'inspectDocx threw', count: 1, effect: errorMessage(e) }] };
  }
}

/** Run one file through inspect -> import -> export -> re-import. Never throws: failures become part of the result. */
export async function runFixture(api: EditorApi, bytes: Uint8Array): Promise<FixtureResult> {
  const loss = await inspect(api, bytes);

  const document = api.createDocument();
  let model: ModelSnapshot;
  try {
    const { warnings } = await api.importFromDocx(bytes, document);
    model = { importError: null, warnings, document: normalizeModel(document.getData()) };
  } catch (e) {
    model = { importError: errorMessage(e), warnings: [], document: null };
  }

  return { loss, model, roundtrip: model.document === null ? emptyRoundtrip('skipped-import-error') : await roundtrip(api, document, model.document) };
}

const emptyRoundtrip = (status: RoundtripSnapshot['status'], extra: Partial<RoundtripSnapshot> = {}): RoundtripSnapshot => ({
  status,
  exportWarnings: [],
  reimportWarnings: [],
  diff: null,
  ...extra,
});

async function roundtrip(api: EditorApi, document: ReturnType<EditorApi['createDocument']>, imported: Json): Promise<RoundtripSnapshot> {
  let exported: Awaited<ReturnType<EditorApi['exportToDocx']>>;
  try {
    exported = await api.exportToDocx(document);
  } catch (e) {
    return emptyRoundtrip('export-error', { error: errorMessage(e) });
  }
  const second = api.createDocument();
  try {
    const { warnings } = await api.importFromDocx(exported.bytes, second);
    return {
      status: 'ok',
      exportWarnings: exported.warnings,
      reimportWarnings: warnings,
      diff: diffModels(imported, normalizeModel(second.getData())),
    };
  } catch (e) {
    return emptyRoundtrip('reimport-error', { error: errorMessage(e), exportWarnings: exported.warnings });
  }
}

/** The result of a run that did not finish (timed out): every kind says so. */
export function crashedResult(message: string): FixtureResult {
  return {
    loss: { format: 'docx', items: [{ id: 'harness.crashed', label: 'the harness run did not finish', count: 1, effect: message }] },
    model: { importError: message, warnings: [], document: null },
    roundtrip: emptyRoundtrip('skipped-import-error'),
  };
}

export function docxAdapter(api: EditorApi): HarnessAdapter<FixtureResult> {
  return {
    kinds: DOCX_KINDS,
    run: (bytes) => runFixture(api, bytes),
    crashed: crashedResult,
    snapshot: (result, kind) => result[kind as DocxKind],
    importError: (result) => result.model.importError,
  };
}
