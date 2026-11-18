import { mkdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { diffModels } from './diff';
import { discoverFixtures, loadManifest } from './manifest';
import { normalizeModel } from './normalize';
import { compareSnapshot, findOrphanGoldens, removeGoldens, snapshotPath } from './snapshot';
import {
  SNAPSHOT_KINDS,
  type EditorApi,
  type FixtureResult,
  type Json,
  type LossSnapshot,
  type ManifestEntry,
  type ModelSnapshot,
  type RoundtripSnapshot,
  type SnapshotKind,
  type SnapshotStatus,
} from './types';

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

export interface FixtureOutcome {
  entry: ManifestEntry;
  result: FixtureResult;
  status: Record<SnapshotKind, SnapshotStatus>;
  /** Problems that are not snapshot mismatches (e.g. `expectImportError` disagreeing with reality). */
  problems: string[];
}

export interface HarnessReport {
  fixtures: FixtureOutcome[];
  orphanGoldens: string[];
  /** True when every snapshot matched (or was just written in update mode) and there are no problems. */
  ok: boolean;
}

export interface HarnessOptions {
  api: EditorApi;
  fixturesDir: string;
  /** Rewrite snapshots that differ (and delete orphans) instead of failing. */
  update?: boolean;
  /** Restrict to these fixture ids. */
  only?: string[];
}

export function goldensDirOf(fixturesDir: string): string {
  return path.join(fixturesDir, 'goldens');
}

/** Run the manifest's fixtures and compare each snapshot with its golden. */
export async function runHarness(options: HarnessOptions): Promise<HarnessReport> {
  const { api, fixturesDir, update = false, only } = options;
  const entries = discoverFixtures(fixturesDir, only);
  const goldens = goldensDirOf(fixturesDir);
  if (update) mkdirSync(goldens, { recursive: true });

  const fixtures: FixtureOutcome[] = [];
  for (const entry of entries) {
    const bytes = new Uint8Array(readFileSync(path.join(fixturesDir, entry.filename)));
    const result = await runFixture(api, bytes);
    const status = {} as Record<SnapshotKind, SnapshotStatus>;
    for (const kind of SNAPSHOT_KINDS) status[kind] = compareSnapshot(snapshotPath(goldens, entry.id, kind), result[kind], update);
    const problems: string[] = [];
    const threw = result.model.importError !== null;
    if (entry.expectImportError && !threw) problems.push('manifest says importFromDocx should throw, but the import succeeded');
    if (!entry.expectImportError && threw) problems.push(`importFromDocx threw unexpectedly: ${result.model.importError}`);
    fixtures.push({ entry, result, status, problems });
  }

  // Orphans are judged against the whole manifest, even when only a subset ran.
  const orphanGoldens = findOrphanGoldens(goldens, loadManifest(fixturesDir).fixtures.map((f) => f.id));
  if (update && orphanGoldens.length) removeGoldens(goldens, orphanGoldens);

  const ok =
    fixtures.every((f) => f.problems.length === 0 && SNAPSHOT_KINDS.every((k) => ['match', 'created', 'updated'].includes(f.status[k]))) &&
    (update || orphanGoldens.length === 0);
  return { fixtures, orphanGoldens: update ? [] : orphanGoldens, ok };
}

/** One-line description of what is wrong with an outcome (empty when fine). */
export function describeOutcome(o: FixtureOutcome): string[] {
  const lines = SNAPSHOT_KINDS.filter((k) => o.status[k] === 'mismatch' || o.status[k] === 'missing').map(
    (k) => `${o.entry.id}: ${k} snapshot ${o.status[k]}`
  );
  return [...lines, ...o.problems.map((p) => `${o.entry.id}: ${p}`)];
}
