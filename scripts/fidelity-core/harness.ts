import { mkdirSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { loadManifest } from './manifest';
import { compareSnapshot, findOrphanGoldens, removeGoldens, snapshotPath } from './snapshot';
import {
  GOOD_STATUSES,
  type ImportExpectation,
  type Manifest,
  type ManifestEntry,
  type ManifestOptions,
  type SnapshotStatus,
} from './types';

/** What a format adapter supplies; the harness itself knows nothing about workbooks, documents or decks. */
export interface HarnessAdapter<R> {
  /** Snapshot kinds written per fixture (`<id>.<kind>.json`). */
  kinds: readonly string[];
  /** Run one file through every kind of measurement. Must never throw: failures become part of the result. */
  run(bytes: Uint8Array, entry: ManifestEntry): Promise<R>;
  /** A result for a run that did not finish (the harness timed out or the adapter threw after all). */
  crashed(message: string): R;
  /** The JSON written to / compared with the golden of one kind. */
  snapshot(result: R, kind: string): unknown;
  /** The import error message, or null when the import succeeded. */
  importError(result: R): string | null;
}

export interface FixtureOutcome<R> {
  entry: ManifestEntry;
  result: R;
  status: Record<string, SnapshotStatus>;
  /** Problems that are not snapshot mismatches (only with `strictImport`). */
  problems: string[];
  /** Whether the import ended the way `expectImportError` says. */
  importExpectation: ImportExpectation;
}

export interface HarnessReport<R> {
  fixtures: Array<FixtureOutcome<R>>;
  orphanGoldens: string[];
  /** True when every snapshot matched (or was just written in update mode) and there are no problems. */
  ok: boolean;
}

export interface HarnessOptions<R> {
  adapter: HarnessAdapter<R>;
  fixturesDir: string;
  manifestOptions: ManifestOptions;
  /** Rewrite snapshots that differ (and delete orphans) instead of failing. */
  update?: boolean;
  /** Restrict to these fixture ids. */
  only?: string[];
  /**
   * Treat an import that disagrees with `expectImportError` as a problem that fails the run (DOCX
   * does). XLSX records today's behaviour in the goldens and reports the disagreement on the
   * scoreboard instead, so a known importer bug does not turn CI red.
   */
  strictImport?: boolean;
  /** Per-file wall-clock limit. */
  timeoutMs?: number;
}

export const DEFAULT_TIMEOUT_MS = 60_000;

const errorMessage = (e: unknown): string => (e instanceof Error ? e.message : String(e));

/** Resolve with the promise's value, or reject when it does not settle within `ms` (the timer never keeps the process alive). */
export function withTimeout<T>(promise: Promise<T>, ms: number, what: string): Promise<T> {
  let timer: NodeJS.Timeout;
  const limit = new Promise<never>((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${what} did not finish within ${ms} ms`)), ms);
    timer.unref();
  });
  return Promise.race([promise, limit]).finally(() => clearTimeout(timer));
}

/** The manifest's committed fixtures, optionally narrowed to `only`; throws on an unknown id. */
export function selectFixtures(manifest: Manifest, only: string[] = []): ManifestEntry[] {
  const committed = manifest.fixtures.filter((f) => !f.generatedAtRuntime);
  if (only.length === 0) return committed;
  const unknown = only.filter((id) => !committed.some((f) => f.id === id));
  if (unknown.length)
    throw new Error(
      `unknown fixture id(s): ${unknown.join(', ')} (known: ${committed.map((f) => f.id).join(', ')})`
    );
  return committed.filter((f) => only.includes(f.id));
}

export function importExpectationOf(
  entry: ManifestEntry,
  importError: string | null
): ImportExpectation {
  const threw = importError !== null;
  if (entry.expectImportError === true) return threw ? 'as-expected' : 'unexpected-success';
  return threw ? 'unexpected-error' : 'as-expected';
}

/** Run one file under the time limit. Never throws. */
export async function runOne<R>(
  adapter: HarnessAdapter<R>,
  entry: ManifestEntry,
  bytes: Uint8Array,
  timeoutMs: number
): Promise<R> {
  try {
    return await withTimeout(adapter.run(bytes, entry), timeoutMs, `fixture ${entry.id}`);
  } catch (e) {
    return adapter.crashed(errorMessage(e));
  }
}

export function goldensDirOf(fixturesDir: string): string {
  return path.join(fixturesDir, 'goldens');
}

/** Run the manifest's committed fixtures and compare each snapshot with its golden. */
export async function runHarness<R>(options: HarnessOptions<R>): Promise<HarnessReport<R>> {
  const {
    adapter,
    fixturesDir,
    manifestOptions,
    update = false,
    only,
    strictImport = false,
    timeoutMs = DEFAULT_TIMEOUT_MS,
  } = options;
  const manifest = loadManifest(fixturesDir, manifestOptions);
  const entries = selectFixtures(manifest, only);
  const goldens = goldensDirOf(fixturesDir);
  if (update) mkdirSync(goldens, { recursive: true });

  const fixtures: Array<FixtureOutcome<R>> = [];
  for (const entry of entries) {
    const bytes = new Uint8Array(readFileSync(path.join(fixturesDir, entry.filename)));
    const result = await runOne(adapter, entry, bytes, timeoutMs);
    const status: Record<string, SnapshotStatus> = {};
    for (const kind of adapter.kinds)
      status[kind] = compareSnapshot(
        snapshotPath(goldens, entry.id, kind),
        adapter.snapshot(result, kind),
        update
      );
    const importError = adapter.importError(result);
    const importExpectation = importExpectationOf(entry, importError);
    const problems: string[] = [];
    if (strictImport && importExpectation === 'unexpected-success')
      problems.push('manifest says the import should throw, but it succeeded');
    if (strictImport && importExpectation === 'unexpected-error')
      problems.push(`the import threw unexpectedly: ${importError}`);
    fixtures.push({ entry, result, status, problems, importExpectation });
  }

  // Orphans are judged against the whole manifest, even when only a subset ran.
  const orphanGoldens = findOrphanGoldens(
    goldens,
    manifest.fixtures.map((f) => f.id),
    adapter.kinds
  );
  if (update && orphanGoldens.length) removeGoldens(goldens, orphanGoldens);

  const ok =
    fixtures.every(
      (f) =>
        f.problems.length === 0 && adapter.kinds.every((k) => GOOD_STATUSES.includes(f.status[k]))
    ) &&
    (update || orphanGoldens.length === 0);
  return { fixtures, orphanGoldens: update ? [] : orphanGoldens, ok };
}

/** What is wrong with an outcome, one line each (empty when fine). */
export function describeOutcome<R>(o: FixtureOutcome<R>, kinds: readonly string[]): string[] {
  const lines = kinds
    .filter((k) => o.status[k] === 'mismatch' || o.status[k] === 'missing')
    .map((k) => `${o.entry.id}: ${k} snapshot ${o.status[k]}`);
  return [...lines, ...o.problems.map((p) => `${o.entry.id}: ${p}`)];
}
