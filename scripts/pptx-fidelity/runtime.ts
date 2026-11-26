/**
 * The opt-in `--runtime` run: the files that are too large to commit (the 1000-slide deck) are written by the stdlib
 * generator into a temporary directory and scanned once, with timing and memory. They have no goldens (timings are not
 * reproducible), so this only prints what happened.
 */
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import type { CliArgs } from '../fidelity-core/args';
import type { CliIo } from '../fidelity-core/cli';
import { withTimeout } from '../fidelity-core/harness';
import type { ManifestEntry } from '../fidelity-core/types';
import { NO_IMPORTER, type PptxEditorApi } from './api';
import { scrubMessage } from './normalize';

export interface RuntimeRun {
  id: string;
  /** `no importer` today; the import error once an importer exists and fails. */
  importError: string | null;
  elapsedMs: number;
  /** Resident set size after the run, in MB (the process peak is higher; this is a sanity figure). */
  rssMb: number;
  lossTypes: number;
  lossCount: number;
}

/** Scan (and, once there is an importer, import) one file with a time limit; never throws. */
export async function runRuntimeFile(
  api: PptxEditorApi,
  id: string,
  bytes: Uint8Array,
  timeoutMs: number
): Promise<RuntimeRun> {
  const started = Date.now();
  let importError: string | null = api.importPptx ? null : NO_IMPORTER;
  let lossTypes = 0;
  let lossCount = 0;
  try {
    const report = await withTimeout(api.inspectPptx(bytes), timeoutMs, `scan of ${id}`);
    lossTypes = report.items.length;
    lossCount = report.items.reduce((n, i) => n + i.count, 0);
    if (api.importPptx) await withTimeout(api.importPptx(bytes), timeoutMs, `import of ${id}`);
  } catch (e) {
    importError = scrubMessage(e instanceof Error ? e.message : String(e), [process.cwd()]);
  }
  return {
    id,
    importError,
    elapsedMs: Date.now() - started,
    rssMb: Math.round(process.memoryUsage().rss / 1048576),
    lossTypes,
    lossCount,
  };
}

/** Write the run-time files with the generator script; throws with the script's output when it fails. */
export function generateRuntimeFiles(
  script: string,
  outDir: string,
  ids: string[],
  timeoutMs = 300_000
): void {
  const run = spawnSync('python3', [script, '--out', outDir, ...ids], {
    encoding: 'utf8',
    timeout: timeoutMs,
  });
  if (run.status !== 0)
    throw new Error(`generator failed (${String(run.status)}): ${run.stderr || String(run.error)}`);
}

export interface RuntimeDeps {
  api: PptxEditorApi;
  fixturesDir: string;
  generate?: (script: string, outDir: string, ids: string[]) => void;
}

/** `CliConfig.runRuntime` for PPTX. */
export async function runRuntime(
  deps: RuntimeDeps,
  entries: ManifestEntry[],
  args: CliArgs,
  io: CliIo
): Promise<void> {
  const dir = mkdtempSync(path.join(tmpdir(), 'pptx-fidelity-runtime-'));
  try {
    const script = path.join(
      deps.fixturesDir,
      entries.find((e) => e.generator)?.generator ?? 'generators/build_runtime.py'
    );
    (deps.generate ?? generateRuntimeFiles)(
      script,
      dir,
      entries.map((e) => e.id)
    );
    io.log('\nRun-time files (no goldens; sanity figures only):');
    for (const entry of entries) {
      const bytes = new Uint8Array(readFileSync(path.join(dir, entry.filename)));
      const r = await runRuntimeFile(deps.api, entry.id, bytes, args.timeoutSeconds * 1000);
      io.log(
        `     ${r.id.padEnd(22)} ${String(r.elapsedMs).padStart(7)} ms  rss ${String(r.rssMb).padStart(5)} MB  loss types ${r.lossTypes} (${r.lossCount} objects)  import: ${r.importError ?? 'ok'}`
      );
    }
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}
