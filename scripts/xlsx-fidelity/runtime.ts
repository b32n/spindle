/**
 * The opt-in `--runtime` run: the files that are too large or too hostile to commit (zip bomb, huge dimension,
 * 100k x 10 perf) are written by the stdlib generator into a temporary directory and imported once, with timing and
 * memory. They have no goldens (sizes and timings are not reproducible), so this only prints what happened.
 */
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import type { CliArgs } from '../fidelity-core/args';
import type { CliIo } from '../fidelity-core/cli';
import { withTimeout } from '../fidelity-core/harness';
import type { ManifestEntry } from '../fidelity-core/types';
import type { XlsxEditorApi } from './api';
import { countWorkbook, normalizeWorkbook, scrubMessage } from './normalize';

export interface RuntimeRun {
  id: string;
  importError: string | null;
  sheets: number;
  cells: number;
  elapsedMs: number;
  /** Resident set size after the import, in MB (the process peak is higher; this is a sanity figure). */
  rssMb: number;
  lossTypes: number;
}

/** Import one file with a time limit; never throws. */
export async function runRuntimeFile(
  api: XlsxEditorApi,
  id: string,
  bytes: Uint8Array,
  timeoutMs: number
): Promise<RuntimeRun> {
  const started = Date.now();
  const roots = [process.cwd()];
  let importError: string | null = null;
  let sheets = 0;
  let cells = 0;
  let lossTypes = 0;
  try {
    const workbook = api.createWorkbook();
    const { loss } = await withTimeout(
      api.importFromXlsx(bytes, workbook),
      timeoutMs,
      `import of ${id}`
    );
    lossTypes = loss.items.length;
    const counts = countWorkbook(normalizeWorkbook(workbook.getData()));
    sheets = counts.sheets;
    cells = counts.cells;
  } catch (e) {
    importError = scrubMessage(e instanceof Error ? e.message : String(e), roots);
  }
  return {
    id,
    importError,
    sheets,
    cells,
    elapsedMs: Date.now() - started,
    rssMb: Math.round(process.memoryUsage().rss / 1048576),
    lossTypes,
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
  api: XlsxEditorApi;
  fixturesDir: string;
  generate?: (script: string, outDir: string, ids: string[]) => void;
}

/** `CliConfig.runRuntime` for XLSX. */
export async function runRuntime(
  deps: RuntimeDeps,
  entries: ManifestEntry[],
  args: CliArgs,
  io: CliIo
): Promise<void> {
  const dir = mkdtempSync(path.join(tmpdir(), 'xlsx-fidelity-runtime-'));
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
      const outcome = r.importError
        ? `import error: ${r.importError}`
        : `${r.sheets} sheets, ${r.cells} cells`;
      io.log(
        `     ${r.id.padEnd(22)} ${String(r.elapsedMs).padStart(7)} ms  rss ${String(r.rssMb).padStart(5)} MB  loss types ${r.lossTypes}  ${outcome}`
      );
    }
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}
