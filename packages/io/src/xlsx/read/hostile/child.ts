import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir, totalmem } from 'node:os';
import { join, resolve } from 'node:path';
import type { ImportLossItem } from '@b32nio/spindle-shared';
import type { ImportLimits } from '../../../common/safe-zip';

// ---------------------------------------------------------------------------
// Memory and time of a hostile package are facts a process proves by surviving, not numbers a test can read off its own heap (which
// depend on the Node version and on when the collector last ran): the shape is imported by `probe.ts` in a child process started with a
// limited heap (`--max-old-space-size`) and a time limit. An out-of-memory exit or the time limit is the failure; the tests assert only
// what the import gave (counts, warnings, loss items), which does not vary by machine.
// ---------------------------------------------------------------------------

const IO_DIR = resolve(__dirname, '../../../..');
const PROBE = resolve(__dirname, 'probe.ts');
export const MB = 1024 * 1024;

/**
 * Whether the heavy shapes run: they need a machine with room for several children of a gigabyte or two beside the test run itself, so
 * they are skipped below 6 GB of memory, or when `XLSX_SKIP_HOSTILE` is set (for a slow or small runner).
 */
export const HEAVY = totalmem() >= 6 * 1024 * MB && !process.env.XLSX_SKIP_HOSTILE;

/** `describe` for the heavy shapes: skipped where `HEAVY` is false. */
export const describeHeavy = HEAVY ? describe : describe.skip;

export interface SheetProbe {
  name: string;
  cells: number;
  merges: number;
  hyperlinks: number;
  columns: number;
  columnWidths: number;
  hiddenCols: number;
  rows: number;
  extLst: number;
  dimension?: string;
  formulaChars: number;
}

export interface Probe {
  ms: number;
  peakRssMb: number;
  retained?: { heapMb: number; externalMb: number };
  sheets: SheetProbe[];
  sharedStrings: number;
  definedNames: number;
  keptRelationships: number;
  warnings: number;
  warningChars: number;
  firstWarnings: string[];
  limitLoss?: ImportLossItem[];
  error?: string;
  errorCode?: string;
}

export interface ChildRun {
  probe: Probe;
  /** Exit status: 0, 1 (the import threw: `probe.error`), or `null` when it was killed (out of memory, or over the time limit: see `signal`). */
  status: number | null;
  signal: string | null;
  stderr: string;
}

const scratch = mkdtempSync(join(tmpdir(), 'spindle-hostile-'));
afterAll(() => rmSync(scratch, { recursive: true, force: true }));

let counter = 0;

export interface ChildOptions {
  /** Heap limit of the child, `--max-old-space-size`. */
  heapMb: number;
  /** Time limit: generous (eight times what a laptop takes or more), so only a hang or a blow-up in cost reaches it. */
  timeoutMs: number;
  limits?: Partial<ImportLimits>;
  /** Collect before reporting what the result keeps (`probe.retained`). */
  gc?: boolean;
}

/** Run `script` (a `.ts` file, by `tsx`) in a child process under `options` with `args`; its last line of output as JSON, and how it ended. */
export function runChild<T>(script: string, args: string[], options: ChildOptions): { out: T; status: number | null; signal: string | null; stderr: string } {
  const flags = [`--max-old-space-size=${options.heapMb}`, ...(options.gc ? ['--expose-gc'] : [])];
  const child = spawnSync(process.execPath, [...flags, '--import', 'tsx', script, ...args], { cwd: IO_DIR, encoding: 'utf8', timeout: options.timeoutMs, maxBuffer: 64 * MB });
  const line = child.stdout.trim().split('\n').pop() ?? '{}';
  return { out: JSON.parse(line.startsWith('{') ? line : '{}') as T, status: child.status, signal: child.signal, stderr: child.stderr.slice(0, 2_000) };
}

/** Import `bytes` in a child process under `options`. */
export function importInChild(bytes: Uint8Array, options: ChildOptions): ChildRun {
  const file = join(scratch, `${process.pid}-${counter++}.xlsx`);
  writeFileSync(file, bytes);
  const { out, status, signal, stderr } = runChild<Probe>(PROBE, options.limits ? [file, JSON.stringify(options.limits)] : [file], options);
  rmSync(file, { force: true });
  return { probe: out, status, signal, stderr };
}

/** What a test asserts first: the child ended by itself with status 0 (it was not killed for memory or time, and the import did not throw). */
export const survived = (run: Pick<ChildRun, 'status' | 'signal' | 'stderr'> & { probe?: Probe }): { status: number | null; signal: string | null; error?: string; stderr: string } => ({
  status: run.status,
  signal: run.signal,
  ...(run.probe?.error === undefined ? {} : { error: run.probe.error }),
  stderr: run.status === 0 ? '' : run.stderr,
});

export const SURVIVED = { status: 0, signal: null, stderr: '' };
