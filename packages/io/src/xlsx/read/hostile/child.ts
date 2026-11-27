import { spawnSync, type SpawnSyncReturns } from 'node:child_process';
import { mkdirSync, mkdtempSync, readdirSync, readFileSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { availableParallelism, hostname, tmpdir, totalmem } from 'node:os';
import { join, resolve } from 'node:path';
import { randomBytes } from 'node:crypto';
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
export const DETECT_PROBE = resolve(__dirname, 'detect-probe.ts');
export const MB = 1024 * 1024;

/** What a process may use: the container's limit when it has one (`constrainedMemory`, a cgroup), else the host's memory (`totalmem` reports the host inside a limited container). */
export function memoryBudget(constrained: number | undefined = process.constrainedMemory?.(), total: number = totalmem()): number {
  // 0 means no limit, and some platforms answer a huge number for it
  return constrained !== undefined && constrained > 0 && constrained < total ? constrained : total;
}

/**
 * Whether the heavy shapes run: they need a machine with room for children of a gigabyte or two beside the test run itself, so they are
 * skipped below 6 GB of memory (the container's limit where there is one, and no more than what is free when the run starts, where
 * `availableMemory` exists), or when `XLSX_SKIP_HOSTILE` is set (for a slow or small runner).
 */
export const heavyAllowed = (budget: number, available: number | undefined, skip: string | undefined): boolean => budget >= 6 * 1024 * MB && (available === undefined || available >= 2 * 1024 * MB) && !skip;
export const HEAVY = heavyAllowed(memoryBudget(), process.availableMemory?.(), process.env.XLSX_SKIP_HOSTILE);

/** How many children run at once across all the test workers: one for each 3 GB of the budget, one to three (a child takes up to 1.3 GB, and the worker that built its zip holds a few hundred MB). */
export const slotCount = (budget: number, cpus: number = availableParallelism()): number => Math.max(1, Math.min(3, cpus, Math.floor(budget / (3 * 1024 * MB))));

const SLOT_PREFIX = 'spindle-hostile-slot-';
const sleep = (ms: number): void => void Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);

/** Whether the process `pid` exists (signal 0 sends nothing). Not a pid at all (empty, NaN, zero or negative: `kill(0, 0)` would answer for the process group) is not alive. */
export function alive(pid: number): boolean {
  if (!Number.isInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (e) {
    // EPERM: it exists, and is somebody else's
    return (e as NodeJS.ErrnoException).code === 'EPERM';
  }
}

/** What the scratch directories and the slots of this file are named (`spindle-hostile-` and what follows), and how old one must be to be swept. */
const OWN = /^spindle-hostile-/;
const STALE_MS = 24 * 60 * 60 * 1000;
const swept = new Set<string>();

/** Remove the directories of this file's own names (`spindle-hostile-*`) in `root` that were last changed over `maxAgeMs` ago: what a killed run leaves (its 55 MB scratch directory, a slot). Once for each `root` in a process. */
export function sweepStale(root: string, maxAgeMs: number = STALE_MS, now: number = Date.now()): void {
  if (swept.has(root)) return;
  swept.add(root);
  for (const name of readdirSync(root)) {
    if (!OWN.test(name)) continue;
    try {
      if (now - statSync(join(root, name)).mtimeMs > maxAgeMs) rmSync(join(root, name), { recursive: true, force: true });
    } catch {
      // gone already, or somebody else's: leave it
    }
  }
}

/** Who a slot names: the process, the host (a container of its own has its own pids, and may share the temp directory) and a token of this process. */
const HOST = hostname();
const ME = `${process.pid}\n${HOST}\n${randomBytes(6).toString('hex')}`;
/** How long a slot that names nobody (or a garbage name) may take to be named before it is dead, and how long a slot of another host may stand before it is taken over (its pid says nothing here). */
const UNNAMED_MS = 10_000;
const FOREIGN_MS = 30 * 60 * 1000;

/** What the slot `dir` says (its `pid` file), or `undefined` when it has none (the holder is between creating the directory and naming itself, or the slot is gone). */
const readHolder = (dir: string): string | undefined => {
  try {
    return readFileSync(join(dir, 'pid'), 'utf8');
  } catch {
    return undefined;
  }
};

/** Whether the holder `seen` of `dir` is gone: by its pid on this host, else (another host, no name, a name that is no name) when the slot is older than that kind of slot may be. */
function holderGone(dir: string, seen: string | undefined, now: number): boolean {
  const [pid, host] = (seen ?? '').split('\n');
  const named = /^\d+$/.test(pid) && host !== undefined && host !== '';
  if (named && host === HOST) return !alive(Number(pid));
  return now - statSync(dir).mtimeMs > (named ? FOREIGN_MS : UNNAMED_MS);
}

/**
 * Put the slot `claimed` (taken from `dir` by a rename) back, when it turned out to be a live holder's: `dir` is made afresh (it fails when
 * another has made it since) and the `pid` file moved into it, so no path that may be another's is renamed over or removed; `claimed` is
 * ours alone. When `dir` is taken the claimed copy is dropped, and its holder (who frees only a slot that names it) goes on unprotected.
 */
function putBack(claimed: string, dir: string): void {
  try {
    mkdirSync(dir);
  } catch {
    return;
  }
  try {
    renameSync(join(claimed, 'pid'), join(dir, 'pid'));
  } catch {
    // nothing to put (the made directory is ours, and left to the next try to find unnamed)
  }
}

/** Take over the slot `dir` when its holder is gone: by renaming it to a name of the taker's own (one waiter wins the rename), and removing it only when it still says what was seen. */
function reap(dir: string): void {
  let seen: string | undefined;
  try {
    seen = readHolder(dir);
    if (!holderGone(dir, seen, Date.now())) return;
  } catch {
    // the slot is gone already
    return;
  }
  const claimed = `${dir}.reap-${randomBytes(6).toString('hex')}`;
  try {
    renameSync(dir, claimed);
  } catch {
    // another waiter won the rename
    return;
  }
  // a slot another has taken since is put back, not removed
  if (readHolder(claimed) !== seen) putBack(claimed, dir);
  rmSync(claimed, { recursive: true, force: true });
}

/**
 * Run `fn` holding one of `slots` machine-wide slots (a directory each, in the temp directory, that names the holder): the jest workers
 * of one run, or of two, do not each start a child of a gigabyte at the same time. A slot whose holder is gone (killed) is taken over
 * (`reap`), also one that was never named (killed between `mkdir` and the write, or the disk full: dead when over 10 s old). After
 * `waitMs` it runs anyway: a stuck neighbour must not fail the tests that wait for it. Directories of this file's names that a killed run
 * left over a day ago are removed first.
 */
export function withSlot<T>(fn: () => T, slots: number = slotCount(memoryBudget()), waitMs = 600_000, root: string = tmpdir()): T {
  sweepStale(root);
  const deadline = Date.now() + waitMs;
  let held: string | undefined;
  while (held === undefined && Date.now() < deadline) {
    for (let i = 0; i < slots && held === undefined; i++) {
      const dir = join(root, `${SLOT_PREFIX}${i}`);
      try {
        mkdirSync(dir);
      } catch {
        // taken: by a live process, or by one that is gone
        reap(dir);
        continue;
      }
      try {
        writeFileSync(join(dir, 'pid'), ME);
        held = dir;
      } catch {
        // cannot name itself (disk full): the directory is ours, and must not stand unnamed
        rmSync(dir, { recursive: true, force: true });
      }
    }
    if (held === undefined) sleep(50);
  }
  try {
    return fn();
  } finally {
    // only a slot that still names this process: after a takeover it is somebody else's
    if (held !== undefined && readHolder(held) === ME) rmSync(held, { recursive: true, force: true });
  }
}

/** `describe` for the heavy shapes: skipped where `heavy` (by default `HEAVY`) is false. */
export const describeIf = (heavy: boolean): jest.Describe => (heavy ? describe : describe.skip);
export const describeHeavy = describeIf(HEAVY);

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

export interface ChildOptions {
  /** Heap limit of the child, `--max-old-space-size`. */
  heapMb: number;
  /** Time limit: generous (eight times what a laptop takes or more), so only a hang or a blow-up in cost reaches it. */
  timeoutMs: number;
  limits?: Partial<ImportLimits>;
  /** Collect before reporting what the result keeps (`probe.retained`). */
  gc?: boolean;
  /** The script that reads the package and reports as `probe.ts` does (the default); `detect-probe.ts` runs the loss detector instead. */
  script?: string;
}

/** Run `script` (a `.ts` file, by `tsx`) in a child process under `options` with `args`; its last line of output as JSON, and how it ended. */
export function runChild<T>(script: string, args: string[], options: ChildOptions): { out: T; status: number | null; signal: string | null; stderr: string } {
  const flags = [`--max-old-space-size=${options.heapMb}`, ...(options.gc ? ['--expose-gc'] : [])];
  const child = withSlot(() => spawnSync(process.execPath, [...flags, '--import', 'tsx', script, ...args], { cwd: IO_DIR, encoding: 'utf8', timeout: options.timeoutMs, maxBuffer: 64 * MB }));
  return endOf<T>(child);
}

/** What a finished (or killed, or never started: `stdout` and `stderr` are then null) child gave. */
export function endOf<T>(child: Pick<SpawnSyncReturns<string>, 'stdout' | 'stderr' | 'status' | 'signal' | 'error'>): { out: T; status: number | null; signal: string | null; stderr: string } {
  return { out: lastJson<T>(child.stdout ?? ''), status: child.status, signal: child.signal, stderr: [child.error?.message ?? '', child.stderr ?? ''].join(' ').trim().slice(0, 2_000) };
}

/** The last line of `stdout` as JSON: `{}` when there is none or it is not (a child killed in the middle of writing it leaves a cut line, and the test then fails on `status` and `signal`, which say why). */
export function lastJson<T>(stdout: string): T {
  const line = stdout.trim().split('\n').pop() as string;
  try {
    return (line.startsWith('{') ? JSON.parse(line) : {}) as T;
  } catch {
    return {} as T;
  }
}

/** Import `bytes` in a child process under `options`. */
export function importInChild(bytes: Uint8Array, options: ChildOptions): ChildRun {
  // a directory of its own for each import, removed whatever happens (no module-level state for a test run to leave behind)
  const dir = mkdtempSync(join(tmpdir(), 'spindle-hostile-'));
  try {
    const file = join(dir, 'package.xlsx');
    writeFileSync(file, bytes);
    const { out, status, signal, stderr } = runChild<Probe>(options.script ?? PROBE, options.limits ? [file, JSON.stringify(options.limits)] : [file], options);
    return { probe: out, status, signal, stderr };
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

/** What a test asserts first: the child ended by itself with status 0 (it was not killed for memory or time, and the import did not throw). */
export const survived = (run: Pick<ChildRun, 'status' | 'signal' | 'stderr'> & { probe?: Probe }): { status: number | null; signal: string | null; error?: string; stderr: string } => ({
  status: run.status,
  signal: run.signal,
  ...(run.probe?.error === undefined ? {} : { error: run.probe.error }),
  stderr: run.status === 0 ? '' : run.stderr,
});

export const SURVIVED = { status: 0, signal: null, stderr: '' };
