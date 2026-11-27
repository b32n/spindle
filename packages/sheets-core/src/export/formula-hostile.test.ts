import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';

// ---------------------------------------------------------------------------
// What a formula in a file can cost on load: array formulas over huge extents (the file chooses the extent), deep nesting, and
// thousands of invented function names. Each import runs in a child process (`hostile/formula-probe.ts`) under a heap limit and a time
// limit, so "bounded" is a fact the process proves by finishing, not a reading of this process's clock or heap: before the array
// budget, 20,000 array formulas over blocks took 110 s, and a legacy array formula over the whole sheet would have asked for 17 billion
// cells. The limits are many times what a laptop needs.
// ---------------------------------------------------------------------------

const PROBE = resolve(__dirname, 'hostile/formula-probe.ts');

interface Out {
  ms?: number;
  zipBytes?: number;
  cells?: number;
  cachedOnly?: number;
  warnings?: string[];
  loss?: Array<[string, number]>;
  unsupportedNames?: number;
  error?: string;
}

function importInChild(shape: string, count: number, timeoutMs: number): { out: Out; status: number | null; signal: string | null; stderr: string } {
  const child = spawnSync(process.execPath, ['--max-old-space-size=1024', '--import', 'tsx', PROBE, shape, String(count)], { cwd: resolve(__dirname, '../..'), encoding: 'utf8', timeout: timeoutMs, maxBuffer: 16 * 1024 * 1024 });
  const line = (child.stdout ?? '').trim().split('\n').pop() as string;
  let out: Out = {};
  try {
    if (line.startsWith('{')) out = JSON.parse(line) as Out;
  } catch {
    // killed in the middle of the line: the status and signal say why
  }
  return { out, status: child.status, signal: child.signal, stderr: [child.error?.message ?? '', child.stderr ?? ''].join(' ').trim().slice(0, 2_000) };
}

const lossOf = (out: Out, id: string): number => (out.loss ?? []).filter(([i]) => i === id).reduce((n, [, c]) => n + c, 0);

describe('importing formulas of a hostile shape, in a child process with a heap and a time limit', () => {
  it.each(['arrays-legacy', 'arrays-dynamic'])('20,000 %s over 26,000 cells each: the budget keeps almost all as saved values, and says so', (shape) => {
    const { out, status, signal, stderr } = importInChild(shape, 20_000, 120_000);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out.error).toBeUndefined();
    expect(out.cells).toBe(20_000);
    expect(out.cachedOnly).toBeGreaterThan(19_000);
    expect(lossOf(out, 'xlsx.package.sizeLimit')).toBe(out.cachedOnly);
    expect(out.zipBytes).toBeLessThan(3 * 1024 * 1024);
  }, 150_000);

  it('an array formula over the whole sheet (17 billion cells) is kept as its saved value', () => {
    const { out, status, signal, stderr } = importInChild('array-huge', 1, 60_000);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out).toMatchObject({ cells: 1, cachedOnly: 1 });
    expect(lossOf(out, 'xlsx.package.sizeLimit')).toBe(1);
  }, 100_000);

  it.each(['nested', 'singles'])('2,000 formulas of 8,000 characters, shape %s: no stack overflow, no timeout', (shape) => {
    const { out, status, signal, stderr } = importInChild(shape, 2_000, 120_000);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out.error).toBeUndefined();
    expect(out.cells).toBe(2_000);
  }, 150_000);

  it('500 formulas of 450 invented function names each: every one is kept as its saved value, and only a few names are listed', () => {
    const { out, status, signal, stderr } = importInChild('names', 500, 120_000);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out).toMatchObject({ cells: 500, cachedOnly: 500 });
    expect(lossOf(out, 'xlsx.formulas.cachedOnly')).toBe(500);
    expect(out.unsupportedNames).toBeLessThanOrEqual(5);
  }, 150_000);
});
