import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';

// A structural edit looks at every cell for formulas to retarget and cached-only text to shift, unless the sheet never held one
// (`CellMap` flags): so a big value-only sheet costs what it did before those scans existed (measured for the PR against master:
// ~300 ms to delete a row of a 1,000,000-cell sheet on both). Each case runs in a child process (`hostile/scan-skip-probe.ts`) with a heap
// and a time limit; the bounds are many times the measured figures, they catch a scan that became quadratic, not a few milliseconds.

const PROBE = resolve(__dirname, 'hostile/scan-skip-probe.ts');

function run(kind: string, count: number): { deleteRowMs: number; insertColMs: number } {
  const child = spawnSync(process.execPath, ['--max-old-space-size=3000', '--import', 'tsx', PROBE, kind, String(count)], { cwd: resolve(__dirname, '../..'), encoding: 'utf8', timeout: 120_000, maxBuffer: 1024 * 1024 });
  expect({ status: child.status, signal: child.signal, stderr: child.stderr }).toEqual({ status: 0, signal: null, stderr: '' });
  return JSON.parse(child.stdout.trim().split('\n').pop() as string) as { deleteRowMs: number; insertColMs: number };
}

describe('the cost of a structural edit by what the sheet holds, in a child process', () => {
  it.each([
    ['values', 1_000_000, 5_000],
    ['cachedOnly', 20_000, 5_000],
    ['formulas', 2_000, 20_000],
  ])('%s: %i cells, a row delete and a column insert finish within %i ms each', (kind, count, boundMs) => {
    const { deleteRowMs, insertColMs } = run(kind, count);
    expect(deleteRowMs).toBeLessThan(boundMs);
    expect(insertColMs).toBeLessThan(boundMs);
  }, 150_000);
});
