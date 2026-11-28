import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';

// ---------------------------------------------------------------------------
// What the text of a cached-only formula can cost on a structural edit: the shifter runs once per cached-only cell on every insert or
// delete of rows or columns, on text a file chose (up to 8,192 characters a formula). Each shape runs in a child process
// (`hostile/ref-shift-probe.ts`) under a heap limit and a time limit, so "bounded" is a fact the process proves by finishing, not a
// reading of this process's clock or heap. The limits are many times what a laptop needs: the work is one pass over each text (the
// per-character cost is the same at 2 KB and at 128 KB, measured for the PR), so a quadratic loop, over the cells or over a text,
// would not finish.
// ---------------------------------------------------------------------------

const PROBE = resolve(__dirname, 'hostile/ref-shift-probe.ts');

interface Out {
  shape?: string;
  count?: number;
  insertMs?: number;
  deleteMs?: number;
  kept?: number;
  restored?: number;
  length?: number;
}

function runInChild(shape: string, count: number, timeoutMs: number, heapMb = 512): { out: Out; status: number | null; signal: string | null; stderr: string } {
  const child = spawnSync(process.execPath, [`--max-old-space-size=${heapMb}`, '--import', 'tsx', PROBE, shape, String(count)], { cwd: resolve(__dirname, '../..'), encoding: 'utf8', timeout: timeoutMs, maxBuffer: 16 * 1024 * 1024 });
  const line = (child.stdout ?? '').trim().split('\n').pop() as string;
  let out: Out = {};
  try {
    if (line.startsWith('{')) out = JSON.parse(line) as Out;
  } catch {
    // killed in the middle of the line: the status and signal say why
  }
  return { out, status: child.status, signal: child.signal, stderr: [child.error?.message ?? '', child.stderr ?? ''].join(' ').trim().slice(0, 2_000) };
}

describe('shifting the references of many cached-only cells, in a child process with a heap and a time limit', () => {
  it('10,000 cached-only cells on one row insert and one row delete: every text is shifted and restored, none loses its mark', () => {
    const { out, status, signal, stderr } = runInChild('typical', 10_000, 120_000);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out).toMatchObject({ count: 10_000, kept: 10_000, restored: 10_000 });
  }, 150_000);

  it.each([
    'words', 'words-bang', 'words-3d', 'refs', 'ranges', 'prefixed', 'brackets', 'brackets-nested', 'quotes', 'strings', 'garbage', 'colons', 'spills', 'dollars',
  ])('500 cells of 8 KB of %s text', (shape) => {
    const { out, status, signal, stderr } = runInChild(shape, 500, 120_000);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out.length).toBeLessThanOrEqual(8_192);
    expect(out).toMatchObject({ count: 500, kept: 500, restored: 500 });
  }, 150_000);
});
