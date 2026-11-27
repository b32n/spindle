import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';

// ---------------------------------------------------------------------------
// Merged ranges are applied to the model after the reader gives them, and a file can ask for hundreds of thousands in a few MB. The
// import of such a file runs in a child process (`hostile/merge-probe.ts`) under a heap limit and a time limit, so "fast enough" is a
// fact the process proves by finishing, not a reading of this process's clock: one `mergeCells` per range was quadratic (10,000 merges
// took 12 s, 20,000 51 s, and 40,000 more than 170 s). The time limits are 15 to 30 times what a laptop takes.
// ---------------------------------------------------------------------------

const PROBE = resolve(__dirname, 'hostile/merge-probe.ts');

interface Out {
  ms?: number;
  zipBytes?: number;
  merges?: number;
  cells?: number;
  canUndo?: boolean;
  warnings?: string[];
  loss?: Array<[string, number]>;
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

describe('importing a file with very many merged ranges, in a child process with a time limit', () => {
  it.each([20_000, 200_000])('%i merges of 2 x 2 blocks: all applied, no history, no loss', (count) => {
    const { out, status, signal, stderr } = importInChild('blocks', count, count > 100_000 ? 120_000 : 60_000);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out).toMatchObject({ merges: count, canUndo: false, warnings: [], loss: [] });
    expect(out.zipBytes).toBeLessThan(3 * 1024 * 1024);
  }, 150_000);

  it('100,000 merges of blocks and one over the first two: only those two are lost, none of the rest (an earlier version left out everything past 20,000)', () => {
    const { out, status, signal, stderr } = importInChild('blocksplus', 100_000, 120_000);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out).toMatchObject({ merges: 100_000 - 2, canUndo: false, warnings: [], loss: [] });
  }, 150_000);

  it('200,000 merges that each overlap the next: all are decided (the last is the one that stays), with no warning and no loss', () => {
    const { out, status, signal, stderr } = importInChild('overlap', 200_000, 120_000);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out).toMatchObject({ merges: 1, warnings: [], loss: [] });
  }, 150_000);

  it('one merge of the whole sheet (17 billion cells) over 5,000 cells clears them by one pass over the cells, not over its area', () => {
    const { out, status, signal, stderr } = importInChild('wide', 5_000, 60_000);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out).toMatchObject({ merges: 1, cells: 1, canUndo: false });
  }, 100_000);
});
