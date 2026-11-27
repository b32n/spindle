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
const EXPORT_PROBE = resolve(__dirname, 'hostile/export-probe.ts');

interface Out {
  ms?: number;
  exportMs?: number;
  exportBytes?: number;
  exportWarnings?: string[];
  zipBytes?: number;
  cells?: number;
  cachedOnly?: number;
  warnings?: string[];
  loss?: Array<[string, number]>;
  unsupportedNames?: number;
  error?: string;
}

function importInChild(shape: string, count: number, timeoutMs: number, maxArrayCells?: number, heapMb = 1024): { out: Out; status: number | null; signal: string | null; stderr: string } {
  const child = spawnSync(process.execPath, [`--max-old-space-size=${heapMb}`, '--import', 'tsx', PROBE, shape, String(count), ...(maxArrayCells === undefined ? [] : [String(maxArrayCells)])], { cwd: resolve(__dirname, '../..'), encoding: 'utf8', timeout: timeoutMs, maxBuffer: 16 * 1024 * 1024 });
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
    // the ones over the budget (sizeLimit), and the legacy ones inside it that overlap each other and keep their saved values for it (cachedOnly)
    expect(lossOf(out, 'xlsx.package.sizeLimit') + lossOf(out, 'xlsx.formulas.cachedOnly')).toBe(out.cachedOnly);
    expect(out.zipBytes).toBeLessThan(3 * 1024 * 1024);
  }, 150_000);

  it('an array formula over the whole sheet (17 billion cells) is kept as its saved value', () => {
    const { out, status, signal, stderr } = importInChild('array-huge', 1, 60_000);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out).toMatchObject({ cells: 1, cachedOnly: 1 });
    expect(lossOf(out, 'xlsx.package.sizeLimit')).toBe(1);
  }, 100_000);

  it('an array formula over 400,000 rows (ref A1:J400000, 4 million cells) is kept as its saved value, imported and exported', () => {
    const { out, status, signal, stderr } = importInChild('rows400k', 1, 60_000);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out).toMatchObject({ cells: 1, cachedOnly: 1, exportWarnings: [] });
    expect(lossOf(out, 'xlsx.package.sizeLimit')).toBe(1);
    expect(out.zipBytes).toBeLessThan(4 * 1024);
    // one cell written, not 4 million
    expect(out.exportBytes).toBeLessThan(10 * 1024);
  }, 100_000);

  it('an array formula over exactly a million cells is spilled in full, and written back in full, within the heap and time limits', () => {
    const { out, status, signal, stderr } = importInChild('cells1m', 1, 240_000);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out).toMatchObject({ cachedOnly: 0, exportWarnings: [] });
    expect(lossOf(out, 'xlsx.package.sizeLimit')).toBe(0);
  }, 300_000);

  it('a 2 million cell array formula a host lets through (maxArrayCells) is written back as its first 1,000,000 cells, and the export says so', () => {
    const { out, status, signal, stderr } = importInChild('cells2m', 1, 240_000, 2_000_000, 2048);
    expect({ status, signal, stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    expect(out.cachedOnly).toBe(0);
    expect(out.exportWarnings).toEqual(['S!A1: the array formula\'s result (200000 x 10 cells) is over 1000000 cells, so only its first 100000 x 10 are written to the file.']);
  }, 300_000);

  it('many array formula results just under the per-result cap cannot add up past the workbook budget, and the warnings are counted', () => {
    // 40 results of 24,999 cells under a cap of 25,000 each and a budget of 100,000 for all of them (a million cells typed in)
    const child = spawnSync(process.execPath, ['--max-old-space-size=1024', '--import', 'tsx', EXPORT_PROBE, '40', '24999', '25000', '100000'], { cwd: resolve(__dirname, '../..'), encoding: 'utf8', timeout: 120_000, maxBuffer: 16 * 1024 * 1024 });
    expect({ status: child.status, signal: child.signal, stderr: child.stderr }).toEqual({ status: 0, signal: null, stderr: '' });
    const out = JSON.parse(child.stdout.trim().split('\n').pop() as string) as { arrayCells: number; warnings: string[]; error?: string };
    expect(out.error).toBeUndefined();
    // 4 results in full (99,996 cells), the fifth cut to the 4 cells left, and each of the other 35 keeps its first cell
    expect(out.arrayCells).toBe(4 * 24_999 + 4 + 35);
    expect(out.warnings).toHaveLength(21);
    expect(out.warnings[0]).toContain('does not fit in the 4 cells left');
    expect(out.warnings[20]).toBe('16 more array formula results were cut in the same way.');
  }, 150_000);

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
