import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { buildShape, type ShapeName } from './hostile/shapes';

// ---------------------------------------------------------------------------
// Memory is a fact a process proves by surviving: each hostile shape found by review (a part of millions of tiny items, entity-dense
// text, rows without cells, a name or value of megabytes) is built as the smallest zip that makes it, written to a file, and imported
// by `hostile/probe.ts` in a child process whose heap is limited with `--max-old-space-size`. A reader that holds more than its
// limit dies with an out-of-memory error and the test fails. The limits are generous (the shape used to cost several gigabytes, or
// the process); the child's own peak (RSS) and time are asserted too. The old reader's numbers for the same files are in the PR body.
// ---------------------------------------------------------------------------

const IO_DIR = resolve(__dirname, '../../../..');
const PROBE = resolve(__dirname, 'hostile/probe.ts');
const MB = 1024 * 1024;

interface Probe {
  ms: number;
  peakRssMb: number;
  sheets: number;
  cells: number;
  sharedStrings: number;
  warnings: number;
  warningChars: number;
  firstWarnings: string[];
  error?: string;
}

const scratch = mkdtempSync(join(tmpdir(), 'spindle-hostile-'));
afterAll(() => rmSync(scratch, { recursive: true, force: true }));

/** Import the shape's package in a child process with `heapMb` of heap; the child's JSON answer, and how it ended. */
async function importInChild(shape: ShapeName, heapMb: number): Promise<{ probe: Probe; zipBytes: number; status: number | null; stderr: string }> {
  const bytes = await buildShape(shape);
  const file = join(scratch, `${shape}.xlsx`);
  writeFileSync(file, bytes);
  const child = spawnSync(process.execPath, [`--max-old-space-size=${heapMb}`, '--import', 'tsx', PROBE, file], { cwd: IO_DIR, encoding: 'utf8', timeout: 240_000, maxBuffer: 16 * MB });
  const line = child.stdout.trim().split('\n').pop() ?? '{}';
  return { probe: JSON.parse(line.startsWith('{') ? line : '{}') as Probe, zipBytes: bytes.length, status: child.status, stderr: child.stderr.slice(0, 2_000) };
}

type Shape = [shape: ShapeName, heapMb: number, maxZipMb: number, maxSeconds: number, check: (p: Probe) => void];

const over = (what: string) => (p: Probe) => {
  expect(p.warnings).toBe(1);
  expect(p.firstWarnings[0]).toMatch(new RegExp(`${what}.* has more than 1000000 (elements|entries) \\(the import limit maxPartElements\\)`));
  expect(p.warningChars).toBeLessThan(300);
};

describe('hostile workbook-level parts: memory stays bounded by what is kept', () => {
  const shapes: Shape[] = [
    ['si-5m-tiny', 512, 14, 60, (p) => expect(p.sharedStrings).toBe(5_000_000)],
    ['si-2m', 512, 6, 30, (p) => expect(p.sharedStrings).toBe(2_000_000)],
    ['runs-400k', 512, 1, 30, (p) => expect(p.sharedStrings).toBe(1)],
    ['border-1m', 768, 4, 60, over('xl/styles.xml')],
    ['fill-1m', 768, 4, 60, over('xl/styles.xml')],
    ['font-1m', 768, 4, 60, over('xl/styles.xml')],
    ['dxf-500k', 768, 2, 60, over('xl/styles.xml')],
    ['theme-1m', 768, 4, 60, over('xl/theme/theme1.xml')],
    ['sheet-rels-1m', 768, 6, 60, over('xl/_rels/workbook.xml.rels')],
    ['content-types-1m', 768, 4, 60, over('\\[Content_Types\\].xml')],
  ];

  it.each(shapes)('%s: imports within %i MB of heap, from a zip of at most %i MB', async (shape, heapMb, maxZipMb, maxSeconds, check) => {
    const { probe, zipBytes, status, stderr } = await importInChild(shape, heapMb);
    expect({ status, stderr: status === 0 ? '' : stderr }).toEqual({ status: 0, stderr: '' });
    expect(zipBytes).toBeLessThan(maxZipMb * MB);
    expect(probe.sheets).toBe(1);
    expect(probe.ms).toBeLessThan(maxSeconds * 1000);
    expect(probe.peakRssMb).toBeLessThan(heapMb + 600);
    check(probe);
  }, 300_000);

  it('defined-names-1m: a workbook part over the budget is refused as corrupt, with the limit named, within 768 MB', async () => {
    const { probe, zipBytes, status } = await importInChild('defined-names-1m', 768);
    expect(status).toBe(1);
    expect(zipBytes).toBeLessThan(6 * MB);
    expect(probe.error).toMatch(/ImportError: The workbook part of the file is damaged .*maxPartElements/);
  }, 300_000);
});

describe('hostile text: entities, rows, names and values', () => {
  it.each([
    ['entities-42mb', 512],
    ['entities-100mb', 512],
    ['entities-cell-100mb', 512],
  ] as Array<[ShapeName, number]>)('%s: a part of entity-dense text decodes within %i MB of heap (the whole-text replace needed gigabytes)', async (shape, heapMb) => {
    const { probe, zipBytes, status, stderr } = await importInChild(shape, heapMb);
    expect({ status, stderr: status === 0 ? '' : stderr }).toEqual({ status: 0, stderr: '' });
    expect(zipBytes).toBeLessThan(2 * MB);
    expect(probe.ms).toBeLessThan(60_000);
    expect(probe.sheets).toBe(1);
    expect(probe.warnings).toBe(0);
  }, 300_000);

  it('rows-8x1m: eight sheets of 1,048,576 cell-less rows with heights stay within 1 GB of heap, keep at most maxRows rows, and warn once per sheet', async () => {
    const { probe, status, stderr } = await importInChild('rows-8x1m', 1024);
    expect({ status, stderr: status === 0 ? '' : stderr }).toEqual({ status: 0, stderr: '' });
    expect(probe.sheets).toBe(8);
    expect(probe.ms).toBeLessThan(120_000);
    // the first two sheets' rows fit the budget (2 x 1,048,576), the other six are left out with one warning each
    expect(probe.warnings).toBe(6);
    expect(probe.firstWarnings[0]).toMatch(/^S3: the settings of 1048576 row\(s\) were left out \(the import limit maxRows\)\.$/);
    expect(probe.warningChars).toBeLessThan(1_000);
  }, 300_000);

  it('name-10mb: a sheet name of 10 MB is cut and its warnings stay small', async () => {
    const { probe, status } = await importInChild('name-10mb', 256);
    expect(status).toBe(0);
    expect(probe.warnings).toBe(2);
    expect(probe.warningChars).toBeLessThan(1_000);
  }, 120_000);

  it('value-1m-digits: a value of a million digits is quoted at 64 characters', async () => {
    const { probe, status } = await importInChild('value-1m-digits', 256);
    expect(status).toBe(0);
    expect(probe.warnings).toBe(2);
    expect(probe.warningChars).toBeLessThan(500);
  }, 120_000);
});
