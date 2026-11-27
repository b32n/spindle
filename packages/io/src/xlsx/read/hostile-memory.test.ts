import { buildShape, type ShapeName } from './hostile/shapes';
import { describeHeavy, importInChild, MB, SURVIVED, survived, type ChildOptions, type ChildRun, type Probe } from './hostile/child';

// ---------------------------------------------------------------------------
// Memory is a fact a process proves by surviving: each hostile shape found by review (a part of millions of tiny items, entity-dense
// text, rows without cells, a name or value of megabytes) is built as the smallest zip that makes it, written to a file, and imported
// by `hostile/probe.ts` in a child process whose heap is limited with `--max-old-space-size` (`hostile/child.ts`). A reader that holds more than its
// limit dies with an out-of-memory error, and one that takes more than the (generous) time limit is killed: either fails the test. The
// limits are generous (the shape used to cost several gigabytes, or the process), and nothing is read off the test process's own heap,
// which depends on the Node version and on when the collector ran; the tests assert what the import gave. The old reader's numbers for the same files are in the PR body.
// ---------------------------------------------------------------------------

/** The child of a shape, and the size of its zip. */
async function build(shape: ShapeName, heapMb: number, timeoutMs: number, limits?: ChildOptions['limits'], gc = false): Promise<{ run: ChildRun; zipBytes: number }> {
  const bytes = await buildShape(shape);
  return { run: importInChild(bytes, { heapMb, timeoutMs, limits, gc }), zipBytes: bytes.length };
}

const sheetCount = (r: { probe: { sheets: unknown[] } }): number => r.probe.sheets.length;

type Shape = [shape: ShapeName, heapMb: number, maxZipMb: number, timeoutSeconds: number, check: (p: Probe) => void];

const over = (what: string) => (p: Probe) => {
  expect(p.warnings).toBe(1);
  expect(p.firstWarnings[0]).toMatch(new RegExp(`${what}.* has more than 1000000 (elements|entries) \\(the import limit maxPartElements\\)`));
  expect(p.warningChars).toBeLessThan(300);
  expect(p.limitLoss?.map((i) => [i.id, i.label])).toEqual([['xlsx.package.sizeLimit', 'part(s) over the element limit']]);
};

describeHeavy('hostile workbook-level parts: memory stays bounded by what is kept', () => {
  // (zip sizes are about twice what the deflate output was, so a different zlib does not decide a test; the time limits are 8 to 15 times what a laptop takes)
  const shapes: Shape[] = [
    ['si-5m-tiny', 512, 28, 120, (p) => expect(p.sharedStrings).toBe(5_000_000)],
    ['si-2m', 512, 12, 60, (p) => expect(p.sharedStrings).toBe(2_000_000)],
    ['runs-400k', 512, 2, 60, (p) => expect(p.sharedStrings).toBe(1)],
    ['border-1m', 768, 8, 120, over('xl/styles.xml')],
    ['fill-1m', 768, 8, 120, over('xl/styles.xml')],
    ['font-1m', 768, 8, 120, over('xl/styles.xml')],
    ['dxf-500k', 768, 4, 120, over('xl/styles.xml')],
    ['theme-1m', 768, 8, 120, over('xl/theme/theme1.xml')],
    ['sheet-rels-1m', 768, 12, 120, over('xl/_rels/workbook.xml.rels')],
    ['content-types-1m', 768, 8, 120, over('\\[Content_Types\\].xml')],
  ];

  it.each(shapes)('%s: imports within %i MB of heap, from a zip of at most %i MB', async (shape, heapMb, maxZipMb, timeoutSeconds, check) => {
    const { run, zipBytes } = await build(shape, heapMb, timeoutSeconds * 1000);
    expect(survived(run)).toEqual(SURVIVED);
    expect(zipBytes).toBeLessThan(maxZipMb * MB);
    expect(sheetCount(run)).toBe(1);
    check(run.probe);
  }, 300_000);

  it('defined-names-1m: a million defined names are read (the part is data, not a document over the element budget), within 768 MB', async () => {
    const { run, zipBytes } = await build('defined-names-1m', 768, 120_000);
    expect(survived(run)).toEqual(SURVIVED);
    expect(zipBytes).toBeLessThan(12 * MB);
    expect(run.probe.definedNames).toBe(1_000_000);
    expect(run.probe.warnings).toBe(0);
    expect(run.probe.limitLoss).toEqual([]);
  }, 300_000);

  it('defined-names-150k under a limit of 100,000: the names past it are counted, with one warning and one loss item, and the import goes on', async () => {
    const { run } = await build('defined-names-150k', 256, 60_000, { maxPartElements: 100_000 });
    expect(survived(run)).toEqual(SURVIVED);
    expect(run.probe.definedNames).toBe(100_000);
    expect(sheetCount(run)).toBe(1);
    expect(run.probe.firstWarnings).toEqual(['The workbook has more than 100000 defined names (the import limit maxPartElements); 50000 were not read.']);
    expect(run.probe.limitLoss).toMatchObject([{ id: 'xlsx.package.sizeLimit', label: 'defined name(s) not kept', count: 50_000, where: ['xl/workbook.xml'] }]);
  }, 120_000);
});

describeHeavy('hostile text: entities, rows, names and values', () => {
  it.each([
    ['entities-42mb', 512],
    ['entities-100mb', 512],
    ['entities-cell-100mb', 512],
  ] as Array<[ShapeName, number]>)('%s: a part of entity-dense text decodes within %i MB of heap (the whole-text replace needed gigabytes)', async (shape, heapMb) => {
    const { run, zipBytes } = await build(shape, heapMb, 300_000);
    expect(survived(run)).toEqual(SURVIVED);
    expect(zipBytes).toBeLessThan(4 * MB);
    expect(sheetCount(run)).toBe(1);
    expect(run.probe.warnings).toBe(0);
  }, 400_000);

  it('rows-8x1m: eight sheets of 1,048,576 cell-less rows with heights stay within 1 GB of heap, keep at most maxRows rows, and warn once per sheet', async () => {
    const { run } = await build('rows-8x1m', 1024, 300_000);
    expect(survived(run)).toEqual(SURVIVED);
    expect(sheetCount(run)).toBe(8);
    // the first two sheets' rows fit the budget (2 x 1,048,576), the other six are left out with one warning each
    expect(run.probe.warnings).toBe(6);
    expect(run.probe.firstWarnings[0]).toMatch(/^S3: the settings of 1048576 row\(s\) were left out \(the import limit maxRows\)\.$/);
    expect(run.probe.warningChars).toBeLessThan(1_000);
    expect(run.probe.limitLoss).toMatchObject([{ id: 'xlsx.package.sizeLimit', label: 'row(s) with settings left out', count: 6 * 1_048_576, where: ['S3', 'S4', 'S5', 'S6', 'S7'] }]);
  }, 400_000);

  it('name-10mb: a sheet name of 10 MB is cut and its warnings stay small', async () => {
    const { run } = await build('name-10mb', 256, 120_000);
    expect(survived(run)).toEqual(SURVIVED);
    expect(run.probe.warnings).toBe(2);
    expect(run.probe.warningChars).toBeLessThan(1_000);
    expect(run.probe.limitLoss).toMatchObject([{ label: 'sheet name(s) cut at 255 characters', count: 1 }]);
  }, 120_000);

  it('value-1m-digits: a value of a million digits is quoted at 64 characters', async () => {
    const { run } = await build('value-1m-digits', 256, 120_000);
    expect(survived(run)).toEqual(SURVIVED);
    expect(run.probe.warnings).toBe(2);
    expect(run.probe.warningChars).toBeLessThan(500);
  }, 120_000);
});
