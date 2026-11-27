import { resolve } from 'node:path';
import { buildShape, type ShapeName } from './hostile/shapes';
import { describeHeavy, importInChild, runChild, SURVIVED, survived, type ChildOptions, type ChildRun, type SheetProbe } from './hostile/child';

const FLAT_PROBE = resolve(__dirname, 'hostile/flat-probe.ts');

// ---------------------------------------------------------------------------
// The shapes of review round 5, imported in a child process with a limited heap (see `hostile-memory.test.ts`; its own file so that the
// two run in parallel): a long number format code under every cell format, strings the result keeps that held the part they were cut from
// alive, and the package-wide budgets (columns, extension lists, relationships, cells, scanned elements). The tests assert what the import
// gave and what a collected child still holds; the time limits are 8 to 15 times what a laptop takes.
// ---------------------------------------------------------------------------

/** The child of a shape, and the size of its zip. */
async function build(shape: ShapeName, heapMb: number, timeoutMs: number, limits?: ChildOptions['limits'], gc = false): Promise<{ run: ChildRun; zipBytes: number }> {
  const bytes = await buildShape(shape);
  return { run: importInChild(bytes, { heapMb, timeoutMs, limits, gc }), zipBytes: bytes.length };
}

const sheetCount = (r: { probe: { sheets: unknown[] } }): number => r.probe.sheets.length;
const sum = (r: { probe: { sheets: SheetProbe[] } }, key: keyof SheetProbe): number => r.probe.sheets.reduce((n, s) => n + (s[key] as number), 0);

describeHeavy('hostile number formats: a long code under every cell format (review round 5)', () => {
  it.each([
    ['numfmt-1mib-2000xf', 2000],
    ['numfmt-4mib-900xf', 900],
  ] as Array<[ShapeName, number]>)('%s: the code is not scanned or copied per format: a few hundred milliseconds in 256 MB (it was 165 s and 3.6 GB, and more than 300 s)', async (shape, cells) => {
    const { run, zipBytes } = await build(shape, 256, 60_000);
    expect(survived(run)).toEqual(SURVIVED);
    expect(zipBytes).toBeLessThan(100_000);
    expect(run.probe.sheets[0].cells).toBe(cells);
    // one warning however many formats use the code, and one loss item
    expect(run.probe.firstWarnings).toEqual(['A number format code is over 255 characters (Excel allows no more); it is kept as written and not read as a date or time.']);
    expect(run.probe.limitLoss).toMatchObject([{ id: 'xlsx.package.sizeLimit', label: 'number format code(s) over 255 characters', count: 1, state: 'approximated' }]);
  }, 120_000);
});

describeHeavy('strings the result keeps do not keep the part they came from alive (review round 5)', () => {
  it('a view of a large string keeps it alive in the engine; `flat` copies, whether the text is one byte or two a character', () => {
    const held = (mode: string, width: string) => runChild<{ heapMb: number; externalMb: number; length: number }>(FLAT_PROBE, [mode, width], { heapMb: 512, timeoutMs: 60_000, gc: true });
    for (const width of ['narrow', 'wide']) {
      const view = held('view', width);
      const copy = held('flat', width);
      expect([view.status, copy.status]).toEqual([0, 0]);
      expect(copy.out.length).toBe(13);
      // the control: without the copy the 13 characters hold the 100 MB text (200 MB when it is two bytes a character)
      expect(view.out.externalMb + view.out.heapMb).toBeGreaterThan(90);
      expect(copy.out.externalMb + copy.out.heapMb).toBeLessThan(25);
    }
  }, 120_000);

  it('dimension-120mb: a 13-character dimension of a 120 MB part holds nothing of the part once the import is done (it kept all of it)', async () => {
    const { run } = await build('dimension-120mb', 512, 120_000, undefined, true);
    expect(survived(run)).toEqual(SURVIVED);
    expect(run.probe.sheets[0].dimension).toBe('A1:XFD1048576');
    expect((run.probe.retained?.heapMb ?? 1e6) + (run.probe.retained?.externalMb ?? 1e6)).toBeLessThan(60);
  }, 240_000);

  it('sheets-500x1mb: 500 sheets of a megabyte with a dimension each (two bytes a character) hold a few megabytes within a 512 MB heap (it was 1 GB for 500 sheets, 2.15 GB for 1,000)', async () => {
    const { run } = await build('sheets-500x1mb', 512, 120_000, undefined, true);
    expect(survived(run)).toEqual(SURVIVED);
    expect(sheetCount(run)).toBe(500);
    expect(run.probe.sheets.every((s) => s.dimension === 'A1:XFD1048576' && s.cells === 1)).toBe(true);
    expect((run.probe.retained?.heapMb ?? 1e6) + (run.probe.retained?.externalMb ?? 1e6)).toBeLessThan(100);
  }, 300_000);
});

describeHeavy('package-wide budgets (review round 5)', () => {
  it('cols-300x16k: 300 sheets of 16,384 <col> are drawn from one budget (maxColumns), with one warning and one loss item (900 of them took 24 s and 1.67 GB, 1.29 GB of it live)', async () => {
    const { run } = await build('cols-300x16k', 768, 300_000);
    expect(survived(run)).toEqual(SURVIVED);
    expect(sheetCount(run)).toBe(300);
    // 16,385 columns (the ranges and what they set on a sheet with a cell in XFD) are what a sheet costs: 64 sheets fit the budget of 1,048,576
    expect(sum(run, 'columnWidths')).toBeLessThanOrEqual(1_048_576);
    expect(sum(run, 'columns')).toBeLessThanOrEqual(1_048_576);
    expect(run.probe.sheets[0].columnWidths).toBe(16_384);
    expect(run.probe.sheets[299].columnWidths).toBe(0);
    expect(run.probe.firstWarnings.filter((w) => /column/.test(w))).toHaveLength(1);
    expect(run.probe.limitLoss).toMatchObject([{ id: 'xlsx.package.sizeLimit', label: 'column setting(s) left out' }]);
  }, 400_000);

  it('extlst-12x8mib: twelve extension lists of 8 MiB keep 64 MiB (maxExtensionChars) in all, with one warning and one loss item', async () => {
    const { run } = await build('extlst-12x8mib', 768, 120_000);
    expect(survived(run)).toEqual(SURVIVED);
    expect(sum(run, 'extLst')).toBeLessThanOrEqual(64 * 1024 * 1024);
    expect(run.probe.sheets.filter((s) => s.extLst > 0)).toHaveLength(8);
    expect(run.probe.firstWarnings).toHaveLength(1);
    expect(run.probe.limitLoss).toMatchObject([{ label: 'extension list(s) not kept', count: 4, where: ['S9', 'S10', 'S11', 'S12'] }]);
  }, 240_000);

  it('rels-4x400k-other: relationships of kinds the reader does not use are counted against maxRelationships and not kept (28 sheets of a million kept 3.4 GB)', async () => {
    const { run } = await build('rels-4x400k-other', 768, 120_000);
    expect(survived(run)).toEqual(SURVIVED);
    expect(run.probe.keptRelationships).toBe(0);
    expect(run.probe.firstWarnings).toEqual(['The package has more than 1048576 relationships (the import limit maxRelationships); the rest, from xl/worksheets/_rels/sheet3.xml.rels on, were not read.']);
    expect(run.probe.limitLoss).toMatchObject([{ label: 'part(s) with relationships not read', count: 2 }]);
  }, 240_000);

  it('rels-3x400k-links: the relationships that are kept stop at the budget too', async () => {
    const { run } = await build('rels-3x400k-links', 768, 120_000, { maxRelationships: 300_000 });
    expect(survived(run)).toEqual(SURVIVED);
    expect(run.probe.keptRelationships).toBeLessThanOrEqual(300_000);
    expect(run.probe.keptRelationships).toBeGreaterThan(290_000);
    expect(run.probe.limitLoss).toMatchObject([{ label: 'part(s) with relationships not read' }]);
  }, 240_000);

  it('cells-4x1500k: four sheets of 1.5 million cells are cut at the 5,000,000 of maxTotalCells, and the loss says so (the host had one warning string)', async () => {
    const { run } = await build('cells-4x1500k', 2048, 300_000);
    expect(survived(run)).toEqual(SURVIVED);
    expect(run.probe.sheets.map((s) => s.cells)).toEqual([1_500_000, 1_500_000, 1_500_000, 500_000]);
    expect(run.probe.firstWarnings).toEqual([expect.stringMatching(/^S4: the workbook has more cells than the import limit allows \(maxTotalCells, 5000000; raise it with the limits option\)/)]);
    expect(run.probe.limitLoss).toEqual([
      {
        id: 'xlsx.package.sizeLimit',
        label: 'sheet(s) cut at the cell limit',
        count: 1,
        where: ['S4'],
        effect: 'lost: the cells past maxTotalCells (5000000) over the workbook were not imported; raise it with the `limits` option of the import',
      },
    ]);
  }, 400_000);

  it('empty-cells-4x5m: four sheets of five million empty cells run into the scan budget (maxScannedElements), not a minute of scanning (it took 57 s)', async () => {
    const { run } = await build('empty-cells-4x5m', 768, 120_000);
    expect(survived(run)).toEqual(SURVIVED);
    expect(sum(run, 'cells')).toBe(0);
    expect(run.probe.firstWarnings).toEqual([expect.stringMatching(/maxScannedElements, 16777216; raise it with the limits option\)/)]);
    expect(run.probe.limitLoss).toMatchObject([{ label: 'sheet(s) cut at the scan limit', where: ['S4'] }]);
  }, 240_000);
});
