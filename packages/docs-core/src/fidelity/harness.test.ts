import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import { DocumentImpl } from '../document';
import { exportToDocx } from '../export/docx';
import { describeOutcome, goldensDirOf, runFixture, runHarness } from '../../../../scripts/docx-fidelity/harness';
import type { EditorApi } from '../../../../scripts/docx-fidelity/types';
import { sourceApi } from './support';

async function docxBytes(text: string): Promise<Uint8Array> {
  const doc = new DocumentImpl('d', 'Title');
  const data = doc.getData();
  data.sections[0].blocks = [{ id: 'b1', type: 'paragraph', content: [{ type: 'text', text }] }];
  doc.setData(data);
  return (await exportToDocx(doc)).bytes;
}

const entry = (id: string, extra: Record<string, unknown> = {}) => ({
  id,
  filename: `${id}.docx`,
  title: id,
  source: { tool: 'test', os: 'linux' },
  license: 'original',
  constructs: ['headings'],
  notes: 'n',
  ...extra,
});

describe('runFixture', () => {
  it('imports, exports and re-imports a valid file with a clean diff', async () => {
    const result = await runFixture(sourceApi, await docxBytes('Hello fidelity'));
    expect(result.model.importError).toBeNull();
    expect(JSON.stringify(result.model.document)).toContain('Hello fidelity');
    expect(result.roundtrip.status).toBe('ok');
    expect(result.roundtrip.diff?.total).toBe(0);
    expect(result.loss).toEqual({ format: 'docx', items: [] });
  });

  it('records an import error for bytes that are not a zip instead of crashing', async () => {
    const result = await runFixture(sourceApi, new TextEncoder().encode('this is not a docx'));
    expect(result.model.importError).toEqual(expect.any(String));
    expect(result.model.document).toBeNull();
    expect(result.roundtrip.status).toBe('skipped-import-error');
    expect(result.loss.items).toEqual([]);
  });

  it('records an import error for a zip with no document body', async () => {
    const zip = new JSZip();
    zip.file('hello.txt', 'hi');
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const result = await runFixture(sourceApi, bytes);
    expect(result.model.importError).toEqual(expect.any(String));
    expect(result.roundtrip.status).toBe('skipped-import-error');
  });

  it('records a truncated package as an import error', async () => {
    const good = await docxBytes('x');
    const result = await runFixture(sourceApi, good.slice(0, Math.floor(good.length / 2)));
    expect(result.model.importError).toEqual(expect.any(String));
  });

  describe('with a failing API', () => {
    const base = (over: Partial<EditorApi>): EditorApi => ({ ...sourceApi, ...over });
    const valid = () => docxBytes('x');

    it('records when inspectDocx throws', async () => {
      const result = await runFixture(base({ inspectDocx: () => Promise.reject(new Error('inspect boom')) }), await valid());
      expect(result.loss.items).toEqual([expect.objectContaining({ id: 'harness.inspectThrew', effect: 'inspect boom' })]);
      expect(result.model.importError).toBeNull();
    });

    it('records when export throws', async () => {
      const result = await runFixture(base({ exportToDocx: () => Promise.reject(new Error('export boom')) }), await valid());
      expect(result.roundtrip).toMatchObject({ status: 'export-error', error: 'export boom', diff: null });
    });

    it('records when the re-import throws', async () => {
      let calls = 0;
      const api = base({
        importFromDocx: (bytes, doc) => (++calls === 2 ? Promise.reject(new Error('reimport boom')) : sourceApi.importFromDocx(bytes, doc)),
        exportToDocx: async (doc) => ({ ...(await sourceApi.exportToDocx(doc)), warnings: ['export warning'] }),
      });
      const result = await runFixture(api, await valid());
      expect(result.roundtrip).toMatchObject({ status: 'reimport-error', error: 'reimport boom', exportWarnings: ['export warning'] });
    });

    it('stringifies non-Error throws and sorts loss items by id', async () => {
      const api = base({
        importFromDocx: () => Promise.reject('plain string'),
        inspectDocx: async () => ({
          format: 'docx',
          items: [
            { id: 'b', label: 'b', count: 1, effect: 'e' },
            { id: 'a', label: 'a', count: 2, effect: 'e' },
          ],
        }),
      });
      const result = await runFixture(api, await valid());
      expect(result.model.importError).toBe('plain string');
      expect(result.loss.items.map((i) => i.id)).toEqual(['a', 'b']);
    });
  });
});

describe('runHarness', () => {
  let dir: string;
  beforeEach(async () => {
    dir = mkdtempSync(path.join(tmpdir(), 'docx-harness-'));
    writeFileSync(path.join(dir, 'good.docx'), await docxBytes('Good one'));
    writeFileSync(path.join(dir, 'broken.docx'), 'not a zip at all');
    writeFileSync(
      path.join(dir, 'manifest.json'),
      JSON.stringify({ version: 1, fixtures: [entry('good'), entry('broken', { expectImportError: true })] })
    );
  });
  afterEach(() => rmSync(dir, { recursive: true, force: true }));

  it('fails until goldens exist, then passes after an update, and the update is stable', async () => {
    const first = await runHarness({ api: sourceApi, fixturesDir: dir });
    expect(first.ok).toBe(false);
    expect(first.fixtures.flatMap(describeOutcome)).toContain('good: model snapshot missing');

    const created = await runHarness({ api: sourceApi, fixturesDir: dir, update: true });
    expect(created.ok).toBe(true);
    expect(created.fixtures[0].status).toEqual({ model: 'created', loss: 'created', roundtrip: 'created' });

    const goldens = goldensDirOf(dir);
    const before = readFileSync(path.join(goldens, 'good.model.json'), 'utf8');
    expect(before.endsWith('\n')).toBe(true);
    expect(JSON.parse(readFileSync(path.join(goldens, 'broken.model.json'), 'utf8'))).toMatchObject({ document: null, importError: expect.any(String) });

    const again = await runHarness({ api: sourceApi, fixturesDir: dir });
    expect(again.ok).toBe(true);
    expect(again.fixtures.every((f) => Object.values(f.status).every((s) => s === 'match'))).toBe(true);
    expect(readFileSync(path.join(goldens, 'good.model.json'), 'utf8')).toBe(before);
  });

  it('fails when the importer output changes and passes again after an update', async () => {
    await runHarness({ api: sourceApi, fixturesDir: dir, update: true });
    const improved: EditorApi = {
      ...sourceApi,
      inspectDocx: async () => ({ format: 'docx', items: [{ id: 'docx.comments', label: 'comment(s)', count: 1, effect: 'not imported' }] }),
    };
    const drift = await runHarness({ api: improved, fixturesDir: dir });
    expect(drift.ok).toBe(false);
    expect(describeOutcome(drift.fixtures.find((f) => f.entry.id === 'good')!)).toEqual(['good: loss snapshot mismatch']);
    const updated = await runHarness({ api: improved, fixturesDir: dir, update: true });
    expect(updated.fixtures.find((f) => f.entry.id === 'good')!.status.loss).toBe('updated');
  });

  it('flags a manifest whose expectImportError disagrees with reality', async () => {
    writeFileSync(
      path.join(dir, 'manifest.json'),
      JSON.stringify({ version: 1, fixtures: [entry('good', { expectImportError: true }), entry('broken')] })
    );
    const report = await runHarness({ api: sourceApi, fixturesDir: dir, update: true });
    expect(report.ok).toBe(false);
    const lines = report.fixtures.flatMap(describeOutcome);
    expect(lines).toContain('good: manifest says importFromDocx should throw, but the import succeeded');
    expect(lines.find((l) => l.startsWith('broken: importFromDocx threw unexpectedly'))).toBeDefined();
  });

  it('restricts to --only ids and reports orphan goldens (deleting them on update)', async () => {
    await runHarness({ api: sourceApi, fixturesDir: dir, update: true });
    writeFileSync(path.join(goldensDirOf(dir), 'removed.model.json'), '{}\n');
    const subset = await runHarness({ api: sourceApi, fixturesDir: dir, only: ['good'] });
    expect(subset.fixtures.map((f) => f.entry.id)).toEqual(['good']);
    expect(subset.orphanGoldens).toEqual(['removed.model.json']);
    expect(subset.ok).toBe(false);
    const cleaned = await runHarness({ api: sourceApi, fixturesDir: dir, update: true });
    expect(cleaned.ok).toBe(true);
    expect((await runHarness({ api: sourceApi, fixturesDir: dir })).orphanGoldens).toEqual([]);
  });

  it('creates the goldens directory on update', async () => {
    mkdirSync(path.join(dir, 'goldens'), { recursive: true });
    const report = await runHarness({ api: sourceApi, fixturesDir: dir, update: true });
    expect(report.ok).toBe(true);
  });
});
