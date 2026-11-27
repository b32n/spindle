import JSZip from 'jszip';
import {
  PPTX_KINDS,
  crashedResult,
  opaquePartNames,
  pptxAdapter,
  runPptxFixture,
} from '../../../../scripts/pptx-fidelity/run';
import type { PptxEditorApi } from '../../../../scripts/pptx-fidelity/api';
import { detectOnlyApi, sourceApi } from './support';

const okLoss = {
  format: 'pptx',
  items: [
    { id: 'pptx.b', label: 'B', count: 1, effect: 'e' },
    { id: 'pptx.a', label: 'A', count: 2, effect: 'e' },
  ],
};
const model = { slides: [{ elements: [{ id: 'a' }, { id: 'b' }] }] };
const api = (over: Partial<PptxEditorApi> = {}): PptxEditorApi => ({
  inspectPptx: async () => okLoss,
  ...over,
});
const withImporter = (over: Partial<PptxEditorApi> = {}): PptxEditorApi =>
  api({
    importPptx: async () => ({ warnings: ['w'], presentation: model }),
    exportPptx: async () => ({ bytes: new Uint8Array([1]), warnings: ['not yet'] }),
    ...over,
  });
const bytes = new Uint8Array([1, 2, 3]);

async function minimalDeck(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types/>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="x"/>');
  zip.file('ppt/media/clip.mp4', 'x');
  return zip.generateAsync({ type: 'uint8array' });
}

describe('runPptxFixture without an importer', () => {
  it('records "no importer", skips the round trip, and still reports loss and ledger', async () => {
    const result = await runPptxFixture(detectOnlyApi, await minimalDeck());
    expect(result.model).toEqual({ importError: 'no importer', warnings: [], presentation: null });
    expect(result.roundtrip).toEqual({
      status: 'skipped-import-error',
      exportWarnings: [],
      reimportWarnings: [],
      diff: null,
    });
    expect(result.loss.items.map((i) => i.id)).toEqual([]);
    expect(result.ledger.package).toBe('zip');
    expect(result.ledger.importError).toBe(true);
    expect(result.ledger.totals).toEqual({ mapped: 0, opaque: 0, dropped: 0, unaccounted: 3, total: 3 });
  });

  it('covers the media part by the loss id once the detector reports it', async () => {
    const result = await runPptxFixture(
      api({ inspectPptx: async () => ({ format: 'pptx', items: [okLoss.items[0], { id: 'pptx.media', label: 'm', count: 1, effect: 'e' }] }) }),
      await minimalDeck()
    );
    expect(result.ledger.parts.find((p) => p.name === 'ppt/media/clip.mp4')).toMatchObject({ status: 'dropped', loss: ['pptx.media'] });
    expect(result.ledger.totals).toMatchObject({ dropped: 1, unaccounted: 2 });
  });

  it('gives bytes that are not a package an unreadable ledger instead of crashing', async () => {
    const result = await runPptxFixture(detectOnlyApi, new Uint8Array([9, 9, 9]));
    expect(result.model.importError).toBe('no importer');
    expect(result.ledger).toMatchObject({ package: 'unreadable', importError: true });
    expect(result.loss).toEqual({ format: 'pptx', items: [] });
  });

  it('knows an OLE compound file when it sees one', async () => {
    const ole = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]);
    const result = await runPptxFixture(detectOnlyApi, ole);
    expect(result.ledger.package).toBe('ole');
    expect(result.loss.items.map((i) => i.id)).toEqual(['pptx.legacyBinary']);
  });

  it('sorts the loss items', async () => {
    expect((await runPptxFixture(api(), bytes)).loss.items.map((i) => i.id)).toEqual(['pptx.a', 'pptx.b']);
  });

  it('records a throwing inspector as a finding', async () => {
    const result = await runPptxFixture(api({ inspectPptx: () => Promise.reject(new Error('inspect broke')) }), bytes);
    expect(result.loss).toEqual({
      format: 'pptx',
      items: [{ id: 'harness.inspectThrew', label: 'inspectPptx threw', count: 1, effect: 'inspect broke' }],
    });
    const stringy = await runPptxFixture(api({ inspectPptx: () => Promise.reject('plain') }), bytes);
    expect(stringy.loss.items[0].effect).toBe('plain');
  });
});

describe('runPptxFixture with the io reader', () => {
  it('imports a deck and round-trips it through the writer', async () => {
    const result = await runPptxFixture(sourceApi, await minimalDeck());
    expect(result.model.importError).toBeNull();
    expect(result.model.presentation).toMatchObject({ slideSize: { w: 1280, h: 720 }, slides: [] });
    expect(result.model.warnings.length).toBeGreaterThan(0);
    expect(result.roundtrip).toMatchObject({ status: 'ok', exportWarnings: [], diff: { added: 1, removed: 0, changed: 0 } }); // the deck had no layout: the writer adds a blank one
    expect(result.ledger.importError).toBe(false);
    expect(result.ledger.totals).toMatchObject({ mapped: 2, dropped: 0, unaccounted: 1, total: 3 });
  });

  it('records the reader\'s named error for bytes that are not a package', async () => {
    const result = await runPptxFixture(sourceApi, new Uint8Array([9, 9, 9]));
    expect(result.model).toEqual({ importError: 'This is not a PowerPoint presentation (.pptx).', warnings: [], presentation: null });
  });
});

describe('runPptxFixture with a scripted importer (what the importer PRs will see)', () => {
  it('imports, round-trips and diffs the two models', async () => {
    const result = await runPptxFixture(withImporter(), bytes);
    expect(result.model).toEqual({ importError: null, warnings: ['w'], presentation: model, textParity: { sourceRuns: 0, modelRuns: 0, outOfScope: 0, inErrorSlides: 0, missing: [], extra: [] } });
    expect(result.roundtrip).toMatchObject({ status: 'ok', reimportWarnings: ['w'] });
    expect(result.roundtrip.diff?.total).toBe(0);
    expect(result.roundtrip.diff?.before).toEqual({ slides: 1, elements: 2 });
    expect(result.ledger.importError).toBe(false);
  });

  it('reports the differences between the first and the second import', async () => {
    let n = 0;
    const result = await runPptxFixture(
      withImporter({
        importPptx: async () => ({ warnings: [], presentation: { slides: [{ elements: n++ === 0 ? [{ id: 'a' }] : [{ id: 'a' }, { id: 'b' }] }] } }),
      }),
      bytes
    );
    expect(result.roundtrip.diff).toMatchObject({ total: 1, added: 1, byField: { 'added slides[].elements[]': 1 } });
  });

  it('records an import failure and normalises the model it gets', async () => {
    const imp = await runPptxFixture(withImporter({ importPptx: () => Promise.reject(new Error('import broke')) }), bytes);
    expect(imp.model).toEqual({ importError: 'import broke', warnings: [], presentation: null });
    expect(imp.roundtrip.status).toBe('skipped-import-error');
    const odd = await runPptxFixture(
      withImporter({ importPptx: async () => ({ warnings: [], presentation: { when: new Date(0), blob: new Uint8Array([1]), skip: undefined } }) }),
      bytes
    );
    expect(odd.model.presentation).toEqual({ when: '1970-01-01T00:00:00.000Z', blob: { bytes: 1, sha256: expect.any(String) } });
  });

  it('records export and re-import failures at the right level', async () => {
    const none = await runPptxFixture(api({ importPptx: async () => ({ warnings: [], presentation: model }) }), bytes);
    expect(none.roundtrip).toEqual({ status: 'export-error', error: 'no exporter', exportWarnings: [], reimportWarnings: [], diff: null });
    const warned = await runPptxFixture(withImporter(), bytes);
    expect(warned.roundtrip).toMatchObject({ status: 'ok', exportWarnings: ['not yet'] });
    const exp = await runPptxFixture(withImporter({ exportPptx: () => Promise.reject(new Error('export broke')) }), bytes);
    expect(exp.roundtrip).toMatchObject({ status: 'export-error', error: 'export broke' });
    let calls = 0;
    const re = await runPptxFixture(
      withImporter({ importPptx: async () => (++calls === 1 ? { warnings: [], presentation: model } : Promise.reject(new Error('second import broke'))) }),
      bytes
    );
    expect(re.roundtrip).toMatchObject({ status: 'reimport-error', error: 'second import broke', exportWarnings: ['not yet'] });
  });

  it('removes the checkout path from error messages', async () => {
    const result = await runPptxFixture(
      withImporter({ importPptx: () => Promise.reject(new Error(`cannot open ${process.cwd()}/x.pptx`)) }),
      bytes
    );
    expect(result.model.importError).toBe('cannot open <repo>/x.pptx');
  });
});

describe('opaquePartNames', () => {
  it('lists the package parts every opaque element refers to, groups included, and nothing else', () => {
    const opaque = (...names: Array<string | undefined>) => ({ kind: 'opaque', parts: names.map((partName) => ({ hash: 'h', role: 'r', ...(partName ? { partName } : {}) })) });
    const presentation = {
      slides: [
        { elements: [opaque('ppt/charts/chart1.xml', undefined, 'ppt/embeddings/a.xlsx'), { kind: 'shape' }, { kind: 'group', children: [opaque('ppt/media/clip.mp4'), { kind: 'group', children: [opaque('ppt/ink/ink1.xml')] }] }] },
        { elements: [opaque('ppt/charts/chart1.xml')] },
        { error: 'damaged' },
      ],
    };
    expect([...opaquePartNames(presentation)].sort()).toEqual(['ppt/charts/chart1.xml', 'ppt/embeddings/a.xlsx', 'ppt/ink/ink1.xml', 'ppt/media/clip.mp4']);
    expect(opaquePartNames(null).size).toBe(0);
    expect(opaquePartNames({}).size).toBe(0);
    expect(opaquePartNames({ slides: [{ elements: 'x' }] }).size).toBe(0);
  });

  it('turns the parts the reader kept into opaque ledger entries', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types/>');
    zip.file('ppt/charts/chart1.xml', '<c/>');
    const presentation = { slides: [{ elements: [{ kind: 'opaque', parts: [{ hash: 'h', role: 'chart', partName: 'ppt/charts/chart1.xml' }] }] }] };
    const r = await runPptxFixture(withImporter({ importPptx: async () => ({ warnings: [], presentation }) }), await zip.generateAsync({ type: 'uint8array' }));
    expect(r.ledger.parts.find((p) => p.name === 'ppt/charts/chart1.xml')).toMatchObject({ status: 'opaque', rule: 'opaque' });
    expect(r.ledger.totals.opaque).toBe(1);
  });
});

describe('crashed results and the adapter', () => {
  it('says every kind did not finish', () => {
    const r = crashedResult(`timed out in ${process.cwd()}`);
    expect(r.model).toEqual({ importError: 'timed out in <repo>', warnings: [], presentation: null });
    expect(r.loss.items[0]).toMatchObject({ id: 'harness.crashed', effect: 'timed out in <repo>' });
    expect(r.roundtrip.status).toBe('skipped-import-error');
    expect(r.ledger.package).toBe('unreadable');
  });

  it('exposes the four kinds and maps results onto the harness contract', async () => {
    const adapter = pptxAdapter(api());
    expect(adapter.kinds).toEqual(PPTX_KINDS);
    expect(PPTX_KINDS).toEqual(['model', 'loss', 'roundtrip', 'ledger']);
    const result = await adapter.run(bytes, {} as never);
    for (const kind of PPTX_KINDS) expect(adapter.snapshot(result, kind)).toBe(result[kind]);
    expect(adapter.importError(result)).toBe('no importer');
    expect(adapter.importError(adapter.crashed('x'))).toBe('x');
  });
});
