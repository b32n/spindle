import { NO_IMPORTER, createApi, loadBuiltApi, type PptxModule } from '../../../../scripts/pptx-fidelity/api';
import { detectPptxLoss } from '@b32nio/spindle-io/pptx';
import { sourceApi } from './support';

describe('createApi', () => {
  it('binds the harness to a module and asks for the full loss detail', async () => {
    const calls: unknown[][] = [];
    const api = createApi({
      detectPptxLoss: async (bytes: Uint8Array, options?: unknown) => (
        calls.push(['inspect', bytes, options]),
        { format: 'pptx', items: [] }
      ),
    } as PptxModule);
    const bytes = new Uint8Array([1]);
    expect(await api.inspectPptx(bytes)).toEqual({ format: 'pptx', items: [] });
    expect(calls).toEqual([['inspect', bytes, { detail: 'full' }]]);
    // no importer and no exporter: the keys are absent, which is how the harness knows
    expect('importPptx' in api).toBe(false);
    expect('exportPptx' in api).toBe(false);
  });

  it('passes an importer and an exporter through once the module has them', async () => {
    const importPptx = async () => ({ warnings: ['w'], presentation: { slides: [] } });
    const exportPptx = async () => new Uint8Array([7]);
    const api = createApi({ detectPptxLoss, importPptx, exportPptx } as unknown as PptxModule);
    expect(api.importPptx).toBe(importPptx);
    expect(api.exportPptx).toBe(exportPptx);
    expect(Object.keys(createApi({ detectPptxLoss, importPptx } as unknown as PptxModule)).sort()).toEqual(['importPptx', 'inspectPptx']);
  });

  it('works against the io source', async () => {
    expect(await sourceApi.inspectPptx(new Uint8Array(0))).toEqual({ format: 'pptx', items: [] });
    expect(sourceApi.importPptx).toBeUndefined();
  });

  it('says what the model snapshot holds when there is no importer', () => {
    expect(NO_IMPORTER).toBe('no importer');
  });
});

describe('loadBuiltApi', () => {
  it('loads the io package by name (mapped to the source under Jest)', async () => {
    const api = await loadBuiltApi();
    expect((await api.inspectPptx(new Uint8Array(0))).format).toBe('pptx');
  });
});
