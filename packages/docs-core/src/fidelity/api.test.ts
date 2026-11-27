import { createApi, loadBuiltApi, type DocsCoreModule } from '../../../../scripts/docx-fidelity/api';
import { DocumentImpl } from '../document';
import { exportToDocx, importFromDocx, inspectDocx } from '../export/docx';

describe('createApi', () => {
  it('binds the harness to a docs-core module and asks for the full loss detail', async () => {
    const calls: unknown[][] = [];
    const core = {
      DocumentImpl: class {
        constructor(
          public id: string,
          public name: string
        ) {}
        getData() {
          return { id: this.id, name: this.name };
        }
      },
      importFromDocx: async (bytes: Uint8Array, doc: unknown) => (calls.push(['import', bytes, doc]), { warnings: [], loss: { format: 'docx', items: [] } }),
      exportToDocx: async (doc: unknown) => (calls.push(['export', doc]), { bytes: new Uint8Array([7]), warnings: [] }),
      inspectDocx: async (bytes: Uint8Array, options?: unknown) => (calls.push(['inspect', bytes, options]), { format: 'docx', items: [] }),
    } as unknown as DocsCoreModule;
    const api = createApi(core);
    const doc = api.createDocument();
    expect(doc.getData()).toEqual({ id: 'fixture', name: 'Untitled' });
    const bytes = new Uint8Array([1]);
    await api.importFromDocx(bytes, doc);
    expect((await api.exportToDocx(doc)).bytes).toEqual(new Uint8Array([7]));
    await api.inspectDocx(bytes);
    expect(calls).toEqual([
      ['import', bytes, doc],
      ['export', doc],
      ['inspect', bytes, { detail: 'full' }],
    ]);
  });

  it('works against docs-core source', async () => {
    const api = createApi({ DocumentImpl, importFromDocx, exportToDocx, inspectDocx } as unknown as DocsCoreModule);
    const doc = api.createDocument();
    expect(doc).toBeInstanceOf(DocumentImpl);
    expect((await api.inspectDocx((await api.exportToDocx(doc)).bytes)).items).toEqual([]);
  });
});

describe('loadBuiltApi', () => {
  it('loads the docs-core package by name (mapped to the source under Jest)', async () => {
    const api = await loadBuiltApi();
    expect(typeof api.createDocument().getData).toBe('function');
    expect((await api.inspectDocx((await api.exportToDocx(api.createDocument())).bytes)).format).toBe('docx');
  });
});
