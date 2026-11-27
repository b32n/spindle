import {
  createApi,
  loadBuiltApi,
  type SheetsCoreModule,
} from '../../../../scripts/xlsx-fidelity/api';
import { WorkbookImpl } from '../workbook';
import { exportToXlsx, importFromXlsx, inspectXlsx } from '../export/xlsx';

describe('createApi', () => {
  it('binds the harness to a sheets-core module and asks for the full loss detail', async () => {
    const calls: unknown[][] = [];
    const core = {
      WorkbookImpl: class {
        constructor(
          public id: string,
          public name: string
        ) {}
        getData() {
          return { id: this.id, name: this.name };
        }
      },
      importFromXlsx: async (bytes: Uint8Array, wb: unknown) => (
        calls.push(['import', bytes, wb]),
        { warnings: [], loss: { format: 'xlsx', items: [] } }
      ),
      exportToXlsx: async (wb: unknown) => (calls.push(['export', wb]), new Uint8Array([7])),
      inspectXlsx: async (bytes: Uint8Array, options?: unknown) => (
        calls.push(['inspect', bytes, options]),
        { format: 'xlsx', items: [] }
      ),
    } as unknown as SheetsCoreModule;
    const api = createApi(core);
    const wb = api.createWorkbook();
    expect(wb.getData()).toEqual({ id: 'fixture', name: 'Untitled' });
    const bytes = new Uint8Array([1]);
    await api.importFromXlsx(bytes, wb);
    expect(await api.exportToXlsx(wb)).toEqual(new Uint8Array([7]));
    await api.inspectXlsx(bytes);
    expect(calls).toEqual([
      ['import', bytes, wb],
      ['export', wb],
      ['inspect', bytes, { detail: 'full' }],
    ]);
  });

  it('works against sheets-core source', async () => {
    const api = createApi({
      WorkbookImpl,
      importFromXlsx,
      exportToXlsx,
      inspectXlsx,
    } as unknown as SheetsCoreModule);
    const wb = api.createWorkbook();
    expect(wb).toBeInstanceOf(WorkbookImpl);
    expect((await api.inspectXlsx(await api.exportToXlsx(wb))).items).toEqual([]);
  });
});

describe('loadBuiltApi', () => {
  it('loads the sheets-core package by name (mapped to the source under Jest)', async () => {
    const api = await loadBuiltApi();
    expect(typeof api.createWorkbook().getData).toBe('function');
    expect((await api.inspectXlsx(await api.exportToXlsx(api.createWorkbook()))).format).toBe(
      'xlsx'
    );
  });
});
