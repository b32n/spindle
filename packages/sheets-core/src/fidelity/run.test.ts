import JSZip from 'jszip';
import {
  crashedResult,
  runXlsxFixture,
  xlsxAdapter,
  XLSX_KINDS,
} from '../../../../scripts/xlsx-fidelity/run';
import type { WorkbookLike, XlsxEditorApi } from '../../../../scripts/xlsx-fidelity/api';
import { WorkbookImpl } from '../workbook';
import { exportToXlsx } from '../export/xlsx';
import { sourceApi } from './support';

const emptyData = {
  id: 'w',
  name: 'n',
  activeSheetId: 's',
  defaultRowHeight: 20,
  defaultColWidth: 100,
  stylePool: {},
  formatPool: {},
  sheets: [{ id: 's', name: 'S', rowCount: 1, colCount: 1, cells: [], config: {} }],
};
const fakeWorkbook = (): WorkbookLike => ({ getData: () => emptyData });
const okLoss = {
  format: 'xlsx',
  items: [
    { id: 'xlsx.b', label: 'B', count: 1, effect: 'e' },
    { id: 'xlsx.a', label: 'A', count: 2, effect: 'e' },
  ],
};
const api = (over: Partial<XlsxEditorApi> = {}): XlsxEditorApi => ({
  createWorkbook: fakeWorkbook,
  importFromXlsx: async () => ({ warnings: ['w'], loss: okLoss }),
  exportToXlsx: async () => new Uint8Array([1]),
  inspectXlsx: async () => okLoss,
  ...over,
});
const bytes = new Uint8Array([1, 2, 3]);

async function sampleWorkbook(): Promise<Uint8Array> {
  const wb = new WorkbookImpl('id', 'name');
  wb.setCellValue(undefined, 0, 0, 5);
  wb.setCellValue(undefined, 1, 0, 'text');
  wb.setFormula(undefined, 2, 0, '=A1*2');
  return exportToXlsx(wb);
}

describe('runXlsxFixture with the real engine', () => {
  it('imports, round-trips, builds the ledger and compares values for a valid workbook', async () => {
    const result = await runXlsxFixture(sourceApi, await sampleWorkbook());
    expect(result.model.importError).toBeNull();
    expect(result.model.workbook).toMatchObject({
      sheets: [
        { cells: { A1: { value: 5 }, A2: { value: 'text' }, A3: { value: 10, formula: '=A1*2' } } },
      ],
    });
    expect(result.roundtrip.status).toBe('ok');
    expect(result.roundtrip.diff?.total).toBe(0);
    expect(result.ledger.package).toBe('zip');
    expect(result.ledger.importError).toBe(false);
    expect(result.ledger.totals.total).toBeGreaterThan(3);
    expect(result.values).toMatchObject({
      status: 'ok',
      formulaCells: 1,
      compared: 1,
      matched: 1,
      mismatched: 0,
    });
    expect(result.loss).toEqual({ format: 'xlsx', items: [] });
  });

  it('records an import error for bytes that are not a package instead of crashing', async () => {
    const result = await runXlsxFixture(sourceApi, new Uint8Array([9, 9, 9]));
    expect(result.model.importError).toEqual(expect.any(String));
    expect(result.model.workbook).toBeNull();
    expect(result.roundtrip.status).toBe('skipped-import-error');
    expect(result.ledger).toMatchObject({ package: 'unreadable', importError: true });
    expect(result.values.status).toBe('unreadable-package');
  });

  it('knows an OLE compound file when it sees one', async () => {
    const ole = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]);
    const result = await runXlsxFixture(sourceApi, ole);
    expect(result.ledger.package).toBe('ole');
    expect(result.values.status).toBe('unreadable-package');
    expect(result.model.importError).toEqual(expect.any(String));
  });

  it('gives a zip that has no workbook an import error but still a ledger and values status', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types/>');
    zip.file('docProps/app.xml', '<a/>');
    const result = await runXlsxFixture(sourceApi, await zip.generateAsync({ type: 'uint8array' }));
    expect(result.model.importError).toEqual(expect.any(String));
    expect(result.ledger.package).toBe('zip');
    expect(result.ledger.totals).toMatchObject({ total: 2, unaccounted: 2, mapped: 0 });
    expect(result.values).toMatchObject({ status: 'skipped-import-error', formulaCells: 0 });
  });
});

describe('runXlsxFixture with a scripted API', () => {
  it('sorts the loss items and keeps the warnings', async () => {
    const result = await runXlsxFixture(api(), bytes);
    expect(result.loss.items.map((i) => i.id)).toEqual(['xlsx.a', 'xlsx.b']);
    expect(result.model.warnings).toEqual(['w']);
    expect(result.roundtrip.status).toBe('ok');
  });

  it('records a throwing inspector as a finding', async () => {
    const result = await runXlsxFixture(
      api({ inspectXlsx: () => Promise.reject(new Error('inspect broke')) }),
      bytes
    );
    expect(result.loss).toEqual({
      format: 'xlsx',
      items: [
        {
          id: 'harness.inspectThrew',
          label: 'inspectXlsx threw',
          count: 1,
          effect: 'inspect broke',
        },
      ],
    });
    const stringy = await runXlsxFixture(
      api({ inspectXlsx: () => Promise.reject('plain') }),
      bytes
    );
    expect(stringy.loss.items[0].effect).toBe('plain');
  });

  it('records import, export and re-import failures at the right level', async () => {
    const imp = await runXlsxFixture(
      api({ importFromXlsx: () => Promise.reject(new Error('import broke')) }),
      bytes
    );
    expect(imp.model).toEqual({ importError: 'import broke', warnings: [], workbook: null });
    expect(imp.roundtrip.status).toBe('skipped-import-error');

    const exp = await runXlsxFixture(
      api({ exportToXlsx: () => Promise.reject(new Error('export broke')) }),
      bytes
    );
    expect(exp.roundtrip).toEqual({
      status: 'export-error',
      error: 'export broke',
      exportWarnings: [],
      reimportWarnings: [],
      diff: null,
    });

    let calls = 0;
    const re = await runXlsxFixture(
      api({
        importFromXlsx: async () =>
          ++calls === 1
            ? { warnings: [], loss: okLoss }
            : Promise.reject(new Error('second import broke')),
      }),
      bytes
    );
    expect(re.roundtrip).toMatchObject({ status: 'reimport-error', error: 'second import broke' });
  });

  it('carries re-import warnings and an empty diff when the second import matches', async () => {
    let calls = 0;
    const result = await runXlsxFixture(
      api({ importFromXlsx: async () => ({ warnings: [`call ${++calls}`], loss: okLoss }) }),
      bytes
    );
    expect(result.roundtrip).toMatchObject({ status: 'ok', reimportWarnings: ['call 2'] });
    expect(result.roundtrip.diff?.total).toBe(0);
  });

  it('removes the checkout path from error messages', async () => {
    const result = await runXlsxFixture(
      api({
        importFromXlsx: () => Promise.reject(new Error(`cannot open ${process.cwd()}/x.xlsx`)),
      }),
      bytes
    );
    expect(result.model.importError).toBe('cannot open <repo>/x.xlsx');
  });
});

describe('crashed results and the adapter', () => {
  it('says every kind did not finish', () => {
    const r = crashedResult(`timed out in ${process.cwd()}`);
    expect(r.model).toEqual({ importError: 'timed out in <repo>', warnings: [], workbook: null });
    expect(r.loss.items[0]).toMatchObject({ id: 'harness.crashed', effect: 'timed out in <repo>' });
    expect(r.roundtrip.status).toBe('skipped-import-error');
    expect(r.ledger.package).toBe('unreadable');
    expect(r.values.status).toBe('unreadable-package');
  });

  it('exposes the five kinds and maps results onto the harness contract', async () => {
    const adapter = xlsxAdapter(api());
    expect(adapter.kinds).toEqual(XLSX_KINDS);
    expect(XLSX_KINDS).toEqual(['model', 'loss', 'roundtrip', 'ledger', 'values']);
    const result = await adapter.run(bytes, {} as never);
    for (const kind of XLSX_KINDS) expect(adapter.snapshot(result, kind)).toBe(result[kind]);
    expect(adapter.importError(result)).toBeNull();
    expect(adapter.importError(adapter.crashed('x'))).toBe('x');
  });
});
