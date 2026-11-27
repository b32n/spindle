import { WorkbookImpl } from '@b32nio/spindle-sheets-core';
import * as core from '@b32nio/spindle-sheets-core';
import { exportWorkbookFile } from './file-io';

// The export's caps are tested in sheets-core (a result over them is a million cells): here the core export is stood in for by one
// that reports a cut, to show that the download path hands the host's `onWarning` to it.
jest.mock('@b32nio/spindle-sheets-core', () => ({ ...jest.requireActual('@b32nio/spindle-sheets-core'), exportToXlsx: jest.fn() }));
const exportToXlsx = core.exportToXlsx as jest.Mock;

describe('exportWorkbookFile warnings', () => {
  beforeEach(() => exportToXlsx.mockReset());

  it('passes onWarning to the xlsx export, which reports what it cut', async () => {
    exportToXlsx.mockImplementation(async (_wb: unknown, options?: { onWarning?: (m: string) => void }) => {
      options?.onWarning?.('Sheet1!A1: cut');
      return new Uint8Array([1]);
    });
    const warnings: string[] = [];
    const wb = new WorkbookImpl('w', 'W');
    const out = await exportWorkbookFile(wb, 'xlsx', 'Big', { onWarning: (m) => warnings.push(m) });
    expect(warnings).toEqual(['Sheet1!A1: cut']);
    expect(out).toMatchObject({ fileName: 'Big.xlsx', data: new Uint8Array([1]) });
  });

  it('exports without a listener, and a csv never asks for one', async () => {
    exportToXlsx.mockResolvedValue(new Uint8Array([2]));
    const wb = new WorkbookImpl('w', 'W');
    await exportWorkbookFile(wb, 'xlsx');
    expect(exportToXlsx).toHaveBeenLastCalledWith(wb, { onWarning: undefined });
    await exportWorkbookFile(wb, 'csv', undefined, { onWarning: () => {} });
    expect(exportToXlsx).toHaveBeenCalledTimes(1);
  });
});
