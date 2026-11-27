import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';
import { WorkbookProvider } from '../context/WorkbookContext';
import { WorkbookCanvas } from './WorkbookCanvas';
import * as fileIo from '../file-io';

// The toolbar is replaced by a stub that keeps the `onOpenFile` the canvas gives it, so the test can open a file as the Open button would.
let onOpenFile: ((file: File) => Promise<void>) | undefined;
let onDownload: ((format: 'xlsx' | 'csv') => Promise<void>) | undefined;
jest.mock('./Toolbar', () => ({
  Toolbar: (props: { onOpenFile?: (file: File) => Promise<void>; onDownload?: (format: 'xlsx' | 'csv') => Promise<void> }) => {
    onOpenFile = props.onOpenFile;
    onDownload = props.onDownload;
    return null;
  },
}));
// the browser download is not what is tested
const mockDownload = jest.fn();
jest.mock('@b32nio/spindle-shared/react', () => ({ ...jest.requireActual('@b32nio/spindle-shared/react'), downloadFile: (...args: unknown[]) => mockDownload(...args) }));

// (the grid draws on a canvas, which is not what is tested)
jest.mock('./CanvasGrid', () => ({ CanvasGrid: () => null }));

describe('WorkbookCanvas importLimits', () => {
  beforeAll(() => {
    // layout effects under server rendering
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });
  afterAll(() => jest.restoreAllMocks());

  const render = (props: React.ComponentProps<typeof WorkbookCanvas>) => {
    onOpenFile = undefined;
    renderToStaticMarkup(
      <WorkbookProvider workbook={new WorkbookImpl('wb', 'WB')}>
        <WorkbookCanvas importConfirmation="never" fileActions {...props} />
      </WorkbookProvider>
    );
    return onOpenFile as unknown as (file: File) => Promise<void>;
  };

  const csv = (): File => new File(['a,b\n1,2\n'], 'data.csv');

  it('hands the limits to the import and the loss, with the warnings, to onFileOpened', async () => {
    const spy = jest.spyOn(fileIo, 'importWorkbookFile');
    const opened: unknown[] = [];
    await render({ importLimits: { maxTotalCells: 7 }, onFileOpened: (info) => opened.push(info) })(csv());
    expect(spy).toHaveBeenCalledWith(expect.any(WorkbookImpl), expect.any(File), { limits: { maxTotalCells: 7 } });
    expect(opened).toMatchObject([{ fileName: 'data.csv', warnings: [], loss: { items: [] } }]);
    spy.mockRestore();
  });

  it('passes no limits when the prop is not given (the reader defaults)', async () => {
    const spy = jest.spyOn(fileIo, 'importWorkbookFile');
    await render({})(csv());
    expect(spy).toHaveBeenCalledWith(expect.any(WorkbookImpl), expect.any(File), { limits: undefined });
    spy.mockRestore();
  });

  it('hands recalculateOnLoad to the import (off unless the host asks)', async () => {
    const spy = jest.spyOn(fileIo, 'importWorkbookFile');
    await render({ recalculateOnLoad: true })(csv());
    expect(spy).toHaveBeenLastCalledWith(expect.any(WorkbookImpl), expect.any(File), { limits: undefined, recalculateOnLoad: true });
    await render({})(csv());
    expect(spy.mock.lastCall?.[2]?.recalculateOnLoad).toBeUndefined();
    spy.mockRestore();
  });

  it('hands maxArrayCells to the import (the reader default unless the host sets it)', async () => {
    const spy = jest.spyOn(fileIo, 'importWorkbookFile');
    await render({ maxArrayCells: 50 })(csv());
    expect(spy).toHaveBeenLastCalledWith(expect.any(WorkbookImpl), expect.any(File), { limits: undefined, maxArrayCells: 50 });
    await render({})(csv());
    expect(spy.mock.lastCall?.[2]?.maxArrayCells).toBeUndefined();
    spy.mockRestore();
  });

  it('hands what an xlsx download cut to onExportWarning, and still downloads the file', async () => {
    const wb = new WorkbookImpl('wb', 'WB');
    const warnings: string[] = [];
    const file = { data: new Uint8Array([1]), fileName: 'WB.xlsx', mimeType: 'x/y' };
    // the export's own caps are tested in sheets-core: here the helper reports a cut, as it does for a result over them
    const spy = jest.spyOn(fileIo, 'exportWorkbookFile').mockImplementation(async (_wb, _format, _name, options) => {
      options?.onWarning?.('Sheet1!A1: cut');
      return file;
    });
    onDownload = undefined;
    mockDownload.mockClear();
    renderToStaticMarkup(
      <WorkbookProvider workbook={wb}>
        <WorkbookCanvas fileActions onExportWarning={(m) => warnings.push(m)} />
      </WorkbookProvider>
    );
    await onDownload!('xlsx');
    expect(warnings).toEqual(['Sheet1!A1: cut']);
    expect(mockDownload).toHaveBeenCalledWith(file.data, 'WB.xlsx', 'x/y');
    spy.mockRestore();
  });

  it('reports a file the reader refuses to onFileError', async () => {
    const errors: Array<[Error, string]> = [];
    const bad = new File([new Uint8Array([1, 2, 3])], 'bad.xlsx');
    await render({ importLimits: { maxEntries: 1 }, onFileError: (e, action) => errors.push([e, action]) })(bad);
    expect(errors).toHaveLength(1);
    expect(errors[0][1]).toBe('open');
  });
});
