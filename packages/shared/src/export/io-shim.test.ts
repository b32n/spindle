import { exportXlsx, importXlsx, detectXlsxLoss } from './xlsx';
import { exportDocx, importDocx, detectDocxLoss } from './docx';

// `@b32nio/spindle-io` is an optional peer and is not resolvable from here, so it is mocked virtually.
const xlsx = {
  exportXlsx: jest.fn(async () => new Uint8Array([1])),
  importXlsx: jest.fn(async () => ({ sheets: [] })),
  detectXlsxLoss: jest.fn(async () => ({ format: 'xlsx', items: [] })),
};
const docx = {
  exportDocx: jest.fn(async () => ({ bytes: new Uint8Array([2]), warnings: [] })),
  importDocx: jest.fn(async () => ({ title: 't', sections: [], warnings: [] })),
  detectDocxLoss: jest.fn(async () => ({ format: 'docx', items: [] })),
};
jest.mock('@b32nio/spindle-io/xlsx', () => xlsx, { virtual: true });
jest.mock('@b32nio/spindle-io/docx', () => docx, { virtual: true });

describe('deprecated export/xlsx and export/docx shims', () => {
  it('delegate to @b32nio/spindle-io', async () => {
    const bytes = new Uint8Array([0]);
    expect(await exportXlsx({ sheets: [] })).toEqual(new Uint8Array([1]));
    await importXlsx(bytes);
    await detectXlsxLoss(bytes);
    expect(xlsx.importXlsx).toHaveBeenCalledWith(bytes);
    expect(xlsx.detectXlsxLoss).toHaveBeenCalledWith(bytes);

    expect((await exportDocx({ title: 't', sections: [] })).bytes).toEqual(new Uint8Array([2]));
    await importDocx(bytes);
    await detectDocxLoss(bytes);
    expect(docx.importDocx).toHaveBeenCalledWith(bytes);
    expect(docx.detectDocxLoss).toHaveBeenCalledWith(bytes);
  });
});
