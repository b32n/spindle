import fs from 'node:fs';
import path from 'node:path';
import { importFromXlsx } from './xlsx';
import { WorkbookImpl } from '../workbook';

// The editor-level result of opening the file e2e/sheets/file-io.spec.ts (S-FIO-01) opens, through the path the app uses: the
// number formats the spec expects. (A LibreOffice file: it escapes the literals in its codes, `\$#,##0.00`.)
describe('importFromXlsx on e2e/fixtures/styled-multisheet.xlsx', () => {
  it('gives the cells the number formats S-FIO-01 expects', async () => {
    const bytes = new Uint8Array(fs.readFileSync(path.resolve(__dirname, '../../../../e2e/fixtures/styled-multisheet.xlsx')));
    const wb = new WorkbookImpl('wb', 'WB');
    await importFromXlsx(bytes, wb);
    const sales = wb.getData().sheets[0].id;
    const format = (row: number, col: number) => wb.getFormatPool().get(wb.getCell(sales, row, col)?.formatId as string);
    expect(format(1, 2)).toEqual({ type: 'custom', pattern: '$#,##0.00' });
    expect(format(1, 4)).toMatchObject({ type: 'percentage', decimalPlaces: 1 });
    expect(format(1, 5)).toMatchObject({ type: 'date', dateFormat: 'YYYY-MM-DD' });
    expect(format(1, 6)).toEqual({ type: 'custom', pattern: '#,##0.0" kg"' });
  });
});
