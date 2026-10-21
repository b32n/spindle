import { test, expect, type Page } from '@playwright/test';
import ExcelJS from 'exceljs';
import { SheetsPage } from '../support/sheets';
import {
  asUpload,
  captureDownload,
  collectAlerts,
  collectPageErrors,
  fixture,
  openViaChooser,
  type Downloaded,
} from '../support/file-io';
import { dumpWorkbook } from '../support/sheets-dump';

const OPEN = 'Open file (.xlsx, .csv)';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

async function download(page: Page, label: RegExp): Promise<Downloaded> {
  await page.getByRole('button', { name: 'Download', exact: true }).click();
  return captureDownload(page, () => page.getByText(label).click());
}
const downloadXlsx = (page: Page) => download(page, /Microsoft Excel/);
const downloadCsv = (page: Page) => download(page, /Comma-separated/);

async function loadXlsx(bytes: Buffer): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes as unknown as ExcelJS.Buffer);
  return wb;
}

test.describe('sheets: file open / download round-trips', () => {
  let sheet: SheetsPage;
  let pageErrors: string[];
  let alerts: string[];
  test.beforeEach(async ({ page }) => {
    pageErrors = collectPageErrors(page);
    alerts = collectAlerts(page);
    sheet = new SheetsPage(page);
    await sheet.goto();
  });
  test.afterEach(() => {
    expect(pageErrors, 'uncaught page errors').toEqual([]);
    expect(alerts, 'unexpected alert dialogs').toEqual([]);
  });

  async function openFixture(page: Page, name: string) {
    await openViaChooser(page, OPEN, fixture(name));
  }

  test('S-FIO-01 opening a LibreOffice .xlsx shows sheets, values, formulas, formats, merges and styles', async ({
    page,
  }) => {
    await openFixture(page, 'styled-multisheet.xlsx');
    await expect(page.getByText('Costs', { exact: true })).toBeVisible();
    await expect(page.getByText('Notes ünï', { exact: true })).toBeVisible();
    await sheet.expectValue(0, 0, 'Item');

    const [sales, costs, notes] = await dumpWorkbook(page);
    expect([sales.name, costs.name, notes.name]).toEqual(['Sales', 'Costs', 'Notes ünï']);

    // values + formulas (calculated, not just cached)
    expect(sales.cells.B2.value).toBe(10);
    expect(sales.cells.D2).toMatchObject({ value: 25, formula: '=B2*C2' });
    expect(sales.cells.E3).toMatchObject({ value: 0.51, formula: '=D3/D$5' });
    expect(sales.cells.B5).toMatchObject({ value: 22, formula: '=SUM(B2:B4)' });
    expect(sales.cells.D5).toMatchObject({ value: 100, formula: '=SUM(D2:D4)' });
    // cross-sheet formulas
    expect(costs.cells.B1).toMatchObject({ value: 100, formula: '=Sales!D5' });
    expect(costs.cells.B2).toMatchObject({ value: 20, formula: '=Sales!D5*0.2' });
    expect(costs.cells.B3).toMatchObject({ value: 80, formula: '=B1-B2' });
    expect(notes.cells.A1.value).toBe('Café — naïve 日本語');

    // number formats: currency, percentage, date, custom number with suffix
    expect(sales.cells.C2.format).toMatchObject({
      type: 'currency',
      currencyCode: 'USD',
      decimalPlaces: 2,
      useThousandsSeparator: true,
    });
    expect(sales.cells.E2.format).toMatchObject({ type: 'percentage', decimalPlaces: 1 });
    expect(sales.cells.F2).toMatchObject({
      value: 45366,
      format: { type: 'date', dateFormat: 'YYYY-MM-DD' },
    });
    expect(sales.cells.G2).toMatchObject({
      value: 1234.5,
      format: { type: 'number', decimalPlaces: 1, useThousandsSeparator: true },
    });

    // styles: bold/fill/font colour/borders/alignment on the header, italic+fill on the banner, underline+colour
    expect(sales.cells.A1.style).toMatchObject({
      bold: true,
      fontColor: '#ffffff',
      backgroundColor: '#1f4e79',
      textAlign: 'center',
      borderTop: '1px solid #000000',
      borderLeft: '1px solid #000000',
    });
    expect(sales.cells.A5.style).toMatchObject({ bold: true });
    expect(sales.cells.D5.style).toMatchObject({ borderTop: '1px solid #000000' });
    expect(sales.cells.A7.style).toMatchObject({
      italic: true,
      backgroundColor: '#ffe699',
      textAlign: 'center',
    });
    expect(sales.cells.A9.style).toMatchObject({
      fontColor: '#c00000',
      textDecoration: 'underline',
    });
    expect(sales.cells.C9.style).toMatchObject({ textAlign: 'right' });

    // merged A7:D8
    expect(sales.merges).toEqual([[6, 0, 7, 3]]);
    // column widths follow the file (A is wider than B)
    expect(sales.colWidths[0]).toBeGreaterThan(sales.colWidths[1]);
    expect(sales.colWidths[0]).toBeGreaterThan(130);
  });

  test('S-FIO-02 sheet tabs switch between imported sheets', async ({ page }) => {
    await openFixture(page, 'styled-multisheet.xlsx');
    await page.getByText('Costs', { exact: true }).click();
    await sheet.expectValue(0, 1, 100);
    await sheet.expectValue(2, 1, 80);
    await expect.poll(() => sheet.cellFormula(0, 1)).toBe('=Sales!D5');
    await page.getByText('Sales', { exact: true }).click();
    await sheet.expectValue(0, 0, 'Item');
  });

  test('S-FIO-03 opening an exceljs-written .xlsx restores frozen panes and column widths', async ({
    page,
  }) => {
    await openFixture(page, 'exceljs-frozen.xlsx');
    await sheet.expectValue(0, 0, 'Frozen header');
    const [frozen] = await dumpWorkbook(page, 4, 4);
    expect(frozen.name).toBe('Frozen');
    expect(frozen.frozen).toEqual({ rows: 2, cols: 1 });
    expect(frozen.colWidths[0]).toBeGreaterThan(frozen.colWidths[1]);
    expect(frozen.cells.D3).toMatchObject({ value: 7.5, formula: '=SUM(B3:C3)' });
    expect(frozen.cells.A1.style).toMatchObject({ bold: true });
  });

  test('S-FIO-04 XLSX -> XLSX: the downloaded file keeps sheets, formulas, formats, merges and styles @smoke', async ({
    page,
  }) => {
    await openFixture(page, 'styled-multisheet.xlsx');
    await sheet.expectValue(0, 0, 'Item');
    const original = await dumpWorkbook(page);

    const file = await downloadXlsx(page);
    expect(file.fileName).toMatch(/\.xlsx$/);
    const wb = await loadXlsx(file.bytes);
    expect(wb.worksheets.map((w) => w.name)).toEqual(['Sales', 'Costs', 'Notes ünï']);

    const ws = wb.getWorksheet('Sales')!;
    expect(ws.getCell('B2').value).toBe(10);
    expect(ws.getCell('D2').value).toMatchObject({ formula: 'B2*C2' });
    expect(ws.getCell('E2').value).toMatchObject({ formula: 'D2/D$5' });
    expect(ws.getCell('D5').value).toMatchObject({ formula: 'SUM(D2:D4)', result: 100 });
    expect(wb.getWorksheet('Costs')!.getCell('B1').value).toMatchObject({
      formula: 'Sales!D5',
      result: 100,
    });
    expect(wb.getWorksheet('Notes ünï')!.getCell('A1').value).toBe('Café — naïve 日本語');

    expect(ws.getCell('C2').numFmt).toMatch(/\$.*#,##0\.00/);
    expect(ws.getCell('E2').numFmt).toBe('0.0%');
    expect(ws.getCell('F2').numFmt).toMatch(/yyyy.?-?mm.?-?dd/i);
    expect(ws.getCell('G2').numFmt).toMatch(/#,##0\.0/);

    expect(ws.getCell('A1').font?.bold).toBe(true);
    expect(ws.getCell('A1').font?.color?.argb?.toLowerCase()).toMatch(/ffffff$/);
    expect((ws.getCell('A1').fill as ExcelJS.FillPattern).fgColor?.argb?.toLowerCase()).toMatch(
      /1f4e79$/
    );
    expect(ws.getCell('A1').alignment?.horizontal).toBe('center');
    expect(ws.getCell('A1').border?.top?.style).toBe('thin');
    expect(ws.getCell('A7').font?.italic).toBe(true);
    expect(ws.getCell('A9').font?.underline).toBeTruthy();
    expect(ws.getCell('A9').font?.color?.argb?.toLowerCase()).toMatch(/c00000$/);
    expect(ws.getCell('C9').alignment?.horizontal).toBe('right');
    expect(
      Object.keys(ws.model.merges ?? {}).length === 0 ? ws.model.merges : ws.model.merges
    ).toEqual(['A7:D8']);
    expect(ws.getColumn(1).width ?? 0).toBeGreaterThan(ws.getColumn(2).width ?? 0);

    // Re-open the downloaded file through the UI: the model must match the first open.
    await openViaChooser(page, OPEN, asUpload(file, XLSX_MIME));
    await expect
      .poll(async () => JSON.stringify(await dumpWorkbook(page)))
      .toBe(JSON.stringify(original));
  });

  test('S-FIO-05 XLSX -> XLSX keeps frozen panes and widths', async ({ page }) => {
    await openFixture(page, 'exceljs-frozen.xlsx');
    await sheet.expectValue(0, 0, 'Frozen header');
    const file = await downloadXlsx(page);
    const ws = (await loadXlsx(file.bytes)).getWorksheet('Frozen')!;
    expect(ws.views[0]).toMatchObject({ state: 'frozen', xSplit: 1, ySplit: 2 });
    expect(ws.getColumn(1).width ?? 0).toBeGreaterThan(ws.getColumn(2).width ?? 0);
    expect(ws.getCell('A1').font?.bold).toBe(true);
    expect(ws.getCell('D3').value).toMatchObject({ formula: 'SUM(B3:C3)' });
    expect(ws.rowCount).toBe(40);
  });

  test('S-FIO-06 an edit made after opening is part of the downloaded .xlsx', async ({ page }) => {
    await openFixture(page, 'styled-multisheet.xlsx');
    await sheet.expectValue(0, 0, 'Item');
    // Imported column widths differ from the grid defaults, so navigate by keyboard, not pixels.
    await sheet.select(0, 0);
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowRight');
    await expect(sheet.canvas).toHaveAttribute('data-active-cell', 'B2');
    await sheet.type('99'); // Sales!B2 10 -> 99
    await sheet.expectValue(1, 3, 247.5); // D2 = 99 * 2.5
    const ws = (await loadXlsx((await downloadXlsx(page)).bytes)).getWorksheet('Sales')!;
    expect(ws.getCell('B2').value).toBe(99);
    expect(ws.getCell('D2').value).toMatchObject({ formula: 'B2*C2', result: 247.5 });
  });

  test('S-FIO-07 CSV with quotes, commas and newlines opens into the right cells', async ({
    page,
  }) => {
    await openFixture(page, 'quoting.csv');
    await sheet.expectValue(0, 0, 'id');
    const [s] = await dumpWorkbook(page);
    expect(s.name).toBe('Sheet1');
    expect(s.cells.B2.value).toBe('Smith, John');
    expect(s.cells.C2.value).toBe('He said "hello" twice');
    expect(s.cells.D2.value).toBe(1234.5);
    expect(s.cells.B3.value).toBe('Multi\nline cell');
    expect(s.cells.C3.value).toBe('Line one\nLine two');
    expect(s.cells.D3.value).toBe(-7);
    expect(s.cells.C4).toBeUndefined(); // empty field stays empty
    expect(s.cells.B5.value).toBe('  padded  ');
    expect(s.cells.D5.value).toBe('007'); // leading zeros survive as text
    expect(s.cells.B6.value).toBe('=1+1 text'); // text, not a live formula
    expect(s.cells.B6.formula).toBeUndefined();
    expect(s.cells.C6.value).toBe('quote " at end"');
    expect(s.cells.D6.value).toBe('1e3');
  });

  test('S-FIO-08 CSV -> CSV: the downloaded text is valid CSV that re-opens to the same cells @smoke', async ({
    page,
  }) => {
    await openFixture(page, 'quoting.csv');
    await sheet.expectValue(0, 0, 'id');
    const original = await dumpWorkbook(page);
    const file = await downloadCsv(page);
    expect(file.fileName).toBe('quoting.csv');
    expect(file.bytes.toString('utf8')).toBe(
      [
        'id,name,notes,amount',
        '1,"Smith, John","He said ""hello"" twice",1234.5',
        '2,"Multi\nline cell","Line one\nLine two",-7',
        '3,plain,,0.1',
        '4,  padded  ,"trailing comma,",007',
        // The export prefixes a quote to a text cell starting with "=" (CSV-injection guard).
        `5,'=1+1 text,"quote "" at end""",1e3`,
      ].join('\n')
    );
    await openViaChooser(page, OPEN, asUpload(file, 'text/csv'));
    await expect.poll(async () => (await dumpWorkbook(page))[0].cells.B6?.value).toBe("'=1+1 text");
    const reopened = (await dumpWorkbook(page))[0];
    // Everything except the guarded formula-looking cell is identical.
    delete reopened.cells.B6;
    const expected = original[0];
    delete expected.cells.B6;
    expect(reopened.cells).toEqual(expected.cells);
  });

  test('S-FIO-09 unicode CSV (accents, CJK, emoji, RTL, CRLF) round-trips byte for byte', async ({
    page,
  }) => {
    await openFixture(page, 'unicode.csv');
    await expect.poll(async () => (await dumpWorkbook(page))[0].cells.B2?.value).toBe('Café crème');
    const [s] = await dumpWorkbook(page);
    expect(s.cells.C2.value).toBe("à la carte, s'il vous plaît");
    expect(s.cells.B4.value).toBe('こんにちは世界');
    expect(s.cells.B5.value).toBe('你好，世界');
    expect(s.cells.B6.value).toBe('🚀 Launch 🎉');
    expect(s.cells.C6.value).toBe('family 👨‍👩‍👧‍👦, and flag 🇯🇵');
    expect(s.cells.B7.value).toBe('مرحبا بالعالم');
    expect(s.cells.B8.value).toBe('שלום עולם');

    const file = await downloadCsv(page);
    const text = file.bytes.toString('utf8');
    // Same content; the export normalises CRLF to LF and drops the trailing newline.
    const source = (await import('node:fs')).readFileSync(fixture('unicode.csv'), 'utf8');
    expect(text).toBe(source.replace(/\r\n/g, '\n').replace(/\n$/, ''));
  });

  test('S-FIO-10 XLSX -> CSV exports only the active sheet, with raw (unformatted) numbers', async ({
    page,
  }) => {
    await openFixture(page, 'styled-multisheet.xlsx');
    await sheet.expectValue(0, 0, 'Item');
    const csv = (await downloadCsv(page)).bytes.toString('utf8').split('\n');
    expect(csv[0]).toBe('Item,Qty,Price,Total,Share,Date,Weight');
    expect(csv[1].split(',').filter((_, i) => i !== 3 && i !== 4)).toEqual([
      'Widget',
      '10',
      '2.5',
      '45366',
      '1234.5',
    ]);
    expect(csv.join('\n')).not.toContain('Sales total'); // the Costs sheet is not included
  });

  // Bug (https://github.com/b32n/spindle/issues/77): CSV export writes the formula text, guarded with a leading apostrophe (`'=B2*C2`),
  // for formula cells instead of their computed values (25), so the CSV is not usable elsewhere.
  test.fixme('S-FIO-14 XLSX -> CSV exports computed values for formula cells', async ({ page }) => {
    await openFixture(page, 'styled-multisheet.xlsx');
    await sheet.expectValue(0, 0, 'Item');
    const csv = (await downloadCsv(page)).bytes.toString('utf8').split('\n');
    expect(csv[1]).toBe('Widget,10,2.5,25,0.25,45366,1234.5');
    expect(csv[4]).toBe('Total,22,,100,1,,');
  });

  test('S-FIO-11 CSV -> XLSX: values, text with newlines/unicode and numbers survive', async ({
    page,
  }) => {
    await openFixture(page, 'quoting.csv');
    await sheet.expectValue(0, 0, 'id');
    const ws = (await loadXlsx((await downloadXlsx(page)).bytes)).worksheets[0];
    expect(ws.getCell('B2').value).toBe('Smith, John');
    expect(ws.getCell('B3').value).toBe('Multi\nline cell');
    expect(ws.getCell('D2').value).toBe(1234.5);
    expect(ws.getCell('D5').value).toBe('007');
    expect(ws.getCell('D6').value).toBe('1e3');
  });

  test('S-FIO-12 semicolon- and tab-delimited CSV are not split (no delimiter detection)', async ({
    page,
  }) => {
    // Current behaviour: only commas delimit. A European (semicolon) export lands in column A.
    await openFixture(page, 'semicolon.csv');
    await sheet.expectValue(0, 0, 'name;qty;price');
    await openFixture(page, 'tabs.tsv.csv');
    await sheet.expectValue(0, 0, 'name\tqty\tprice');
  });

  test('S-FIO-13 opening a 10,000-row CSV stays within budget and the grid stays usable', async ({
    page,
  }, testInfo) => {
    const rows = ['id,name,value'];
    for (let i = 1; i <= 10_000; i++) rows.push(`${i},"Row ${i}, with comma",${i * 3}`);
    const started = Date.now();
    await openViaChooser(page, OPEN, {
      name: 'big.csv',
      mimeType: 'text/csv',
      buffer: Buffer.from(rows.join('\n')),
    });
    await sheet.expectValue(0, 0, 'id');
    await expect
      .poll(async () =>
        page.evaluate(() => (window as any).__spindle.workbook.getCellValue(undefined, 10_000, 2))
      )
      .toBe(30_000);
    const loadMs = Date.now() - started;
    testInfo.annotations.push({ type: 'load-ms', description: String(loadMs) });
    expect(loadMs, 'open + render of 10k rows').toBeLessThan(10_000);

    // Still responsive: select and edit a cell, then export everything.
    await sheet.enter(2, 1, 'edited');
    await sheet.expectValue(2, 1, 'edited');
    const csv = (await downloadCsv(page)).bytes.toString('utf8').split('\n');
    expect(csv).toHaveLength(10_001);
    expect(csv[0]).toBe('id,name,value');
    expect(csv[2]).toBe('2,edited,6');
    expect(csv[10_000]).toBe('10000,"Row 10000, with comma",30000');
  });
});
