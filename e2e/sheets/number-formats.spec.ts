import { test, expect, type Page } from '@playwright/test';
import ExcelJS from 'exceljs';
import { SheetsPage } from '../support/sheets';
import { collectAlerts, collectPageErrors, openViaChooser } from '../support/file-io';

const OPEN = 'Open file (.xlsx, .csv)';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// One cell per kind of number format, in column A: [value, Excel code, what the cell shows].
const CASES: Array<[number, string, string]> = [
  [1234.5, '"$"#,##0.00', '$1,234.50'], // currency (maps to the structured currency type)
  [46024, 'dddd, mmmm d, yyyy', 'Friday, January 2, 2026'], // a date no structured type has: kept as a custom code
  [46024.75, 'h:mm AM/PM', '6:00 PM'], // time
  [1.5, '[h]:mm:ss', '36:00:00'], // elapsed time
  [1500000, '0.0,,"M"', '1.5M'], // scaling comma and literal text
  [150, '[>100]"big";"small"', 'big'], // condition
];

async function workbookWithFormats(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Formats');
  CASES.forEach(([value, code], row) => {
    const cell = sheet.getCell(row + 1, 1);
    cell.value = value;
    cell.numFmt = code;
  });
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

const cellFormat = (page: Page, row: number, col: number) =>
  page.evaluate(
    ([r, c]) => {
      const wb = (window as any).__spindle.workbook;
      const cell = wb.getCell(undefined, r, c);
      return cell?.formatId ? (wb.getFormatPool().get(cell.formatId) ?? null) : null;
    },
    [row, col]
  );

async function openFormatCells(page: Page, sheet: SheetsPage, row: number) {
  const p = await sheet.cellCenter(row, 0);
  await page.mouse.click(p.x, p.y, { button: 'right' });
  await page.getByText('Format Cells...').click();
  await expect(page.getByTestId('format-cells-dialog')).toBeVisible();
}

test.describe('sheets: imported number formats', () => {
  let sheet: SheetsPage;
  let pageErrors: string[];
  let alerts: string[];
  test.beforeEach(async ({ page }) => {
    pageErrors = collectPageErrors(page);
    alerts = collectAlerts(page);
    await page.route(/fonts\.b32n\.io/, (r) => r.abort());
    sheet = new SheetsPage(page);
    await sheet.goto();
    await openViaChooser(page, OPEN, {
      name: 'formats.xlsx',
      mimeType: XLSX_MIME,
      buffer: await workbookWithFormats(),
    });
    await sheet.expectValue(0, 0, 1234.5);
  });
  test.afterEach(() => {
    expect(pageErrors, 'uncaught page errors').toEqual([]);
    expect(alerts, 'unexpected alert dialogs').toEqual([]);
  });

  test('S-FMC-06 imported currency, date, time, elapsed, scaled and conditional formats show as Excel shows them, and the Custom tab edits a code live', async ({
    page,
  }) => {
    // The grid is a canvas; the Format Cells dialog previews the selected cell with the format it carries.
    for (const [row, [value, , shown]] of CASES.entries()) {
      await openFormatCells(page, sheet, row);
      await expect(page.getByTestId('format-sample')).toHaveText(shown);
      await page.getByRole('button', { name: 'Cancel' }).click();
      await expect(page.getByTestId('format-cells-dialog')).toBeHidden();
      await sheet.expectValue(row, 0, value); // the stored value is the number from the file
    }
    // codes with a structured type keep it; the others are stored verbatim, so they survive a save
    expect(await cellFormat(page, 0, 0)).toMatchObject({
      type: 'currency',
      currencyCode: 'USD',
      decimalPlaces: 2,
    });
    expect(await cellFormat(page, 1, 0)).toEqual({ type: 'custom', pattern: 'dddd, mmmm d, yyyy' });
    expect(await cellFormat(page, 4, 0)).toEqual({ type: 'custom', pattern: '0.0,,"M"' });

    // The Custom tab opens on the stored code, previews as you type, refuses an invalid code and applies a valid one.
    await openFormatCells(page, sheet, 1);
    await expect(page.getByTestId('format-tab-custom')).toHaveAttribute('aria-pressed', 'true');
    const pattern = page.getByTestId('format-custom-pattern');
    await expect(pattern).toHaveValue('dddd, mmmm d, yyyy');

    await pattern.fill('dddd');
    await expect(page.getByTestId('format-sample')).toHaveText('Friday');

    await pattern.fill('yyyy 0 abc'); // letters that are not format tokens: Excel refuses the code
    await expect(page.getByTestId('format-custom-error')).toBeVisible();
    await expect(pattern).toHaveAttribute('aria-invalid', 'true');
    await expect(page.getByTestId('format-apply')).toBeDisabled();

    await pattern.fill('mmm d, yyyy');
    await expect(page.getByTestId('format-custom-error')).toBeHidden();
    await expect(page.getByTestId('format-sample')).toHaveText('Jan 2, 2026');
    await page.getByTestId('format-apply').click();
    await expect
      .poll(() => cellFormat(page, 1, 0))
      .toEqual({ type: 'custom', pattern: 'mmm d, yyyy' });
    await sheet.expectValue(1, 0, 46024); // formatting never changes the stored value
  });
});
