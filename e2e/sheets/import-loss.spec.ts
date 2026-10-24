import { test, expect, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { SheetsPage } from '../support/sheets';
import { openViaChooser, fixture } from '../support/file-io';

const OPEN = 'Open file (.xlsx, .csv)';

type Opened = {
  fileName: string;
  loss: { format: string; items: Array<{ id: string; count: number }> };
};
const lastOpened = (page: Page) =>
  page.evaluate(() => (window as unknown as { __lastOpened?: Opened }).__lastOpened);

/** A copy of the frozen-header fixture with a conditional format and a chart bolted on. */
async function lossyXlsx(): Promise<{ name: string; mimeType: string; buffer: Buffer }> {
  const zip = await JSZip.loadAsync(
    (await import('node:fs')).readFileSync(fixture('exceljs-frozen.xlsx'))
  );
  const path = Object.keys(zip.files).find((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))!;
  const xml = await zip.file(path)!.async('string');
  zip.file(
    path,
    xml.replace(
      '</sheetData>',
      '</sheetData><conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="equal"><formula>1</formula></cfRule></conditionalFormatting>'
    )
  );
  zip.file('xl/charts/chart1.xml', '<c:chartSpace/>');
  return {
    name: 'features.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(await zip.generateAsync({ type: 'uint8array' })),
  };
}

test.describe('sheets: import loss report and confirmation', () => {
  let sheet: SheetsPage;
  const dialog = (page: Page) => page.getByRole('dialog', { name: /Import/ });
  test.beforeEach(async ({ page }) => {
    await page.route(/fonts\.b32n\.io/, (r) => r.abort());
    sheet = new SheetsPage(page);
  });

  test('S-FIO-16 a file with unsupported features asks first, listing what would be lost', async ({
    page,
  }) => {
    await sheet.goto();
    await openViaChooser(page, OPEN, await lossyXlsx());
    await expect(dialog(page)).toBeVisible();
    await expect(dialog(page).getByRole('heading')).toHaveText('Import “features.xlsx”?');
    const items = dialog(page).getByTestId('import-loss-list').locator('li');
    await expect(items).toHaveCount(2);
    await expect(items.filter({ hasText: 'conditional formatting' })).toHaveCount(1);
    await expect(items.filter({ hasText: 'chart' })).toHaveCount(1);
    // Nothing has been replaced yet.
    await sheet.expectValue(0, 0, 'Product');
  });

  test('S-FIO-17 Cancel (and Escape) keep the open workbook untouched', async ({ page }) => {
    await sheet.goto();
    await openViaChooser(page, OPEN, await lossyXlsx());
    await dialog(page).getByRole('button', { name: 'Cancel' }).click();
    await expect(dialog(page)).toBeHidden();
    await sheet.expectValue(0, 0, 'Product');
    expect(await lastOpened(page)).toBeUndefined();

    await openViaChooser(page, OPEN, await lossyXlsx());
    await expect(dialog(page)).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(dialog(page)).toBeHidden();
    await sheet.expectValue(0, 0, 'Product');
    expect(await lastOpened(page)).toBeUndefined();
  });

  test('S-FIO-18 "Import anyway" imports the file and reports the same loss to the host', async ({
    page,
  }) => {
    await sheet.goto();
    await openViaChooser(page, OPEN, await lossyXlsx());
    await dialog(page).getByRole('button', { name: 'Import anyway' }).click();
    await expect(dialog(page)).toBeHidden();
    await sheet.expectValue(0, 0, 'Frozen header');
    await expect
      .poll(async () => (await lastOpened(page))?.loss.items.map((i) => i.id).sort())
      .toEqual(['xlsx.charts', 'xlsx.conditionalFormatting']);
  });

  test('S-FIO-19 a file that loses nothing opens without a dialog', async ({ page }) => {
    await sheet.goto();
    await openViaChooser(page, OPEN, fixture('exceljs-frozen.xlsx'));
    await sheet.expectValue(0, 0, 'Frozen header');
    await expect(dialog(page)).toHaveCount(0);
    expect((await lastOpened(page))!.loss.items).toEqual([]);
    await openViaChooser(page, OPEN, fixture('quoting.csv'));
    await expect.poll(async () => (await lastOpened(page))?.fileName).toBe('quoting.csv');
    await expect(dialog(page)).toHaveCount(0);
    expect((await lastOpened(page))!.loss).toEqual({ format: 'csv', items: [] });
  });

  test('S-FIO-20 a semicolon-delimited CSV asks first (it would import into one column)', async ({
    page,
  }) => {
    await sheet.goto();
    await openViaChooser(page, OPEN, fixture('semicolon.csv'));
    await expect(dialog(page).getByTestId('import-loss-list').locator('li')).toHaveText([
      /semicolon-separated/,
    ]);
    await dialog(page).getByRole('button', { name: 'Import anyway' }).click();
    await expect
      .poll(async () => (await lastOpened(page))?.loss.items.map((i) => i.id))
      .toEqual(['csv.delimiter']);
  });

  test("S-FIO-21 importConfirmation 'never' imports without asking", async ({ page }) => {
    await sheet.goto('?importConfirmation=never');
    await openViaChooser(page, OPEN, await lossyXlsx());
    await sheet.expectValue(0, 0, 'Frozen header');
    await expect(dialog(page)).toHaveCount(0);
  });

  test('S-FIO-22 a host callback decides instead of the dialog (headless API)', async ({
    page,
  }) => {
    await sheet.goto('?importConfirmation=host-no');
    await openViaChooser(page, OPEN, await lossyXlsx());
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as unknown as { __hostConfirmCalls?: number[] }).__hostConfirmCalls
        )
      )
      .toEqual([2]);
    await expect(dialog(page)).toHaveCount(0);
    await sheet.expectValue(0, 0, 'Product'); // the host said no

    await sheet.goto('?importConfirmation=host-yes');
    await openViaChooser(page, OPEN, await lossyXlsx());
    await sheet.expectValue(0, 0, 'Frozen header');
    await expect(dialog(page)).toHaveCount(0);
  });
});
