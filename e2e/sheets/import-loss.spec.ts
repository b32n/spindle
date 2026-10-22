import { test, expect, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { SheetsPage } from '../support/sheets';
import { openViaChooser, fixture } from '../support/file-io';

const OPEN = 'Open file (.xlsx, .csv)';

async function lastOpened(page: Page) {
  return page.evaluate(
    () =>
      (
        window as unknown as {
          __lastOpened?: {
            fileName: string;
            loss: { format: string; items: Array<{ id: string; count: number }> };
          };
        }
      ).__lastOpened
  );
}

async function xlsxWith(
  extra: (zip: JSZip) => Promise<void> | void
): Promise<{ name: string; mimeType: string; buffer: Buffer }> {
  const zip = await JSZip.loadAsync(
    (await import('node:fs')).readFileSync(fixture('exceljs-frozen.xlsx'))
  );
  await extra(zip);
  return {
    name: 'features.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: Buffer.from(await zip.generateAsync({ type: 'uint8array' })),
  };
}

test.describe('sheets: import loss report', () => {
  let sheet: SheetsPage;
  test.beforeEach(async ({ page }) => {
    await page.route(/fonts\.b32n\.io/, (r) => r.abort());
    sheet = new SheetsPage(page);
    await sheet.goto();
  });

  test('S-FIO-16 opening a workbook with unsupported features reports what was not imported', async ({
    page,
  }) => {
    const file = await xlsxWith(async (zip) => {
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
    });
    await openViaChooser(page, OPEN, file);
    await sheet.expectValue(0, 0, 'Frozen header');
    await expect
      .poll(async () => (await lastOpened(page))?.loss.items.map((i) => i.id).sort())
      .toEqual(['xlsx.charts', 'xlsx.conditionalFormatting']);
  });

  test('S-FIO-17 opening a plain workbook and a comma CSV reports no loss', async ({ page }) => {
    await openViaChooser(page, OPEN, fixture('exceljs-frozen.xlsx'));
    await sheet.expectValue(0, 0, 'Frozen header');
    await expect.poll(async () => (await lastOpened(page))?.fileName).toBe('exceljs-frozen.xlsx');
    expect((await lastOpened(page))!.loss.items).toEqual([]);
    await openViaChooser(page, OPEN, fixture('quoting.csv'));
    await expect.poll(async () => (await lastOpened(page))?.fileName).toBe('quoting.csv');
    expect((await lastOpened(page))!.loss).toEqual({ format: 'csv', items: [] });
  });

  test('S-FIO-18 a semicolon-delimited CSV is flagged (it imports into one column)', async ({
    page,
  }) => {
    await openViaChooser(page, OPEN, fixture('semicolon.csv'));
    await expect
      .poll(async () => (await lastOpened(page))?.loss.items.map((i) => i.id))
      .toEqual(['csv.delimiter']);
  });
});
