import { test, expect } from '@playwright/test';
import { SheetsPage } from '../support/sheets';

test.describe('mobile: sheets touch smoke', () => {
  test('M-SHEETS-01 tapping a cell selects it; typing edits it @mobile', async ({
    page,
    hasTouch,
  }) => {
    expect(hasTouch).toBe(true);
    const sheet = new SheetsPage(page);
    await sheet.goto();
    const p = await sheet.cellCenter(3, 1);
    await page.touchscreen.tap(p.x, p.y);
    await expect(sheet.canvas).toHaveAttribute('data-active-cell', 'B4');
    await sheet.type('tapped');
    await sheet.expectValue(3, 1, 'tapped');
    // Tapping another cell moves the selection and commits nothing extra.
    const q = await sheet.cellCenter(5, 2);
    await page.touchscreen.tap(q.x, q.y);
    await expect(sheet.canvas).toHaveAttribute('data-active-cell', 'C6');
    await sheet.expectValue(5, 2, null);
  });

  test('M-SHEETS-02 double-tapping a cell opens the editor @mobile', async ({ page }) => {
    const sheet = new SheetsPage(page);
    await sheet.goto();
    const p = await sheet.cellCenter(4, 0);
    await page.touchscreen.tap(p.x, p.y);
    await page.touchscreen.tap(p.x, p.y);
    await page.keyboard.type('dbl');
    await page.keyboard.press('Enter');
    await sheet.expectValue(4, 0, 'dbl');
  });
});
