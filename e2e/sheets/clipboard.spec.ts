import { test, expect } from '@playwright/test';
import { SheetsPage, MOD } from '../support/sheets';

test.describe('sheets: clipboard', () => {
  let sheet: SheetsPage;
  test.beforeEach(async ({ page }) => {
    sheet = new SheetsPage(page);
    await sheet.goto();
  });

  test('S-CLP-01 copy and paste a cell @smoke', async ({ page }) => {
    await sheet.select(1, 0);
    await page.keyboard.press(`${MOD}+c`);
    await sheet.select(6, 0);
    await page.keyboard.press(`${MOD}+v`);
    await sheet.expectValue(6, 0, 'Widget');
    await sheet.expectValue(1, 0, 'Widget');
  });

  test('S-CLP-02 cut moves content', async ({ page }) => {
    await sheet.select(1, 0);
    await page.keyboard.press(`${MOD}+x`);
    await sheet.expectValue(1, 0, null);
    await sheet.select(6, 0);
    await page.keyboard.press(`${MOD}+v`);
    await sheet.expectValue(6, 0, 'Widget');
  });

  test('S-CLP-03 copy and paste a range keeps its shape', async ({ page }) => {
    await sheet.drag(await sheet.cellCenter(1, 0), await sheet.cellCenter(2, 1));
    await page.keyboard.press(`${MOD}+c`);
    await sheet.select(8, 2);
    await page.keyboard.press(`${MOD}+v`);
    await sheet.expectValue(8, 2, 'Widget');
    await sheet.expectValue(8, 3, 10);
    await sheet.expectValue(9, 2, 'Gadget');
    await sheet.expectValue(9, 3, 20);
  });

  test('S-CLP-04 copying a whole column and pasting it elsewhere', async ({ page }) => {
    await sheet.click(await sheet.colHeaderCenter(0));
    await page.keyboard.press(`${MOD}+c`);
    await sheet.select(0, 6);
    await page.keyboard.press(`${MOD}+v`);
    await sheet.expectValue(0, 6, 'Product');
    await sheet.expectValue(3, 6, 'Thing');
  });

  test('S-CLP-05 copying a whole row and pasting it elsewhere', async ({ page }) => {
    await sheet.click(await sheet.rowHeaderCenter(1));
    await page.keyboard.press(`${MOD}+c`);
    await sheet.select(10, 0);
    await page.keyboard.press(`${MOD}+v`);
    await sheet.expectValue(10, 0, 'Widget');
    await sheet.expectValue(10, 3, 50);
  });

  test('S-CLP-06 pasting a formula rebases relative references', async ({ page }) => {
    await sheet.enter(5, 0, '=B2*2');
    await sheet.expectValue(5, 0, 20);
    await sheet.select(5, 0);
    await page.keyboard.press(`${MOD}+c`);
    await sheet.select(6, 0);
    await page.keyboard.press(`${MOD}+v`);
    await expect.poll(() => sheet.cellFormula(6, 0)).toBe('=B3*2');
    await sheet.expectValue(6, 0, 40);
  });

  test('S-CLP-07 paste is undoable', async ({ page }) => {
    await sheet.select(1, 0);
    await page.keyboard.press(`${MOD}+c`);
    await sheet.select(6, 0);
    await page.keyboard.press(`${MOD}+v`);
    await sheet.expectValue(6, 0, 'Widget');
    await page.keyboard.press(`${MOD}+z`);
    await sheet.expectValue(6, 0, null);
  });

  test('S-CLP-08 pasting plain tab-separated text from the system clipboard', async ({
    page,
    browserName,
  }) => {
    test.skip(browserName !== 'chromium', 'clipboard-read/write permissions are Chromium-only');
    await page.evaluate(() => navigator.clipboard.writeText('a\tb\n1\t2'));
    await sheet.select(12, 0);
    await page.keyboard.press(`${MOD}+v`);
    await sheet.expectValue(12, 0, 'a');
    await sheet.expectValue(12, 1, 'b');
    await sheet.expectValue(13, 0, 1);
    await sheet.expectValue(13, 1, 2);
  });

  test('S-CLP-09 copy writes TSV to the system clipboard', async ({ page, browserName }) => {
    test.skip(browserName !== 'chromium', 'clipboard-read/write permissions are Chromium-only');
    await sheet.drag(await sheet.cellCenter(0, 0), await sheet.cellCenter(1, 1));
    await page.keyboard.press(`${MOD}+c`);
    await expect
      .poll(() => page.evaluate(() => navigator.clipboard.readText()))
      .toBe('Product\tPrice\nWidget\t10');
  });

  test('S-CLP-10 context menu copy and paste', async ({ page }) => {
    await sheet.select(2, 0);
    const p = await sheet.cellCenter(2, 0);
    await page.mouse.click(p.x, p.y, { button: 'right' });
    await page.getByText('Copy (Ctrl+C)').click();
    const q = await sheet.cellCenter(7, 0);
    await page.mouse.click(q.x, q.y, { button: 'right' });
    await page.getByText(/^Paste/).click();
    await sheet.expectValue(7, 0, 'Gadget');
  });

  // https://github.com/b32n/spindle/issues/69
  test('S-CLP-11 paste preserves cell formatting', async ({ page }) => {
    await sheet.select(1, 0);
    await page.keyboard.press(`${MOD}+b`);
    await page.keyboard.press(`${MOD}+c`);
    await sheet.select(6, 0);
    await page.keyboard.press(`${MOD}+v`);
    await sheet.expectBold(6, 0);
  });

  test('S-CLP-12 paste carries number formats and colours, and the paste is one undo step', async ({
    page,
  }) => {
    await sheet.enter(8, 0, '1234.5');
    await sheet.select(8, 0);
    await page.evaluate(() => {
      const wb = (window as any).__spindle.workbook;
      const sheet = wb.getSheet();
      const cell = sheet.getCell(8, 0);
      const styleId = wb
        .getStylePool()
        .getOrCreate({ fontColor: '#ff0000', backgroundColor: '#ffff00' });
      const formatId = wb.getFormatPool().getOrCreate({ type: 'currency', decimals: 2 });
      wb.setCell(undefined, 8, 0, { ...cell, styleId, formatId });
    });
    await page.keyboard.press(`${MOD}+c`);
    await sheet.select(10, 1);
    await page.keyboard.press(`${MOD}+v`);
    await sheet.expectValue(10, 1, 1234.5);
    const style = await sheet.cellStyle(10, 1);
    expect(style.fontColor).toBe('#ff0000');
    expect(style.backgroundColor).toBe('#ffff00');
    expect(
      await page.evaluate(() => {
        const wb = (window as any).__spindle.workbook;
        const id = wb.getCell(undefined, 10, 1)?.formatId;
        return id ? wb.getFormatPool().get(id)?.type : null;
      })
    ).toBe('currency');
    await page.keyboard.press(`${MOD}+z`);
    await expect.poll(async () => (await sheet.cellStyle(10, 1)).fontColor ?? null).toBeNull();
  });
});
