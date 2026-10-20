import { test, expect } from '@playwright/test';
import { SheetsPage, MOD } from '../support/sheets';

test.describe('sheets: formatting toolbar & shortcuts', () => {
  let sheet: SheetsPage;
  test.beforeEach(async ({ page }) => {
    sheet = new SheetsPage(page);
    await sheet.goto();
    await sheet.select(1, 1);
  });

  test('S-FMT-01 Bold button and Ctrl/Cmd+B toggle bold @smoke', async ({ page }) => {
    await sheet.toolbar(/^Bold/).click();
    await sheet.expectBold(1, 1);
    await sheet.toolbar(/^Bold/).click();
    await sheet.expectBold(1, 1, false);
    await sheet.canvas.focus();
    await page.keyboard.press(`${MOD}+b`);
    await sheet.expectBold(1, 1);
  });

  test('S-FMT-02 Italic via button and shortcut', async ({ page }) => {
    await sheet.toolbar(/^Italic/).click();
    await expect.poll(async () => (await sheet.cellStyle(1, 1)).italic).toBe(true);
    await sheet.canvas.focus();
    await page.keyboard.press(`${MOD}+i`);
    await expect.poll(async () => !!(await sheet.cellStyle(1, 1)).italic).toBe(false);
  });

  test('S-FMT-03 Underline via button and shortcut', async ({ page }) => {
    await sheet.toolbar(/^Underline/).click();
    await expect.poll(async () => (await sheet.cellStyle(1, 1)).textDecoration).toBe('underline');
    await sheet.canvas.focus();
    await page.keyboard.press(`${MOD}+u`);
    await expect.poll(async () => (await sheet.cellStyle(1, 1)).textDecoration).toBe('none');
  });

  test('S-FMT-04 Strikethrough button', async () => {
    await sheet.toolbar('Strikethrough').click();
    await expect.poll(async () => (await sheet.cellStyle(1, 1)).textDecoration).toBe('line-through');
  });

  test('S-FMT-05 horizontal alignment buttons', async () => {
    await sheet.toolbar('Align center').click();
    await expect.poll(async () => (await sheet.cellStyle(1, 1)).textAlign).toBe('center');
    await sheet.toolbar('Align right').click();
    await expect.poll(async () => (await sheet.cellStyle(1, 1)).textAlign).toBe('right');
    await sheet.toolbar('Align left').click();
    await expect.poll(async () => (await sheet.cellStyle(1, 1)).textAlign).toBe('left');
  });

  test('S-FMT-06 wrap text toggles', async () => {
    await sheet.toolbar('Wrap text').click();
    await expect.poll(async () => (await sheet.cellStyle(1, 1)).textWrap).toBe(true);
  });

  test('S-FMT-07 font size increase / decrease', async () => {
    await sheet.toolbar('Increase font size').click();
    await expect.poll(async () => (await sheet.cellStyle(1, 1)).fontSize).toBe(12);
    await sheet.toolbar('Decrease font size').click();
    await sheet.toolbar('Decrease font size').click();
    await expect.poll(async () => (await sheet.cellStyle(1, 1)).fontSize).toBe(10);
  });

  test('S-FMT-08 currency / percentage formats change the stored format', async ({ page }) => {
    const formatType = () =>
      page.evaluate(() => {
        const wb = (window as any).__spindle.workbook;
        const cell = wb.getCell(undefined, 1, 1);
        return cell?.formatId ? wb.getFormatPool().get(cell.formatId)?.type : null;
      });
    await sheet.toolbar('Currency format').click();
    await expect.poll(formatType).toBe('currency');
    await sheet.toolbar('Percentage format').click();
    await expect.poll(formatType).toBe('percentage');
  });

  test('S-FMT-09 formatting applies to a multi-cell range', async () => {
    await sheet.drag(await sheet.cellCenter(1, 0), await sheet.cellCenter(2, 1));
    await sheet.toolbar(/^Bold/).click();
    for (const [r, c] of [[1, 0], [1, 1], [2, 0], [2, 1]]) await sheet.expectBold(r, c);
    await sheet.expectBold(3, 0, false);
  });

  test('S-FMT-10 formatting is reverted by undo', async ({ page }) => {
    await sheet.toolbar(/^Bold/).click();
    await sheet.expectBold(1, 1);
    await sheet.toolbar(/^Undo/).click();
    await sheet.expectBold(1, 1, false);
    await page.keyboard.press(`${MOD}+y`);
    await sheet.expectBold(1, 1);
  });
});
