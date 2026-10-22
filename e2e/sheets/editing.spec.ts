import { test, expect } from '@playwright/test';
import { SheetsPage, MOD } from '../support/sheets';

test.describe('sheets: editing, navigation & undo/redo', () => {
  let sheet: SheetsPage;
  test.beforeEach(async ({ page }) => {
    sheet = new SheetsPage(page);
    await sheet.goto();
  });

  test('S-EDT-01 sheet renders with demo data @smoke', async ({ page }) => {
    await expect(page.getByRole('heading', { name: 'Spindle Sheets Demo' })).toBeVisible();
    await expect(page.getByText('Sheet1', { exact: true })).toBeVisible();
    await sheet.expectValue(0, 0, 'Product');
    await sheet.expectValue(3, 3, 30);
  });

  test('S-EDT-02 typing in a selected cell and pressing Enter commits @smoke', async () => {
    await sheet.enter(5, 0, 'hello');
    await sheet.expectValue(5, 0, 'hello');
  });

  test('S-EDT-03 numbers are stored as numbers', async () => {
    await sheet.enter(5, 1, '42.5');
    await sheet.expectValue(5, 1, 42.5);
  });

  test('S-EDT-04 Escape cancels an edit', async () => {
    await sheet.select(5, 0);
    await sheet.type('discard', 'Escape');
    await sheet.expectValue(5, 0, null);
  });

  test('S-EDT-05 double-click edits an existing cell', async ({ page }) => {
    await sheet.dblClickCell(1, 0);
    await page.keyboard.press('End');
    await page.keyboard.type('X');
    await page.keyboard.press('Enter');
    await sheet.expectValue(1, 0, 'WidgetX');
  });

  test('S-EDT-06 Enter on a selected cell starts editing; F2-less editing keeps content', async ({
    page,
  }) => {
    await sheet.select(1, 0);
    await page.keyboard.press('Enter');
    await page.keyboard.press('End');
    await page.keyboard.type('!');
    await page.keyboard.press('Enter');
    await sheet.expectValue(1, 0, 'Widget!');
  });

  test('S-EDT-07 Tab commits and moves right', async ({ page }) => {
    await sheet.select(5, 0);
    await sheet.type('a', 'Tab');
    await sheet.type('b', 'Tab');
    await sheet.expectValue(5, 0, 'a');
    await sheet.expectValue(5, 1, 'b');
    await expect(sheet.canvas).toHaveAttribute('data-active-cell', 'C6');
    await page.keyboard.press('Shift+Tab');
    await expect(sheet.canvas).toHaveAttribute('data-active-cell', 'B6');
  });

  test('S-NAV-01 arrow keys move the active cell', async ({ page }) => {
    await sheet.select(2, 2);
    await page.keyboard.press('ArrowDown');
    await expect(sheet.canvas).toHaveAttribute('data-active-cell', 'C4');
    await page.keyboard.press('ArrowRight');
    await expect(sheet.canvas).toHaveAttribute('data-active-cell', 'D4');
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('ArrowLeft');
    await expect(sheet.canvas).toHaveAttribute('data-active-cell', 'C3');
  });

  test('S-NAV-02 arrow keys stop at the sheet edges', async ({ page }) => {
    await sheet.select(0, 0);
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('ArrowLeft');
    await expect(sheet.canvas).toHaveAttribute('data-active-cell', 'A1');
  });

  test('S-NAV-03 shift+arrows extend a range selection', async ({ page }) => {
    await sheet.select(1, 1);
    await page.keyboard.press('Shift+ArrowDown');
    await page.keyboard.press('Shift+ArrowRight');
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'B2:C3');
  });

  test('S-SEL-16 mouse drag selects a block; shift-click extends', async () => {
    await sheet.drag(await sheet.cellCenter(1, 1), await sheet.cellCenter(3, 2));
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'B2:C4');
    await sheet.select(0, 0);
    await sheet.clickCell(2, 2, { modifiers: ['Shift'] });
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'A1:C3');
  });

  test('S-EDT-08 typing into the formula bar edits the active cell', async ({ page }) => {
    await sheet.select(6, 0);
    await sheet.formulaInput.fill('from bar');
    await sheet.formulaInput.press('Enter');
    await sheet.expectValue(6, 0, 'from bar');
    await sheet.select(1, 0);
    await expect(sheet.formulaInput).toHaveValue('Widget');
    await expect(page.getByTestId('name-box')).toHaveText('A2');
  });

  test('S-UND-01 undo and redo via shortcuts and toolbar @smoke', async ({ page }) => {
    await sheet.enter(5, 0, 'one');
    await sheet.expectValue(5, 0, 'one');
    await page.keyboard.press(`${MOD}+z`);
    await sheet.expectValue(5, 0, null);
    await page.keyboard.press(`${MOD}+y`);
    await sheet.expectValue(5, 0, 'one');
    await sheet.toolbar(/^Undo/).click();
    await sheet.expectValue(5, 0, null);
    await sheet.toolbar(/^Redo/).click();
    await sheet.expectValue(5, 0, 'one');
  });

  test('S-UND-07 a host loop of 5,000 unbatched edits stays fast and is one undo step', async ({
    page,
  }) => {
    const ms = await page.evaluate(() => {
      const wb = (window as any).__spindle.workbook;
      const t = performance.now();
      for (let i = 0; i < 5000; i++) wb.setCellValue(undefined, i, 40, i);
      return performance.now() - t;
    });
    expect(ms).toBeLessThan(3000);
    await sheet.select(0, 0);
    await page.keyboard.press(`${MOD}+z`);
    await expect
      .poll(() =>
        page.evaluate(() => {
          const wb = (window as any).__spindle.workbook;
          return [wb.getCellValue(undefined, 0, 40), wb.getCellValue(undefined, 4999, 40)];
        })
      )
      .toEqual([null, null]);
  });

  test('S-UND-02 undo reverts formatting', async ({ page }) => {
    await sheet.select(1, 0);
    await page.keyboard.press(`${MOD}+b`);
    await sheet.expectBold(1, 0);
    await page.keyboard.press(`${MOD}+z`);
    await sheet.expectBold(1, 0, false);
  });

  test('S-UND-03 undo reverts a clear', async ({ page }) => {
    await sheet.select(1, 0);
    await page.keyboard.press('Delete');
    await sheet.expectValue(1, 0, null);
    await page.keyboard.press(`${MOD}+z`);
    await sheet.expectValue(1, 0, 'Widget');
  });
});
