import { test, expect } from '@playwright/test';
import { SheetsPage, MOD } from '../support/sheets';

// Row / column selection (S-SEL-*). Regression coverage for the bug where the
// header click was immediately collapsed back to a single cell.
test.describe('sheets: row & column selection', () => {
  let sheet: SheetsPage;
  test.beforeEach(async ({ page }) => {
    sheet = new SheetsPage(page);
    await sheet.goto();
  });

  test('S-SEL-01 clicking a column header selects the whole column @smoke', async () => {
    await sheet.click(await sheet.colHeaderCenter(1));
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'B:B');
    await expect(sheet.nameBox).toHaveText('B:B');
  });

  test('S-SEL-02 clicking a row header selects the whole row @smoke', async () => {
    await sheet.click(await sheet.rowHeaderCenter(2));
    await expect(sheet.canvas).toHaveAttribute('data-selection', '3:3');
    await expect(sheet.nameBox).toHaveText('3:3');
  });

  test('S-SEL-03 shift-click on column headers extends the selection', async () => {
    await sheet.click(await sheet.colHeaderCenter(1));
    await sheet.click(await sheet.colHeaderCenter(3), ['Shift']);
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'B:D');
  });

  test('S-SEL-04 shift-click on row headers extends the selection', async () => {
    await sheet.click(await sheet.rowHeaderCenter(1));
    await sheet.click(await sheet.rowHeaderCenter(4), ['Shift']);
    await expect(sheet.canvas).toHaveAttribute('data-selection', '2:5');
  });

  test('S-SEL-05 dragging across column headers selects a column range', async () => {
    await sheet.drag(await sheet.colHeaderCenter(0), await sheet.colHeaderCenter(2));
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'A:C');
  });

  test('S-SEL-06 dragging across row headers selects a row range', async () => {
    await sheet.drag(await sheet.rowHeaderCenter(0), await sheet.rowHeaderCenter(3));
    await expect(sheet.canvas).toHaveAttribute('data-selection', '1:4');
  });

  test('S-SEL-07 the select-all corner selects the entire sheet @smoke', async () => {
    await sheet.click(await sheet.corner());
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'A1:CV1000');
    const sel = await sheet.modelSelection();
    expect(sel.ranges[0]).toMatchObject({ startRow: 0, endRow: 999, startCol: 0, endCol: 99 });
  });

  test('S-SEL-08 Ctrl/Cmd+Space selects the active column', async ({ page }) => {
    await sheet.select(2, 2);
    await page.keyboard.press(`${MOD}+Space`);
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'C:C');
  });

  test('S-SEL-09 Shift+Space selects the active row', async ({ page }) => {
    await sheet.select(2, 2);
    await page.keyboard.press('Shift+Space');
    await expect(sheet.canvas).toHaveAttribute('data-selection', '3:3');
  });

  test('S-SEL-10 Ctrl/Cmd+A selects all cells', async ({ page }) => {
    await sheet.select(1, 1);
    await page.keyboard.press(`${MOD}+a`);
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'A1:CV1000');
  });

  test('S-SEL-11 clicking a cell after a column selection collapses to that cell', async () => {
    await sheet.click(await sheet.colHeaderCenter(1));
    await sheet.select(3, 2);
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'C4');
  });

  test('S-SEL-12 bold applies to every cell of a selected column', async ({ page }) => {
    await sheet.click(await sheet.colHeaderCenter(0));
    await page.keyboard.press(`${MOD}+b`);
    await sheet.expectBold(0, 0);
    await sheet.expectBold(3, 0);
    await sheet.expectBold(500, 0);
    await sheet.expectBold(1, 1, false);
  });

  test('S-SEL-13 the Bold toolbar button applies to a selected row', async () => {
    await sheet.click(await sheet.rowHeaderCenter(1));
    await sheet.toolbar(/^Bold/).click();
    await sheet.expectBold(1, 0);
    await sheet.expectBold(1, 3);
    await sheet.expectBold(1, 50);
    await sheet.expectBold(2, 0, false);
  });

  test('S-SEL-14 Delete clears every cell of a selected column', async ({ page }) => {
    await sheet.click(await sheet.colHeaderCenter(1));
    await page.keyboard.press('Delete');
    await sheet.expectValue(0, 1, null);
    await sheet.expectValue(3, 1, null);
    await sheet.expectValue(1, 0, 'Widget');
  });

  test('S-SEL-15 Backspace clears every cell of a selected row', async ({ page }) => {
    await sheet.click(await sheet.rowHeaderCenter(1));
    await page.keyboard.press('Backspace');
    await sheet.expectValue(1, 0, null);
    await sheet.expectValue(1, 3, null);
    await sheet.expectValue(2, 0, 'Gadget');
  });
});
