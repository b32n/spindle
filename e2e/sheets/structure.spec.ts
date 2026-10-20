import { test, expect } from '@playwright/test';
import { SheetsPage, MOD } from '../support/sheets';

test.describe('sheets: freeze, merge, sort, find & replace, context menus, tabs', () => {
  let sheet: SheetsPage;
  test.beforeEach(async ({ page }) => {
    sheet = new SheetsPage(page);
    await sheet.goto();
  });

  const sheetState = (page: import('@playwright/test').Page, expr: string) =>
    page.evaluate((e) => {
      const sh = (window as any).__spindle.workbook.getSheet();
      return new Function('sh', `return ${e}`)(sh);
    }, expr);

  test('S-FRZ-01 Freeze top row / first column / both / unfreeze', async ({ page }) => {
    await sheet.toolbar('Freeze panes').click();
    await page.getByText('Freeze top row', { exact: true }).click();
    await expect.poll(() => sheetState(page, 'sh.getFrozenRows()')).toBe(1);
    await sheet.toolbar('Freeze panes').click();
    await page.getByText('Freeze first column', { exact: true }).click();
    await expect
      .poll(() => sheetState(page, '[sh.getFrozenRows(), sh.getFrozenCols()]'))
      .toEqual([0, 1]);
    await sheet.toolbar('Freeze panes').click();
    await page.getByText('Freeze first row & column', { exact: true }).click();
    await expect
      .poll(() => sheetState(page, '[sh.getFrozenRows(), sh.getFrozenCols()]'))
      .toEqual([1, 1]);
  });

  test('S-FRZ-02 selection still works with frozen panes', async ({ page }) => {
    await sheet.toolbar('Freeze panes').click();
    await page.getByText('Freeze first row & column', { exact: true }).click();
    await sheet.select(2, 2);
    await expect(sheet.canvas).toHaveAttribute('data-active-cell', 'C3');
    await sheet.click(await sheet.colHeaderCenter(1));
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'B:B');
  });

  test('S-MRG-01 merge cells toggles merge and unmerge @smoke', async ({ page }) => {
    await sheet.drag(await sheet.cellCenter(5, 0), await sheet.cellCenter(5, 2));
    await sheet.toolbar('Merge cells').click();
    await expect.poll(() => sheetState(page, 'sh.getMergedRegions().length')).toBe(1);
    // Clicking inside the merge selects the whole region.
    await sheet.clickCell(5, 1);
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'A6:C6');
    await sheet.toolbar('Merge cells').click();
    await expect.poll(() => sheetState(page, 'sh.getMergedRegions().length')).toBe(0);
  });

  test('S-SRT-01 sort sheet A→Z / Z→A by active column', async ({ page }) => {
    await sheet.select(1, 1); // column B (Price): 10, 20, 15
    await sheet.toolbar(/Sort column/).click();
    await page.getByText('Sort sheet Z → A').click();
    await sheet.expectValue(1, 1, 20);
    await sheet.expectValue(2, 1, 15);
    await sheet.expectValue(3, 1, 10);
    await sheet.expectValue(1, 0, 'Gadget');
    await sheet.toolbar(/Sort column/).click();
    await page.getByText('Sort sheet A → Z').click();
    await sheet.expectValue(1, 1, 10);
    await sheet.expectValue(3, 1, 20);
    await sheet.expectValue(1, 0, 'Widget');
  });

  test('S-FND-01 find highlights matches and Next navigates; Escape-free close @smoke', async ({
    page,
  }) => {
    await sheet.select(0, 0);
    await page.keyboard.press(`${MOD}+f`);
    const dialog = page.getByPlaceholder('Find');
    await expect(dialog).toBeVisible();
    await dialog.fill('Gadget');
    await page.getByRole('button', { name: 'Next' }).click();
    await expect(sheet.canvas).toHaveAttribute('data-active-cell', 'A3');
    await page.getByRole('button', { name: 'Close' }).click();
    await expect(dialog).toBeHidden();
  });

  test('S-FND-02 replace and replace all', async ({ page }) => {
    await sheet.enter(5, 0, 'foo');
    await sheet.enter(6, 0, 'foo bar');
    await sheet.select(0, 0);
    await page.keyboard.press(`${MOD}+f`);
    await page.getByPlaceholder('Find').fill('foo');
    await page.getByPlaceholder('Replace with').fill('baz');
    await page.getByRole('button', { name: 'Replace all' }).click();
    await sheet.expectValue(5, 0, 'baz');
    await sheet.expectValue(6, 0, 'baz bar');
  });

  test('S-FND-03 match case and entire-cell options', async ({ page }) => {
    await sheet.enter(5, 0, 'Foo');
    await sheet.select(0, 0);
    await page.keyboard.press(`${MOD}+f`);
    await page.getByPlaceholder('Find').fill('foo');
    await page.getByPlaceholder('Replace with').fill('x');
    await page.getByLabel('Match case').check();
    await expect(page.getByRole('button', { name: 'Replace all' })).toBeDisabled();
    await page.getByLabel('Match case').uncheck();
    await page.getByLabel('Match entire cell contents').check();
    await page.getByRole('button', { name: 'Replace all' }).click();
    await sheet.expectValue(5, 0, 'x');
  });

  test('S-CTX-01 cell context menu: insert row above / delete row', async ({ page }) => {
    const p = await sheet.cellCenter(1, 0);
    await page.mouse.click(p.x, p.y, { button: 'right' });
    await page.getByText('Insert Row Above').click();
    await sheet.expectValue(1, 0, null);
    await sheet.expectValue(2, 0, 'Widget');
    const q = await sheet.cellCenter(1, 0);
    await page.mouse.click(q.x, q.y, { button: 'right' });
    await page.getByText('Delete Row', { exact: true }).click();
    await sheet.expectValue(1, 0, 'Widget');
  });

  test('S-CTX-02 cell context menu: insert column left / delete column', async ({ page }) => {
    const p = await sheet.cellCenter(0, 1);
    await page.mouse.click(p.x, p.y, { button: 'right' });
    await page.getByText('Insert Column Left').click();
    await sheet.expectValue(0, 1, null);
    await sheet.expectValue(0, 2, 'Price');
    const q = await sheet.cellCenter(0, 1);
    await page.mouse.click(q.x, q.y, { button: 'right' });
    await page.getByText('Delete Column', { exact: true }).click();
    await sheet.expectValue(0, 1, 'Price');
  });

  test('S-CTX-03 header context menu keeps selection and hides a column', async ({ page }) => {
    await sheet.click(await sheet.colHeaderCenter(1));
    const p = await sheet.colHeaderCenter(1);
    await page.mouse.click(p.x, p.y, { button: 'right' });
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'B:B');
    await page.getByText('Hide Column').click();
    await expect
      .poll(() => sheetState(page, 'sh.isColHidden ? sh.isColHidden(1) : true'))
      .toBe(true);
  });

  test('S-CTX-04 context menu Delete clears the selected cells', async ({ page }) => {
    await sheet.drag(await sheet.cellCenter(1, 0), await sheet.cellCenter(1, 1));
    const p = await sheet.cellCenter(1, 1);
    await page.mouse.click(p.x, p.y, { button: 'right' });
    await page.getByText('Delete', { exact: true }).first().click();
    await sheet.expectValue(1, 0, null);
    await sheet.expectValue(1, 1, null);
  });

  test('S-TAB-01 add a sheet tab and switch between tabs', async ({ page }) => {
    await page.getByRole('button', { name: '+', exact: true }).click();
    await expect(page.getByText('Sheet2', { exact: true })).toBeVisible();
    await sheet.select(0, 0);
    await sheet.expectValue(0, 0, null);
    await page.getByText('Sheet1', { exact: true }).click();
    await sheet.select(0, 0);
    await sheet.expectValue(0, 0, 'Product');
  });
});
