import { test, expect } from '@playwright/test';
import { SheetsPage, MOD } from '../support/sheets';

test.describe('sheets: filter, comments, hyperlinks, format dialog', () => {
  let sheet: SheetsPage;
  test.beforeEach(async ({ page }) => {
    sheet = new SheetsPage(page);
    await sheet.goto();
  });

  const rowHidden = (page: import('@playwright/test').Page, row: number) =>
    page.evaluate((r) => {
      const { workbook, FilterManager } = (window as any).__spindle;
      const sh = workbook.getSheet();
      return !FilterManager.getFilteredRows(sh, sh.getFilters()).has(r);
    }, row);

  test('S-FLT-01 filter a column by value hides non-matching rows; clear restores them @smoke', async ({
    page,
  }) => {
    await sheet.select(1, 1);
    await sheet.toolbar(/Filter column/).click();
    await expect(page.getByRole('heading', { name: /Filter Column/ })).toBeVisible();
    await page.getByRole('spinbutton', { name: 'Enter value' }).fill('20');
    await page.getByRole('button', { name: 'Apply Filter' }).click();
    await expect.poll(() => rowHidden(page, 1)).toBe(true); // 10 hidden
    await expect.poll(() => rowHidden(page, 2)).toBe(false); // 20 shown
    await expect.poll(() => rowHidden(page, 3)).toBe(true); // 15 hidden
    await sheet.toolbar(/Filter column/).click();
    await page.getByRole('button', { name: 'Clear Filter' }).click();
    await expect.poll(() => rowHidden(page, 1)).toBe(false);
    await expect.poll(() => rowHidden(page, 3)).toBe(false);
  });

  test('S-FLT-02 filter with Greater than', async ({ page }) => {
    await sheet.select(1, 1);
    await sheet.toolbar(/Filter column/).click();
    await page.getByRole('combobox').selectOption({ label: 'Greater than' });
    await page.getByRole('spinbutton').fill('12');
    await page.getByRole('button', { name: 'Apply Filter' }).click();
    await expect.poll(() => rowHidden(page, 1)).toBe(true);
    await expect.poll(() => rowHidden(page, 2)).toBe(false);
    await expect.poll(() => rowHidden(page, 3)).toBe(false);
  });

  test('S-FLT-03 Cancel leaves rows untouched', async ({ page }) => {
    await sheet.select(1, 1);
    await sheet.toolbar(/Filter column/).click();
    await page.getByRole('button', { name: 'Cancel' }).click();
    await expect(page.getByRole('heading', { name: /Filter Column/ })).toBeHidden();
    expect(await rowHidden(page, 1)).toBe(false);
  });

  test('S-CMT-01 add a comment from the panel @smoke', async ({ page }) => {
    await sheet.select(1, 1);
    await sheet.toolbar('Comments').click();
    await page.getByPlaceholder(/Add a comment/).fill('Looks good');
    await page.getByRole('button', { name: 'Comment', exact: true }).click();
    await expect(page.getByText('Looks good')).toBeVisible();
  });

  test('S-CMT-02 comments are per-cell; the All tab lists every thread', async ({ page }) => {
    await sheet.select(1, 1);
    await sheet.toolbar('Comments').click();
    await page.getByPlaceholder(/Add a comment/).fill('on B2');
    await page.getByRole('button', { name: 'Comment', exact: true }).click();
    await expect(page.getByText('on B2')).toBeVisible();
    await sheet.select(2, 1);
    await expect(page.getByText('on B2')).toBeHidden();
    await page.getByRole('button', { name: 'All', exact: true }).click();
    await expect(page.getByText('on B2')).toBeVisible();
  });

  test('S-CMT-03 Comments panel closes', async ({ page }) => {
    await sheet.select(1, 1);
    await sheet.toolbar('Comments').click();
    await expect(page.getByText('Comment on')).toBeVisible();
    await page.getByRole('button', { name: 'Close' }).click();
    await expect(page.getByText('Comment on')).toBeHidden();
  });

  test('S-LNK-01 insert a hyperlink on a cell', async ({ page }) => {
    await sheet.select(5, 0);
    await sheet.toolbar('Insert link').click();
    await page.getByPlaceholder('https://example.com').fill('https://example.org');
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).__spindle.workbook.getCell(undefined, 5, 0)?.hyperlink)
      )
      .toBe('https://example.org');
  });

  test('S-FMC-01 Format Cells dialog opens from the context menu and cancels', async ({ page }) => {
    const p = await sheet.cellCenter(1, 1);
    await page.mouse.click(p.x, p.y, { button: 'right' });
    await page.getByText('Format Cells...').click();
    await expect(page.getByRole('heading', { name: 'Format Cells' })).toBeVisible();
    await page.getByRole('button', { name: 'Cancel' }).click();
    await expect(page.getByRole('heading', { name: 'Format Cells' })).toBeHidden();
  });
});

test.describe('sheets: read-only mode', () => {
  let sheet: SheetsPage;
  test.beforeEach(async ({ page }) => {
    sheet = new SheetsPage(page);
    await sheet.goto();
    await page.getByRole('button', { name: /Read-only: OFF/ }).click();
    await expect(page.getByRole('button', { name: /Read-only: ON/ })).toBeVisible();
  });

  test('S-RO-01 typing does not change cells @smoke', async ({ page }) => {
    await sheet.select(1, 0);
    await page.keyboard.type('hacked');
    await page.keyboard.press('Enter');
    await sheet.expectValue(1, 0, 'Widget');
    await sheet.dblClickCell(1, 0);
    await page.keyboard.type('x');
    await page.keyboard.press('Enter');
    await sheet.expectValue(1, 0, 'Widget');
  });

  test('S-RO-02 Delete / cut / paste / format / undo are blocked', async ({ page }) => {
    await sheet.select(1, 0);
    await page.keyboard.press('Delete');
    await sheet.expectValue(1, 0, 'Widget');
    await page.keyboard.press(`${MOD}+x`);
    await sheet.expectValue(1, 0, 'Widget');
    await page.keyboard.press(`${MOD}+c`);
    await sheet.select(6, 0);
    await page.keyboard.press(`${MOD}+v`);
    await sheet.expectValue(6, 0, null);
    await page.keyboard.press(`${MOD}+b`);
    await sheet.expectBold(6, 0, false);
  });

  test('S-RO-03 formula bar is read-only and selection still works', async () => {
    await sheet.select(1, 0);
    await expect(sheet.formulaInput).toHaveJSProperty('readOnly', true);
    await sheet.click(await sheet.colHeaderCenter(1));
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'B:B');
  });

  test('S-RO-04 context menu is not offered and resize is blocked', async ({ page }) => {
    const p = await sheet.cellCenter(1, 0);
    await page.mouse.click(p.x, p.y, { button: 'right' });
    await expect(page.getByText('Insert Row Above')).toBeHidden();
    const edge = await sheet.point(200, 10);
    await sheet.drag(edge, { x: edge.x + 50, y: edge.y });
    expect(
      await page.evaluate(() => (window as any).__spindle.workbook.getSheet().getColWidth(0))
    ).toBe(100);
  });
});
