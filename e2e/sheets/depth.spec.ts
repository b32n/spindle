import { test, expect, type Page } from '@playwright/test';
import { SheetsPage, MOD } from '../support/sheets';

/** The stored CellFormat (number/date/custom) of a cell, or null. */
const cellFormat = (page: Page, row: number, col: number) =>
  page.evaluate(
    ([r, c]) => {
      const wb = (window as any).__spindle.workbook;
      const cell = wb.getCell(undefined, r, c);
      return cell?.formatId ? (wb.getFormatPool().get(cell.formatId) ?? null) : null;
    },
    [row, col]
  );

async function openFormatCells(page: Page, sheet: SheetsPage, row: number, col: number) {
  const p = await sheet.cellCenter(row, col);
  await page.mouse.click(p.x, p.y, { button: 'right' });
  await page.getByText('Format Cells...').click();
  await expect(page.getByRole('heading', { name: 'Format Cells' })).toBeVisible();
}

test.describe('sheets: Format Cells dialog, colours, borders, font', () => {
  let sheet: SheetsPage;
  test.beforeEach(async ({ page }) => {
    // The font picker loads Google Fonts on demand; keep the suite offline-safe.
    await page.route(/fonts\.(googleapis|gstatic)\.com/, (r) => r.abort());
    sheet = new SheetsPage(page);
    await sheet.goto();
  });

  test('S-FMC-02 Number category: decimal places apply to the cell and the preview shows them', async ({
    page,
  }) => {
    await sheet.enter(8, 0, '1234.5');
    await openFormatCells(page, sheet, 8, 0);
    const decimals = page.locator('input[type="number"]').first();
    await decimals.fill('3');
    // The dialog previews the sample value with the chosen format before applying.
    await expect(page.getByText('1,234.500').or(page.getByText('1234.500')).first()).toBeVisible();
    await page.getByRole('button', { name: 'Apply' }).click();
    await expect(page.getByRole('heading', { name: 'Format Cells' })).toBeHidden();
    await expect
      .poll(() => cellFormat(page, 8, 0))
      .toMatchObject({ type: 'number', decimalPlaces: 3 });
    await sheet.expectValue(8, 0, 1234.5); // formatting never changes the stored value
  });

  test('S-FMC-03 Date/Time category: a date pattern is stored on the cell', async ({ page }) => {
    await sheet.enter(8, 0, '45000');
    await openFormatCells(page, sheet, 8, 0);
    await page.getByRole('button', { name: 'Date/Time', exact: true }).click();
    await page.locator('select').nth(1).selectOption('YYYY-MM-DD');
    await expect(page.getByText('2023-03-15', { exact: true })).toBeVisible(); // serial 45000
    await page.getByRole('button', { name: 'Apply' }).click();
    await expect
      .poll(() => cellFormat(page, 8, 0))
      .toMatchObject({ type: 'date', dateFormat: 'YYYY-MM-DD' });
    await sheet.expectValue(8, 0, 45000);
  });

  test('S-FMC-04 Custom category: a supported pattern is stored and previewed', async ({
    page,
  }) => {
    await sheet.enter(8, 0, '1234.5');
    await openFormatCells(page, sheet, 8, 0);
    await page.getByRole('button', { name: 'Custom', exact: true }).click();
    await page.getByPlaceholder(/Enter format pattern/).fill('#,##0.00');
    await expect(page.getByText('1,234.50', { exact: true })).toBeVisible();
    await page.getByRole('button', { name: 'Apply' }).click();
    await expect
      .poll(() => cellFormat(page, 8, 0))
      .toMatchObject({ type: 'custom', pattern: '#,##0.00' });
  });

  // Bug: Custom number formats are a stub: only #,##0.00, 0.0% and MM/DD/YYYY are recognised, any other pattern prints the pattern with every 0 replaced by the value, e.g. 7.77" pts" (https://github.com/b32n/spindle/issues/83).
  test.fixme('S-FMC-04b Custom category: a pattern with a literal suffix renders 7.00 pts', async ({
    page,
  }) => {
    await sheet.enter(8, 0, '7');
    await openFormatCells(page, sheet, 8, 0);
    await page.getByRole('button', { name: 'Custom', exact: true }).click();
    await page.getByPlaceholder(/Enter format pattern/).fill('0.00" pts"');
    await expect(page.getByText('7.00 pts', { exact: true })).toBeVisible();
  });

  test('S-FMC-05 Format Cells applies to the whole selected range, and undo reverts it', async ({
    page,
  }) => {
    await sheet.enter(8, 0, '1');
    await sheet.enter(9, 0, '2');
    await sheet.drag(await sheet.cellCenter(8, 0), await sheet.cellCenter(9, 0));
    await openFormatCells(page, sheet, 9, 0);
    await page.locator('input[type="number"]').first().fill('4');
    await page.getByRole('button', { name: 'Apply' }).click();
    for (const r of [8, 9])
      await expect.poll(() => cellFormat(page, r, 0)).toMatchObject({ decimalPlaces: 4 });
    await sheet.focusGrid();
    await page.keyboard.press(`${MOD}+z`);
    for (const r of [8, 9]) await expect.poll(() => cellFormat(page, r, 0)).toBeNull();
  });

  test('S-CLR-01 fill colour and text colour are stored on the whole selection', async ({
    page,
  }) => {
    await sheet.drag(await sheet.cellCenter(8, 0), await sheet.cellCenter(8, 1));
    await sheet.toolbar('Fill color').click();
    await page.getByTitle('#ff0000', { exact: true }).first().click();
    await sheet.toolbar('Text color').click();
    await page.getByTitle('#0000ff', { exact: true }).first().click();
    for (const c of [0, 1]) {
      await expect.poll(async () => (await sheet.cellStyle(8, c)).backgroundColor).toBe('#ff0000');
      await expect.poll(async () => (await sheet.cellStyle(8, c)).fontColor).toBe('#0000ff');
    }
    expect((await sheet.cellStyle(9, 0)).backgroundColor).toBeUndefined();
  });

  test('S-BRD-01 border menu: all / bottom / none set and clear the cell borders', async ({
    page,
  }) => {
    await sheet.select(8, 0);
    await sheet.toolbar('Borders').click();
    await page.getByText('All borders', { exact: true }).click();
    await expect
      .poll(async () => {
        const s = await sheet.cellStyle(8, 0);
        return [s.borderTop, s.borderRight, s.borderBottom, s.borderLeft].every(Boolean);
      })
      .toBe(true);
    await sheet.toolbar('Borders').click();
    await page.getByText('No border', { exact: true }).click();
    await expect
      .poll(async () => {
        const s = await sheet.cellStyle(8, 0);
        return [s.borderTop, s.borderRight, s.borderBottom, s.borderLeft].some(Boolean);
      })
      .toBe(false);
    await sheet.select(9, 0);
    await sheet.toolbar('Borders').click();
    await page.getByText('Bottom border', { exact: true }).click();
    await expect.poll(async () => !!(await sheet.cellStyle(9, 0)).borderBottom).toBe(true);
    const s = await sheet.cellStyle(9, 0);
    expect(s.borderTop || s.borderLeft || s.borderRight).toBeFalsy();
  });

  test('S-FNT-01 font family from the picker is stored on the cell', async ({ page }) => {
    await sheet.select(8, 0);
    await page.getByRole('button', { name: /^Arial/ }).click();
    await page.getByPlaceholder('Search fonts…').fill('Lato');
    await page.getByText('Lato', { exact: true }).first().click();
    await expect.poll(async () => (await sheet.cellStyle(8, 0)).fontFamily).toBe('Lato');
  });
});

const sheetEval = <T>(page: Page, expr: string): Promise<T> =>
  page.evaluate(
    (e) =>
      new Function('wb', `const sh = wb.getSheet(); return (${e});`)(
        (window as any).__spindle.workbook
      ),
    expr
  );

test.describe('sheets: tabs, merged cells, structural undo', () => {
  let sheet: SheetsPage;
  test.beforeEach(async ({ page }) => {
    sheet = new SheetsPage(page);
    await sheet.goto();
  });
  const sheetNames = (page: Page) =>
    page.evaluate(() =>
      [...(window as any).__spindle.workbook.sheets.values()].map((s: any) => s.name)
    );

  test('S-TAB-02 double-click renames a sheet; Escape cancels; the data stays', async ({
    page,
  }) => {
    await page.getByText('Sheet1', { exact: true }).dblclick();
    const input = page.locator('.sheet-tabs input');
    await input.fill('Budget');
    await input.press('Escape');
    expect(await sheetNames(page)).toEqual(['Sheet1']);
    await page.getByText('Sheet1', { exact: true }).dblclick();
    await page.locator('.sheet-tabs input').fill('Budget');
    await page.locator('.sheet-tabs input').press('Enter');
    await expect.poll(() => sheetNames(page)).toEqual(['Budget']);
    await expect(page.getByText('Budget', { exact: true })).toBeVisible();
    await sheet.expectValue(1, 0, 'Widget');
  });

  test('S-TAB-03 deleting a sheet removes it and falls back to the remaining one; the last sheet cannot be deleted', async ({
    page,
  }) => {
    await expect(page.locator('.sheet-tabs button').filter({ hasText: '×' })).toHaveCount(0);
    await page.getByRole('button', { name: '+', exact: true }).click();
    await expect.poll(() => sheetNames(page)).toEqual(['Sheet1', 'Sheet2']);
    await sheet.select(0, 0);
    await sheet.type('only on sheet 2');
    const tab2 = page.locator('.sheet-tabs > div').filter({ hasText: 'Sheet2' }).first();
    await tab2.locator('button').click();
    await expect.poll(() => sheetNames(page)).toEqual(['Sheet1']);
    await sheet.select(0, 0);
    await sheet.expectValue(0, 0, 'Product');
    await expect(page.locator('.sheet-tabs button').filter({ hasText: '×' })).toHaveCount(0);
  });

  test('S-MRG-02 editing and selecting through a merged region uses its top-left cell', async ({
    page,
  }) => {
    await sheet.drag(await sheet.cellCenter(6, 0), await sheet.cellCenter(7, 1));
    await sheet.toolbar('Merge cells').click();
    await expect.poll(() => sheetEval<number>(page, 'sh.getMergedRegions().length')).toBe(1);
    // Clicking the bottom-right cell of the region selects the whole region (A7:B8) ...
    await sheet.clickCell(7, 1);
    await expect(sheet.canvas).toHaveAttribute('data-selection', 'A7:B8');
    // ... and typing writes to the top-left cell only.
    await sheet.type('merged');
    await sheet.expectValue(6, 0, 'merged');
    await sheet.expectValue(7, 1, null);
    await sheet.expectValue(6, 1, null);
    // Merged regions are one unit for keyboard navigation too: Down from the region skips it.
    await sheet.clickCell(6, 0);
    await page.keyboard.press('ArrowDown');
    await expect(sheet.canvas).toHaveAttribute('data-active-cell', 'A9');
    // Unmerge keeps the value in the top-left cell.
    await sheet.clickCell(6, 0);
    await sheet.toolbar('Merge cells').click();
    await expect.poll(() => sheetEval<number>(page, 'sh.getMergedRegions().length')).toBe(0);
    await sheet.expectValue(6, 0, 'merged');
  });

  test('S-UND-01 undo / redo of insert row, delete row and merge', async ({ page }) => {
    const rows = () => sheetEval<number>(page, 'sh.rowCount');
    const before = await rows();
    // insert row above row 2 (index 1)
    let p = await sheet.cellCenter(1, 0);
    await page.mouse.click(p.x, p.y, { button: 'right' });
    await page.getByText('Insert Row Above').click();
    await sheet.expectValue(1, 0, null);
    await sheet.expectValue(2, 0, 'Widget');
    await sheet.focusGrid();
    await page.keyboard.press(`${MOD}+z`);
    await sheet.expectValue(1, 0, 'Widget');
    await expect.poll(rows).toBe(before);
    await page.keyboard.press(`${MOD}+y`);
    await sheet.expectValue(2, 0, 'Widget');

    // delete a row, then undo brings the data back in place
    await page.keyboard.press(`${MOD}+z`); // back to the original layout
    await sheet.expectValue(1, 0, 'Widget');
    p = await sheet.cellCenter(1, 0);
    await page.mouse.click(p.x, p.y, { button: 'right' });
    await page.getByText('Delete Row', { exact: true }).click();
    await sheet.expectValue(1, 0, 'Gadget');
    await sheet.focusGrid();
    await page.keyboard.press(`${MOD}+z`);
    await sheet.expectValue(1, 0, 'Widget');
    await sheet.expectValue(2, 0, 'Gadget');
    await expect.poll(rows).toBe(before);

    // merge, then undo
    await sheet.drag(await sheet.cellCenter(8, 0), await sheet.cellCenter(8, 2));
    await sheet.toolbar('Merge cells').click();
    await expect.poll(() => sheetEval<number>(page, 'sh.getMergedRegions().length')).toBe(1);
    await sheet.focusGrid();
    await page.keyboard.press(`${MOD}+z`);
    await expect.poll(() => sheetEval<number>(page, 'sh.getMergedRegions().length')).toBe(0);
  });

  test('S-UND-02 undo of a sort restores the original order', async ({ page }) => {
    await sheet.select(1, 1); // Price column: 10, 20, 15
    const column = async () => [
      await sheet.cellValue(1, 1),
      await sheet.cellValue(2, 1),
      await sheet.cellValue(3, 1),
    ];
    const original = await column();
    await sheet.toolbar(/Sort column/).click();
    await page.getByText('Sort sheet Z → A').click();
    await expect.poll(column).toEqual([...original].sort((a, b) => (b as number) - (a as number)));
    await sheet.focusGrid();
    await page.keyboard.press(`${MOD}+z`);
    await expect.poll(column).toEqual(original);
    await page.keyboard.press(`${MOD}+y`);
    await expect.poll(column).not.toEqual(original);
  });

  test('S-UND-03 undo of a column filter shows the hidden rows again', async ({ page }) => {
    const hidden = (row: number) =>
      page.evaluate((r) => {
        const { workbook, FilterManager } = (window as any).__spindle;
        const sh = workbook.getSheet();
        return !FilterManager.getFilteredRows(sh, sh.getFilters()).has(r);
      }, row);
    await sheet.select(1, 1);
    await sheet.toolbar(/Filter column/).click();
    await page.getByRole('spinbutton', { name: 'Enter value' }).fill('20');
    await page.getByRole('button', { name: 'Apply Filter' }).click();
    await expect.poll(() => hidden(1)).toBe(true);
    await sheet.focusGrid();
    await page.keyboard.press(`${MOD}+z`);
    await expect.poll(() => hidden(1)).toBe(false);
    await expect.poll(() => hidden(3)).toBe(false);
    await page.keyboard.press(`${MOD}+y`);
    await expect.poll(() => hidden(1)).toBe(true);
  });
});
