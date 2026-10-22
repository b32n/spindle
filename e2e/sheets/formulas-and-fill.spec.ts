import { test, expect } from '@playwright/test';
import { SheetsPage, COL_W, ROW_H, HEADER_H, HEADER_W } from '../support/sheets';

test.describe('sheets: formulas, autofill, resize', () => {
  let sheet: SheetsPage;
  test.beforeEach(async ({ page }) => {
    sheet = new SheetsPage(page);
    await sheet.goto();
  });

  test('S-FRM-01 simple arithmetic formula @smoke', async () => {
    await sheet.enter(5, 0, '=1+2*3');
    await sheet.expectValue(5, 0, 7);
    expect(await sheet.cellFormula(5, 0)).toBe('=1+2*3');
  });

  test('S-FRM-02 SUM over a range and live recalculation', async () => {
    await sheet.enter(5, 1, '=SUM(B2:B4)');
    await sheet.expectValue(5, 1, 45);
    await sheet.enter(1, 1, '100');
    await sheet.expectValue(5, 1, 135);
  });

  test('S-FRM-03 formula bar shows the formula, cell shows the value', async () => {
    await sheet.enter(5, 2, '=C2+C3');
    await sheet.select(5, 2);
    await expect(sheet.formulaInput).toHaveValue('=C2+C3');
    await sheet.expectValue(5, 2, 8);
  });

  test('S-FRM-04 errors surface in the cell', async () => {
    await sheet.enter(5, 0, '=1/0');
    await expect.poll(async () => String(await sheet.cellValue(5, 0))).toMatch(/#DIV\/0!/);
  });

  test('S-FRM-05 text functions', async () => {
    await sheet.enter(5, 0, '=UPPER(A2)&"-"&LEN(A3)');
    await sheet.expectValue(5, 0, 'WIDGET-6');
  });

  test('S-FRM-06 cross-sheet reference', async ({ page }) => {
    await page.getByRole('button', { name: '+', exact: true }).click();
    await sheet.enter(0, 0, '7');
    await page.getByText('Sheet1', { exact: true }).click();
    await sheet.enter(8, 0, '=Sheet2!A1*2');
    await sheet.expectValue(8, 0, 14);
  });

  test('S-FRM-07 clicking a cell while typing a formula inserts its reference', async ({
    page,
  }) => {
    await sheet.select(6, 0);
    await page.keyboard.type('=');
    await sheet.clickCell(1, 1);
    await expect(page.locator('input:focus')).toHaveValue('=B2');
    await page.keyboard.type('+');
    await expect(page.locator('input:focus')).toHaveValue('=B2+');
    await sheet.clickCell(2, 1);
    await page.keyboard.press('Enter');
    await expect.poll(() => sheet.cellFormula(6, 0)).toBe('=B2+B3');
    await sheet.expectValue(6, 0, 30);
  });

  test('S-FRM-08 formula autocomplete suggests functions', async ({ page }) => {
    await sheet.select(6, 0);
    await page.keyboard.type('=SU');
    await expect(page.getByText('SUM', { exact: true }).first()).toBeVisible();
    await page.keyboard.press('Escape');
  });

  test('S-FRM-09 array formulas spill into neighbouring cells', async () => {
    await sheet.enter(5, 5, '=SEQUENCE(3)');
    await sheet.expectValue(5, 5, 1);
    await sheet.expectValue(6, 5, 2);
    await sheet.expectValue(7, 5, 3);
  });

  // https://github.com/b32n/spindle/issues/66 (unit: formula-conformance.test.ts)
  test('S-FRM-10 ISBLANK of an empty cell is TRUE', async () => {
    await sheet.enter(5, 0, '=ISBLANK(F9)');
    await sheet.expectValue(5, 0, true);
  });

  // https://github.com/b32n/spindle/issues/67
  test('S-FRM-11 SUMPRODUCT of an array expression multiplies element-wise', async () => {
    for (const [i, v] of [1, 2, 3].entries()) {
      await sheet.enter(5 + i, 6, String(v));
      await sheet.enter(5 + i, 7, String(v + 3));
    }
    await sheet.enter(9, 0, '=SUMPRODUCT(G6:G8*H6:H8)');
    await sheet.expectValue(9, 0, 32);
    await sheet.enter(10, 0, '=SUMPRODUCT((G6:G8>1)*1)');
    await sheet.expectValue(10, 0, 2);
  });

  test('S-FRM-15 a blocked spill (#SPILL!) fills in by itself once the blocker is cleared', async ({
    page,
  }) => {
    await sheet.enter(6, 10, 'blocker');
    await sheet.enter(5, 10, '=SEQUENCE(3)');
    await sheet.expectValue(5, 10, '#SPILL!');
    await sheet.select(6, 10);
    await page.keyboard.press('Delete');
    await sheet.expectValue(5, 10, 1);
    await sheet.expectValue(6, 10, 2);
    await sheet.expectValue(7, 10, 3);
  });

  test('S-FRM-14 an operator over ranges spills element-wise', async () => {
    for (const [i, v] of [1, 2, 3].entries()) {
      await sheet.enter(5 + i, 6, String(v));
      await sheet.enter(5 + i, 7, String(v + 3));
    }
    await sheet.enter(12, 0, '=G6:G8*H6:H8');
    await sheet.expectValue(12, 0, 4);
    await sheet.expectValue(13, 0, 10);
    await sheet.expectValue(14, 0, 18);
  });

  // https://github.com/b32n/spindle/issues/68
  test('S-FRM-12 unary minus binds tighter than ^', async () => {
    await sheet.enter(5, 0, '=-2^2');
    await sheet.expectValue(5, 0, 4);
    await expect.poll(() => sheet.cellFormula(5, 0)).toBe('=-2^2');
  });

  // The stored formula text is re-rendered from the parsed formula; it must keep the parentheses
  // (it is what XLSX export and a reload re-parse).
  test('S-FRM-13 parentheses survive in the stored formula', async () => {
    await sheet.enter(5, 0, '=(1+2)*3');
    await sheet.expectValue(5, 0, 9);
    await expect.poll(() => sheet.cellFormula(5, 0)).toBe('=(1+2)*3');
    await sheet.enter(6, 0, '=-(1+2)');
    await sheet.expectValue(6, 0, -3);
    await expect.poll(() => sheet.cellFormula(6, 0)).toBe('=-(1+2)');
  });

  test('S-FIL-01 dragging the fill handle copies a value down', async ({ page }) => {
    await sheet.select(1, 1);
    const from = await sheet.point(HEADER_W + 2 * COL_W - 3, HEADER_H + 2 * ROW_H - 3);
    const to = await sheet.cellCenter(4, 1);
    await sheet.drag(from, to);
    await sheet.expectValue(2, 1, 10);
    await sheet.expectValue(4, 1, 10);
    void page;
  });

  test('S-FIL-02 fill extends a numeric series', async () => {
    await sheet.enter(6, 0, '1');
    await sheet.enter(7, 0, '2');
    await sheet.drag(await sheet.cellCenter(6, 0), await sheet.cellCenter(7, 0));
    const from = await sheet.point(HEADER_W + COL_W - 3, HEADER_H + 8 * ROW_H - 3);
    await sheet.drag(from, await sheet.cellCenter(10, 0));
    await sheet.expectValue(8, 0, 3);
    await sheet.expectValue(10, 0, 5);
  });

  test('S-FIL-03 fill rebases formulas', async () => {
    await sheet.enter(5, 0, '=B2*2');
    await sheet.select(5, 0);
    const from = await sheet.point(HEADER_W + COL_W - 3, HEADER_H + 6 * ROW_H - 3);
    await sheet.drag(from, await sheet.cellCenter(7, 0));
    await expect.poll(() => sheet.cellFormula(7, 0)).toBe('=B4*2');
    await sheet.expectValue(7, 0, 30);
  });

  test('S-RSZ-01 dragging a column header border resizes the column', async ({ page }) => {
    const width = () =>
      page.evaluate(() => (window as any).__spindle.workbook.getSheet().getColWidth(0));
    expect(await width()).toBe(100);
    const edge = await sheet.point(HEADER_W + COL_W, HEADER_H / 2);
    await sheet.drag(edge, { x: edge.x + 60, y: edge.y });
    await expect.poll(width).toBe(160);
  });

  test('S-RSZ-02 dragging a row header border resizes the row', async ({ page }) => {
    const height = () =>
      page.evaluate(() => (window as any).__spindle.workbook.getSheet().getRowHeight(0));
    const edge = await sheet.point(HEADER_W / 2, HEADER_H + ROW_H);
    await sheet.drag(edge, { x: edge.x, y: edge.y + 20 });
    await expect.poll(height).toBe(40);
  });

  test('S-RSZ-03 resize is undoable', async ({ page }) => {
    const width = () =>
      page.evaluate(() => (window as any).__spindle.workbook.getSheet().getColWidth(0));
    const edge = await sheet.point(HEADER_W + COL_W, HEADER_H / 2);
    await sheet.drag(edge, { x: edge.x + 60, y: edge.y });
    await expect.poll(width).toBe(160);
    await sheet.toolbar(/^Undo/).click();
    await expect.poll(width).toBe(100);
  });
});
