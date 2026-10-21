import { test, expect } from '@playwright/test';
import { SheetsPage } from '../support/sheets';
import { asUpload, collectPageErrors, openViaChooser } from '../support/file-io';
import { bigCsv } from '../support/generated';

const ROWS = 10_000;
// Generous budgets (>=10x the time measured locally (load ~0.3 s, SUM ~0.5 s, scroll ~0.1 s, edit ~0.3 s)): they
// catch "the grid became unusable", not small regressions. Measured values are in the PR notes.
const LOAD_BUDGET = 15_000;
const NAV_BUDGET = 8_000;
const EDIT_BUDGET = 8_000;

test.describe('sheets: large data', () => {
  test('S-BIG-01 a 10,000-row CSV opens, scrolls to the bottom and an edit near the end reflects within budget', async ({
    page,
  }, info) => {
    const errors = collectPageErrors(page);
    const sheet = new SheetsPage(page);
    await sheet.goto();

    const t0 = Date.now();
    await openViaChooser(
      page,
      'Open file (.xlsx, .csv)',
      asUpload({ fileName: 'big.csv', bytes: bigCsv(ROWS) }, 'text/csv')
    );
    await expect.poll(() => sheet.cellValue(ROWS, 0), { timeout: LOAD_BUDGET }).toBe(ROWS);
    const loadMs = Date.now() - t0;
    expect(await sheet.cellValue(ROWS, 1)).toBe(`item ${ROWS}`);
    expect(await sheet.cellValue(1, 1)).toBe('item 1');
    expect(await sheet.cellValue(ROWS + 1, 0)).toBeNull();

    // A formula over the whole column recalculates and is correct.
    const t1 = Date.now();
    await sheet.select(0, 6);
    await sheet.type(`=SUM(C2:C${ROWS + 1})`);
    let expected = 0;
    for (let i = 1; i <= ROWS; i++) expected += i % 97;
    await expect.poll(() => sheet.cellValue(0, 6), { timeout: NAV_BUDGET }).toBe(expected);
    const sumMs = Date.now() - t1;

    // Wheel-scroll to the very bottom and check the viewport followed: the cell under the pointer
    // near the bottom of the canvas is one of the last data rows.
    const t2 = Date.now();
    const box = (await sheet.canvas.boundingBox())!;
    await page.mouse.move(box.x + 300, box.y + 200);
    await page.mouse.wheel(0, 10_000_000);
    const bottomY = 640;
    let row = 0;
    await expect
      .poll(
        async () => {
          await page.mouse.click(box.x + 150, bottomY);
          const label = (await sheet.canvas.getAttribute('data-active-cell')) ?? '';
          row = Number(label.replace(/^[A-Z]+/, ''));
          return row;
        },
        { timeout: NAV_BUDGET, message: 'bottom-most visible row after scrolling to the end' }
      )
      .toBeGreaterThan(ROWS - 60);
    const navMs = Date.now() - t2;

    // Edit the bottom-most visible cell (column B) and make sure the model reflects it and the
    // neighbouring rows are intact.
    const t3 = Date.now();
    await page.mouse.click(box.x + 250, bottomY);
    await expect(sheet.canvas).toHaveAttribute('data-active-cell', `B${row}`);
    await sheet.type('edited at the end');
    await expect
      .poll(() => sheet.cellValue(row - 1, 1), { timeout: EDIT_BUDGET })
      .toBe('edited at the end');
    expect(await sheet.cellValue(row - 1, 0)).toBe(row - 1);
    expect(await sheet.cellValue(row - 2, 1)).toBe(`item ${row - 2}`);
    const editMs = Date.now() - t3;

    // Scrolling back to the top still works.
    await page.mouse.move(box.x + 300, box.y + 200);
    await page.mouse.wheel(0, -10_000_000);
    await expect
      .poll(async () => {
        await page.mouse.click(box.x + 150, box.y + 40);
        const label = (await sheet.canvas.getAttribute('data-active-cell')) ?? '';
        return Number(label.replace(/^[A-Z]+/, ''));
      })
      .toBeLessThan(6);

    info.annotations.push({
      type: 'timings',
      description: `load ${loadMs} ms, SUM ${sumMs} ms, scroll-to-end ${navMs} ms, edit ${editMs} ms`,
    });
    console.log(
      `[S-BIG-01] load ${loadMs} ms, SUM ${sumMs} ms, scroll-to-end ${navMs} ms, edit ${editMs} ms`
    );
    expect(errors, 'uncaught page errors').toEqual([]);
  });
});
