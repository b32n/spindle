import { test, expect, type Browser } from '@playwright/test';
import { SheetsPage } from '../support/sheets';

const WS = 'ws://127.0.0.1:1234';

async function peer(browser: Browser, room: string, user: string) {
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  const page = await context.newPage();
  const sheet = new SheetsPage(page);
  await sheet.goto(`?ws=${encodeURIComponent(WS)}&room=${room}&user=${user}`);
  return { context, page, sheet };
}

// Each test uses its own room so tests stay independent on the shared relay.
test.describe('sheets: two-client live collaboration', () => {
  test('S-COL-01 an edit made by one client appears in the other @smoke', async ({
    browser,
  }, info) => {
    const room = `sheets-${info.testId}`;
    const a = await peer(browser, room, 'Alice');
    const b = await peer(browser, room, 'Bob');
    await a.sheet.enter(6, 0, 'from alice');
    await b.sheet.expectValue(6, 0, 'from alice');
    await b.sheet.enter(7, 1, '42');
    await a.sheet.expectValue(7, 1, 42);
    await a.context.close();
    await b.context.close();
  });

  test('S-COL-02 concurrent edits to different cells converge', async ({ browser }, info) => {
    const room = `sheets-${info.testId}`;
    const a = await peer(browser, room, 'Alice');
    const b = await peer(browser, room, 'Bob');
    await Promise.all([a.sheet.enter(10, 0, 'A-side'), b.sheet.enter(11, 0, 'B-side')]);
    for (const p of [a, b]) {
      await p.sheet.expectValue(10, 0, 'A-side');
      await p.sheet.expectValue(11, 0, 'B-side');
    }
    await a.context.close();
    await b.context.close();
  });

  test('S-COL-03 concurrent edits to the same cell converge on one value', async ({
    browser,
  }, info) => {
    const room = `sheets-${info.testId}`;
    const a = await peer(browser, room, 'Alice');
    const b = await peer(browser, room, 'Bob');
    await Promise.all([a.sheet.enter(12, 2, 'alice'), b.sheet.enter(12, 2, 'bob')]);
    await expect
      .poll(async () => {
        const [x, y] = [await a.sheet.cellValue(12, 2), await b.sheet.cellValue(12, 2)];
        return x === y && (x === 'alice' || x === 'bob');
      })
      .toBe(true);
    await a.context.close();
    await b.context.close();
  });

  test('S-COL-04 formatting and formulas propagate', async ({ browser }, info) => {
    const room = `sheets-${info.testId}`;
    const a = await peer(browser, room, 'Alice');
    const b = await peer(browser, room, 'Bob');
    await a.sheet.enter(8, 0, '=B2+B3');
    await b.sheet.expectValue(8, 0, 30);
    await a.sheet.select(1, 0);
    await a.sheet.toolbar(/^Bold/).click();
    await b.sheet.expectBold(1, 0);
    // A remote edit to a precedent recalculates the formula on the other side.
    await b.sheet.enter(1, 1, '100');
    await a.sheet.expectValue(8, 0, 120);
    await a.context.close();
    await b.context.close();
  });

  test('S-COL-05 a late joiner receives existing state', async ({ browser }, info) => {
    const room = `sheets-${info.testId}`;
    const a = await peer(browser, room, 'Alice');
    await a.sheet.enter(9, 0, 'history');
    const late = await peer(browser, room, 'Carol');
    await late.sheet.expectValue(9, 0, 'history');
    // Demo seed data must not be duplicated or overwritten when joining.
    await late.sheet.expectValue(0, 0, 'Product');
    await a.context.close();
    await late.context.close();
  });

  test('S-COL-06 row/column clears propagate', async ({ browser }, info) => {
    const room = `sheets-${info.testId}`;
    const a = await peer(browser, room, 'Alice');
    const b = await peer(browser, room, 'Bob');
    await b.sheet.expectValue(1, 1, 10);
    await a.sheet.click(await a.sheet.colHeaderCenter(1));
    await a.page.keyboard.press('Delete');
    await b.sheet.expectValue(1, 1, null);
    await b.sheet.expectValue(3, 1, null);
    await b.sheet.expectValue(1, 0, 'Widget');
    await a.context.close();
    await b.context.close();
  });

  test('S-COL-07 separate rooms are isolated', async ({ browser }, info) => {
    const a = await peer(browser, `sheets-${info.testId}-1`, 'Alice');
    const b = await peer(browser, `sheets-${info.testId}-2`, 'Bob');
    await a.sheet.enter(6, 0, 'only in room 1');
    await a.sheet.expectValue(6, 0, 'only in room 1');
    await expect.poll(() => b.sheet.cellValue(6, 0)).toBe(null);
    await a.context.close();
    await b.context.close();
  });

  test('S-COL-08 the in-process two-pane demo mounts two independent editors', async ({ page }) => {
    const sheet = new SheetsPage(page);
    await sheet.goto();
    await page.getByRole('button', { name: 'Collab demo' }).click();
    const canvases = page.getByTestId('sheet-canvas');
    await expect(canvases).toHaveCount(2);
    // Type in the left pane; the right pane's selection state is independent.
    const box = (await canvases.first().boundingBox())!;
    await page.mouse.click(box.x + 100 + 50, box.y + 20 + 5 * 20 + 10);
    await page.keyboard.type('shared');
    await page.keyboard.press('Enter');
    await expect(canvases.first()).toHaveAttribute('data-active-cell', 'A7');
    await expect(canvases.nth(1)).toHaveAttribute('data-active-cell', '');
  });
});
