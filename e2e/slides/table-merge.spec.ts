import { test, expect } from '@playwright/test';
import { SlidesPage } from '../support/slides';

test.describe('slides: table merge and borders', () => {
  let s: SlidesPage;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
    await s.toolbar('Table (3×3)').click();
    await expect.poll(async () => (await s.els()).some((e) => e.type === 'table')).toBe(true);
  });

  const cell = (id: string, rc: string) =>
    s.stage.locator(`[data-element-id="${id}"] td[data-cell="${rc}"]`);

  async function centerOf(id: string, rc: string) {
    const b = await cell(id, rc).boundingBox();
    if (!b) throw new Error(`cell ${rc} not rendered`);
    return { x: b.x + b.width / 2, y: b.y + b.height / 2 };
  }

  test('SL-TBL-01 merging two cells draws one spanning cell and splitting restores them', async ({
    page,
  }) => {
    const table = (await s.els()).find((e) => e.type === 'table')!;
    await expect(s.stage.locator(`[data-element-id="${table.id}"] td`)).toHaveCount(9);
    const a = await centerOf(table.id, '0,0');
    await page.mouse.click(a.x, a.y); // selects the table
    const from = await centerOf(table.id, '0,0');
    const to = await centerOf(table.id, '0,1');
    await page.mouse.move(from.x, from.y);
    await page.mouse.down();
    await page.mouse.move(to.x, to.y, { steps: 6 });
    await page.mouse.up();
    await page.mouse.click(from.x, from.y, { button: 'right' });
    await page.getByText('Merge cells', { exact: true }).click();

    await expect(s.stage.locator(`[data-element-id="${table.id}"] td`)).toHaveCount(8);
    await expect(cell(table.id, '0,0')).toHaveAttribute('colspan', '2');
    await expect(cell(table.id, '0,1')).toHaveCount(0);
    const spans = await page.evaluate((i) => {
      const t = (window as any).__spindle.deck.getElement(i);
      return [t.cells[0][0].gridSpan, t.cells[0][1].hMerge];
    }, table.id);
    expect(spans).toEqual([2, true]);

    // Right-click inside the left half: the middle of the merged cell sits on a column resize handle.
    const box = (await cell(table.id, '0,0').boundingBox())!;
    await page.mouse.click(box.x + box.width * 0.25, box.y + box.height / 2, { button: 'right' });
    await page.getByText('Split cell', { exact: true }).click();
    await expect(s.stage.locator(`[data-element-id="${table.id}"] td`)).toHaveCount(9);
    expect(await cell(table.id, '0,0').getAttribute('colspan')).toBeNull();
  });

  test('SL-TBL-02 a border command styles the selected cell edges', async ({ page }) => {
    const table = (await s.els()).find((e) => e.type === 'table')!;
    const c = await centerOf(table.id, '1,1');
    await page.mouse.click(c.x, c.y);
    await page.mouse.click(c.x, c.y);
    await page.mouse.click(c.x, c.y, { button: 'right' });
    await page.getByText('No borders', { exact: true }).click();
    await expect
      .poll(() => cell(table.id, '1,1').evaluate((n) => getComputedStyle(n).borderTopStyle))
      .toBe('none');
    // The neighbour above agrees on the shared edge.
    await expect
      .poll(() => cell(table.id, '0,1').evaluate((n) => getComputedStyle(n).borderBottomStyle))
      .toBe('none');
  });
});
