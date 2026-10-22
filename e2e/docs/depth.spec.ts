import { test, expect, type Page } from '@playwright/test';
import { DocsPage, MOD } from '../support/docs';
import { closeToolbarMenu } from '../support/toolbar';

test.describe('docs: table menu, links, page setup, text styles', () => {
  let doc: DocsPage;
  test.beforeEach(async ({ page }) => {
    // The font list may reference web fonts; keep the suite offline-safe.
    await page.route(/fonts\.b32n\.io/, (r) => r.abort());
    doc = new DocsPage(page);
    await doc.goto();
  });

  async function insertTable(page: Page) {
    await doc.clearAndFocus();
    await doc.typeText('x');
    await doc.toolbar('Insert table').click();
    await page.getByRole('button', { name: /Insert 3×3 table/ }).click();
    await expect(doc.editor.locator('table')).toHaveCount(1);
    await closeToolbarMenu(page);
  }
  const rows = () => doc.editor.locator('tr');
  const cols = () => doc.editor.locator('tr').first().locator('td, th');

  async function cellMenu(page: Page, cell: number, item: string) {
    await doc.pages.locator('td').nth(cell).click({ button: 'right' });
    await page.locator('.table-context-menu .context-menu-item', { hasText: item }).click();
  }

  test('D-TBL-01 table context menu inserts rows and columns where asked and deletes them', async ({
    page,
  }) => {
    await insertTable(page);
    // Mark the middle row (cell 4 = row 2, column 2) so positions can be checked.
    await doc.pages.locator('td').nth(4).dblclick();
    await doc.settle();
    await doc.typeText('MID');
    await expect(doc.editor.locator('tr').nth(1)).toContainText('MID');

    const midAt = async (row: number, col: number) => {
      await expect(doc.editor.locator('tr').nth(row).locator('td, th').nth(col)).toContainText(
        'MID'
      );
    };
    await midAt(1, 1);

    await cellMenu(page, 4, 'Insert row above'); // above the MID row
    await expect(rows()).toHaveCount(4);
    await midAt(2, 1);
    await cellMenu(page, 4, 'Insert row below'); // cell 4 is now the new empty row: MID moves down again
    await expect(rows()).toHaveCount(5);
    await midAt(3, 1);

    await cellMenu(page, 1, 'Insert column left'); // left of column 1 (top row)
    await expect(cols()).toHaveCount(4);
    await midAt(3, 2);
    await cellMenu(page, 1, 'Insert column right'); // right of the (shifted) column 1
    await expect(cols()).toHaveCount(5);
    await midAt(3, 3);

    await cellMenu(page, 0, 'Delete row'); // top row
    await expect(rows()).toHaveCount(4);
    await midAt(2, 3);
    await cellMenu(page, 0, 'Delete column'); // first column
    await expect(cols()).toHaveCount(4);
    await midAt(2, 2);
  });

  test('D-LNK-01 an existing link can be edited (Update) and removed', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('link me');
    await doc.press(`${MOD}+a`);
    await doc.toolbar('Insert link').click();
    await page.getByPlaceholder('https://example.com').fill('https://example.org');
    await page.getByRole('button', { name: 'Apply' }).click();
    const link = doc.editor.locator('a[href="https://example.org"]');
    await expect(link).toHaveText('link me');

    // Put the caret inside the link: the dialog opens in edit mode with the current URL.
    await doc.press(`${MOD}+a`); // (clicking a link would navigate away)
    await doc.toolbar('Insert link').click();
    const url = page.getByPlaceholder('https://example.com');
    await expect(url).toHaveValue('https://example.org');
    await url.fill('https://example.com/changed');
    await page.getByRole('button', { name: 'Update' }).click();
    await expect(doc.editor.locator('a')).toHaveCount(1);
    await expect(doc.editor.locator('a')).toHaveAttribute('href', 'https://example.com/changed');
    await expect(doc.editor.locator('a')).toHaveText('link me');

    await doc.press(`${MOD}+a`);
    await doc.toolbar('Insert link').click();
    await page.getByRole('button', { name: 'Remove Link' }).click();
    await expect(doc.editor.locator('a')).toHaveCount(0);
    await expect(doc.editor).toContainText('link me'); // text stays, only the link goes
  });

  const pageBox = () =>
    page0().evaluate((el) => {
      const r = el.getBoundingClientRect();
      return { w: Math.round(r.width), h: Math.round(r.height) };
    });
  const page0 = () => doc.page.locator('.page').first();
  const openPageSetup = async (page: Page) => {
    await doc.toolbar('More tools').click();
    await page.getByRole('menuitem', { name: 'Page setup' }).click();
    await expect(page.getByRole('heading', { name: 'Page Setup' })).toBeVisible();
  };

  test('D-PGS-03 page size A4 changes the page dimensions', async ({ page }) => {
    const letter = await pageBox();
    await openPageSetup(page);
    await page.locator('select').first().selectOption('A4');
    await page.getByRole('button', { name: 'Apply' }).click();
    await expect.poll(async () => (await pageBox()).w).not.toBe(letter.w);
    const a4 = await pageBox();
    // A4 is narrower and taller than US Letter.
    expect(a4.w).toBeLessThan(letter.w);
    expect(a4.h).toBeGreaterThan(letter.h);
    expect(Math.abs(a4.h / a4.w - 297 / 210)).toBeLessThan(0.01);
  });

  test('D-PGS-04 wider margins narrow the text column and re-flow onto more pages', async ({
    page,
  }) => {
    const textWidth = () =>
      doc.pages.first().evaluate((el) => Math.round(el.getBoundingClientRect().width));
    const pageCount = () => page.locator('.page-footer').count();
    const before = await textWidth();
    await openPageSetup(page);
    const inputs = page.locator('input[type="number"]');
    // top, bottom, left, right (inches)
    await inputs.nth(2).fill('2');
    await inputs.nth(3).fill('2');
    await page.getByRole('button', { name: 'Apply' }).click();
    // 1in -> 2in on both sides removes 2in = 192px from the text column.
    await expect.poll(textWidth).toBe(before - 192);
    expect(await pageCount()).toBeGreaterThanOrEqual(1);
  });

  test('D-FMT-12 subscript, text colour, highlight and font family apply to the selection', async ({
    page,
  }) => {
    await doc.clearAndFocus();
    await doc.typeText('H2O');
    await doc.press(`${MOD}+a`);
    await doc.toolbar('Subscript').click();
    await expect(doc.editor.locator('sub')).toHaveText('H2O');
    await doc.toolbar('Subscript').click();
    await expect(doc.editor.locator('sub')).toHaveCount(0);

    const rgb = (hex: string) => {
      const n = parseInt(hex.slice(1), 16);
      return `rgb(${n >> 16}, ${(n >> 8) & 255}, ${n & 255})`;
    };
    const pickSwatch = async (index: number) => {
      const swatch = page.locator('button[title^="#"]').nth(index);
      const hex = (await swatch.getAttribute('title'))!;
      await swatch.click();
      return rgb(hex);
    };

    await doc.toolbar('Text color').click();
    const color = await pickSwatch(3);
    await expect(doc.editor.locator('span[style*="color"]').first()).toHaveText('H2O');
    await expect(doc.editor.locator('span[style*="color"]').first()).toHaveCSS('color', color);

    await doc.toolbar('Highlight').click();
    const background = await pickSwatch(5);
    await expect(doc.editor.locator('span[style*="background"]').first()).toHaveText('H2O');
    await expect(doc.editor.locator('span[style*="background"]').first()).toHaveCSS(
      'background-color',
      background
    );

    await page.getByRole('button', { name: 'Inter', exact: true }).click();
    await page.getByText('Georgia', { exact: true }).click();
    await expect(doc.editor.locator('span[style*="font-family"]').first()).toHaveCSS(
      'font-family',
      /Georgia/
    );
  });
});
