import { test, expect } from '@playwright/test';
import { DocsPage, MOD } from '../support/docs';

// 1x1 transparent PNG
const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';

test.describe('docs: insert, dialogs, comments, pagination, zoom', () => {
  let doc: DocsPage;
  test.beforeEach(async ({ page }) => {
    doc = new DocsPage(page);
    await doc.goto();
  });

  test('D-INS-01 insert a link on selected text', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('link me');
    await doc.press(`${MOD}+a`);
    await doc.toolbar('Insert link').click();
    await page.getByPlaceholder('https://example.com').fill('https://example.org');
    await page.getByRole('button', { name: 'Apply' }).click();
    await expect(doc.editor.locator('a[href="https://example.org"]')).toHaveText('link me');
  });

  test('D-INS-02 link dialog can be cancelled', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('plain');
    await doc.press(`${MOD}+a`);
    await doc.toolbar('Insert link').click();
    await page.getByRole('button', { name: 'Cancel' }).click();
    await expect(page.getByPlaceholder('https://example.com')).toBeHidden();
    await expect(doc.editor.locator('a')).toHaveCount(0);
  });

  test('D-INS-03 insert an image from a URL', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('x');
    await doc.toolbar('Insert image').click();
    await page.getByRole('button', { name: 'By URL' }).click();
    await page.getByPlaceholder('https://example.com/image.png').fill(PNG);
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await expect(doc.editor.locator('img')).toHaveCount(1);
  });

  test('D-INS-04 insert a table from the size picker', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('x');
    await doc.toolbar('Insert table').click();
    await page.getByRole('button', { name: /Insert 3×3 table/ }).click();
    await expect(doc.editor.locator('table')).toHaveCount(1);
    await expect(doc.editor.locator('tr')).toHaveCount(3);
    await expect(doc.editor.locator('tr').first().locator('td, th')).toHaveCount(3);
  });

  test('D-INS-05 double-clicking a table cell lets you type into it', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('x');
    await doc.toolbar('Insert table').click();
    await page.getByRole('button', { name: /Insert 3×3 table/ }).click();
    await expect(doc.editor.locator('table')).toHaveCount(1);
    // The caret is placed after the table; double-click the first cell to edit it
    // (a single click selects the cell without entering it).
    await doc.pages.locator('td').first().dblclick();
    await expect(doc.editor).toBeFocused();
    await doc.settle();
    await doc.typeText('A1');
    await expect(doc.editor.locator('tr').first().locator('td, th').nth(0)).toContainText('A1');
  });

  // Bug: there is no Tab key binding for table cells, so Tab does not move to the
  // next cell (Word / Google Docs behaviour); it leaves the caret in place.
  test.fixme('D-INS-06 Tab moves to the next table cell', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('x');
    await doc.toolbar('Insert table').click();
    await page.getByRole('button', { name: /Insert 3×3 table/ }).click();
    await doc.pages.locator('td').first().dblclick();
    await doc.settle();
    await doc.typeText('A1');
    await doc.press('Tab');
    await page.keyboard.type('B1');
    await expect(doc.editor.locator('tr').first().locator('td, th').nth(1)).toContainText('B1');
  });

  test('D-PAG-01 the demo paginates into several pages with page-number footers @smoke', async () => {
    await expect(doc.pages).toHaveCount(2);
    await expect(doc.page.locator('.page-footer').nth(0)).toContainText('Page 1 of 2');
    await expect(doc.page.locator('.page-footer').nth(1)).toContainText('Page 2 of 2');
  });

  test('D-PAG-02 explicit page break starts a new page and footers update', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('first');
    await expect(doc.pages).toHaveCount(1);
    await doc.toolbar('More tools').click();
    await page.getByRole('menuitem', { name: 'Page break' }).click();
    await doc.typeText('second');
    await expect(doc.pages).toHaveCount(2);
    await expect(page.locator('.page-footer').nth(1)).toContainText('Page 2 of 2');
  });

  test('D-PAG-03 typing enough paragraphs flows onto a new page', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('line');
    for (let i = 0; i < 45; i++) await page.keyboard.press('Enter');
    await expect(doc.pages).toHaveCount(2);
  });

  test('D-HDR-01 the first-page header can be edited', async ({ page }) => {
    await page.getByText('Click to add header').first().click();
    await expect(page.getByRole('button', { name: 'Done' })).toBeVisible();
    // Focus moves to the header editor asynchronously; keys sent earlier hit the body editor.
    await expect(page.locator('.header-footer-pm-editor')).toBeFocused();
    await page.keyboard.type('My Header');
    await page.getByRole('button', { name: 'Done' }).click();
    // The demo uses a different first-page header, which is what this edits.
    await expect(page.locator('.page-header').first()).toContainText('My Header');
  });

  test('D-ZOOM-01 zoom menu scales the page', async ({ page }) => {
    const width = () =>
      page
        .locator('.page')
        .first()
        .evaluate((e) => e.getBoundingClientRect().width);
    expect(await width()).toBeCloseTo(816, 0);
    await doc.toolbar('100%').click();
    await page.getByText('50%', { exact: true }).click();
    await expect.poll(width).toBeCloseTo(408, 0);
    await doc.toolbar('50%').click();
    await page.getByText('200%', { exact: true }).click();
    await expect.poll(width).toBeCloseTo(1632, 0);
  });

  test('D-PGS-01 page setup switches to landscape', async ({ page }) => {
    await doc.toolbar('More tools').click();
    await page.getByRole('menuitem', { name: 'Page setup' }).click();
    await expect(page.getByRole('heading', { name: 'Page Setup' })).toBeVisible();
    await page.getByRole('button', { name: 'Landscape' }).click();
    await page.getByRole('button', { name: 'Apply' }).click();
    await expect
      .poll(() =>
        page
          .locator('.page')
          .first()
          .evaluate((e) => e.getBoundingClientRect().width)
      )
      .toBeCloseTo(1056, 0);
  });

  test('D-PGS-02 page setup can be cancelled', async ({ page }) => {
    await doc.toolbar('More tools').click();
    await page.getByRole('menuitem', { name: 'Page setup' }).click();
    await page.getByRole('button', { name: 'Cancel' }).click();
    await expect(page.getByRole('heading', { name: 'Page Setup' })).toBeHidden();
  });

  test('D-CMT-01 add a comment to selected text @smoke', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('hello');
    await doc.press(`${MOD}+a`);
    await doc.toolbar('More tools').click();
    await page.getByRole('menuitem', { name: 'Comment' }).click();
    await page.getByPlaceholder('Add a comment…').fill('Nice word');
    await page.getByRole('button', { name: 'Comment', exact: true }).click();
    await expect(page.getByText('Nice word')).toBeVisible();
    await expect(doc.editor.locator('[data-comment-thread]').first()).toBeAttached();
  });

  test('D-CMT-02 comment can be cancelled and the panel closes', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('hello');
    await doc.press(`${MOD}+a`);
    await doc.toolbar('More tools').click();
    await page.getByRole('menuitem', { name: 'Comment' }).click();
    await page.getByRole('button', { name: 'Cancel' }).click();
    await expect(page.getByPlaceholder('Add a comment…')).toBeHidden();
  });
});
