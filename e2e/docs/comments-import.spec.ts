import { test, expect, type Page } from '@playwright/test';
import path from 'node:path';
import { DocsPage } from '../support/docs';
import { collectPageErrors, openViaChooser } from '../support/file-io';

const OPEN = 'Open file (.docx)';
/** The DOCX fidelity corpus (a LibreOffice review file: two comment threads and tracked changes). */
const REVIEW_DOCX = path.resolve(
  __dirname,
  '..',
  '..',
  'fixtures',
  'docx',
  'libreoffice-review.docx'
);

test.describe('docs: comments imported from a .docx', () => {
  let doc: DocsPage;
  let pageErrors: string[];
  test.beforeEach(async ({ page }) => {
    await page.route(/fonts\.b32n\.io/, (r) => r.abort());
    pageErrors = collectPageErrors(page);
    doc = new DocsPage(page);
    await doc.goto();
  });
  test.afterEach(() => {
    expect(pageErrors, 'uncaught page errors').toEqual([]);
  });

  async function openReview(page: Page) {
    await openViaChooser(page, OPEN, REVIEW_DOCX);
    // The file has tracked changes, so the import is confirmed first.
    await page
      .getByRole('dialog', { name: /Import/ })
      .getByRole('button', { name: 'Import anyway' })
      .click();
  }

  test('D-CMT-03 a .docx with comment threads shows them anchored in the text and listed in the sidebar @smoke', async ({
    page,
  }) => {
    await openReview(page);

    // Both anchored ranges carry a comment highlight on the rendered pages (the ProseMirror
    // instance is hidden off-screen, so it cannot be clicked).
    const anchors = doc.pages.locator('[data-comment-thread]');
    await expect(anchors).not.toHaveCount(0);
    await expect(anchors.filter({ hasText: 'Fees are set annually' })).toHaveCount(1);

    // Clicking a highlight opens the sidebar on its thread: the root comment and both replies, in order.
    await anchors.filter({ hasText: 'hours' }).first().click();
    await expect(page.getByText('Comments', { exact: true })).toBeVisible();
    const sidebarText = page.getByText(/Is 48 hours long enough for visiting vessels\?/);
    await expect(sidebarText).toBeVisible();
    await expect(page.getByText('Council guidance says 72 hours.')).toBeVisible();
    await expect(page.getByText('Agreed, changing it.')).toBeVisible();
    await expect(page.getByText('Alice Archer').first()).toBeVisible();

    // The second thread is listed too.
    await expect(page.getByText('Cite the 2019 statute here.')).toBeVisible();
  });
});
