import { test, expect } from '@playwright/test';
import path from 'node:path';
import { DocsPage } from '../support/docs';
import { collectPageErrors, openViaChooser } from '../support/file-io';

const OPEN = 'Open file (.docx)';
/** The DOCX fidelity corpus: a document whose table of contents links to bookmarked headings further down. */
const KITCHEN_DOCX = path.resolve(__dirname, '..', '..', 'fixtures', 'docx', 'kitchen-sink.docx');

test.describe('docs: internal links from a .docx', () => {
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

  test('D-LNK-01 clicking a table-of-contents entry scrolls to the bookmarked heading', async ({
    page,
  }) => {
    await openViaChooser(page, OPEN, KITCHEN_DOCX);
    // The file has tracked changes, notes and more that cannot be imported, so the import is confirmed first.
    await page
      .getByRole('dialog', { name: /Import/ })
      .getByRole('button', { name: 'Import anyway' })
      .click();

    // An entry of the table of contents is an internal link (the ProseMirror instance is hidden
    // off-screen, so the rendered pages are what is clicked).
    const entry = doc.pages.locator('a', { hasText: 'Charts' });
    await expect(entry.first()).toBeVisible();
    // The heading it points at is further down, off screen until the link is followed.
    const heading = doc.pages.getByText('Charts', { exact: true }).and(page.locator(':not(a)'));
    await expect(heading.first()).not.toBeInViewport();
    await entry.first().click();

    // It is on screen now (and the address did not change).
    await expect(heading.first()).toBeInViewport();
    expect(new URL(page.url()).hash).toBe('');
  });
});
