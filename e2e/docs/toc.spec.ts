import { test, expect, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { DocsPage } from '../support/docs';
import path from 'node:path';
import { captureDownload, collectPageErrors, openDocxConfirmingLoss } from '../support/file-io';

const NOTES_DOCX = path.resolve(__dirname, '..', '..', 'fixtures', 'docx', 'notes-toc-fields.docx');

test.describe('docs: table of contents', () => {
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

  /** Make the block the caret is in a heading of the given level, through the block type menu. */
  async function setHeading(page: Page, level: number) {
    await page.getByRole('button', { name: /^(Paragraph|Heading \d)$/ }).click();
    await page.getByText(`Heading ${level}`, { exact: true }).click();
    await doc.settle();
  }

  /** Three headings, each on a page of its own, and the caret left in the first one. */
  async function threeHeadings(page: Page) {
    await doc.clearAndFocus();
    await doc.typeText('Chapter one');
    await setHeading(page, 1);
    await page.keyboard.press('Enter');
    await doc.typeText('Some text.');
    await doc.toolbar('Page break').click();
    await doc.typeText('Chapter two');
    await setHeading(page, 1);
    await page.keyboard.press('Enter');
    await doc.toolbar('Page break').click();
    await doc.typeText('Section three');
    await setHeading(page, 2);
    await page.locator('.true-layout-viewport').evaluate((el) => (el.scrollTop = 0));
    await doc.clickText('Chapter one');
  }

  const menuItem = async (page: Page, name: RegExp) => {
    await doc.toolbar('Table of contents').click();
    await page.getByText(name).click();
  };
  /** The lines of the table of contents: a heading's text, and the page number after it. */
  const entry = (title: string) =>
    doc.pages.locator('.line').filter({ hasText: new RegExp(`^${title}\\d+$`) });

  test('D-TOC-01 Insert -> Table of contents lists the headings with the page each is on, after the block the caret is in @smoke', async ({
    page,
  }) => {
    await threeHeadings(page);
    await menuItem(page, /Insert table of contents/);

    // One entry per heading, with its page number; a level 2 heading is indented under them.
    await expect(entry('Chapter one')).toHaveText('Chapter one1');
    await expect(entry('Chapter two')).toHaveText('Chapter two2');
    await expect(entry('Section three')).toHaveText('Section three3');
    const indent = async (title: string) => (await entry(title).boundingBox())!.x;
    expect(await indent('Section three')).toBeGreaterThan((await indent('Chapter two')) + 10);
  });

  test('D-TOC-02 Update lists a heading added later, with the page it is on; entries are not changed until then', async ({
    page,
  }) => {
    await threeHeadings(page);
    await menuItem(page, /Insert table of contents/);
    await expect(entry('Chapter one')).toHaveText('Chapter one1');

    // A fourth heading on a new page at the end of the document.
    await doc.clickText('Chapter one');
    await doc.press('Control+End');
    await doc.toolbar('Page break').click();
    await doc.typeText('Chapter four');
    await setHeading(page, 1);
    await expect(entry('Chapter four')).toHaveCount(0);

    await menuItem(page, /Update table of contents/);
    await expect(entry('Chapter four')).toHaveText('Chapter four4');
    await expect(entry('Chapter one')).toHaveText('Chapter one1');
  });

  test('D-TOC-03 clicking an entry goes to its heading', async ({ page }) => {
    await threeHeadings(page);
    await menuItem(page, /Insert table of contents/);
    await expect(entry('Section three')).toHaveText('Section three3');
    // The heading is on the third page, out of view while the first one is shown.
    const heading = doc.pages.locator('.line').filter({ hasText: /^\u200BSection three$/ });
    await expect(heading).not.toBeInViewport();
    await entry('Section three').locator('a').click();
    await expect(heading).toBeInViewport();
  });

  test('D-TOC-04 the table of contents downloads as a TOC field Word can refresh, with PAGEREF fields and bookmarks on the headings', async ({
    page,
  }) => {
    await threeHeadings(page);
    await menuItem(page, /Insert table of contents/);
    await expect(entry('Chapter two')).toHaveText('Chapter two2');
    const file = await captureDownload(page, () =>
      page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
    );
    const xml =
      (await (await JSZip.loadAsync(file.bytes)).file('word/document.xml')?.async('string')) ?? '';
    expect(xml).toMatch(/<w:instrText[^>]*> TOC \\o &quot;1-3&quot; \\h \\z \\u <\/w:instrText>/);
    expect(xml.match(/<w:instrText[^>]*> PAGEREF _Toc\d{9} \\h <\/w:instrText>/g)).toHaveLength(3);
    expect(xml.match(/<w:bookmarkStart [^>]*w:name="_Toc\d{9}"/g)).toHaveLength(3);
    expect(xml.match(/<w:hyperlink [^>]*w:anchor="_Toc\d{9}"/g)).toHaveLength(3);
  });

  test('D-TOC-05 Remove deletes the table of contents', async ({ page }) => {
    await threeHeadings(page);
    await menuItem(page, /Insert table of contents/);
    await expect(entry('Chapter one')).toBeVisible();
    await menuItem(page, /Remove table of contents/);
    await expect(entry('Chapter one')).toHaveCount(0);
    await expect(
      doc.pages.locator('.line').filter({ hasText: /^\u200BChapter one$/ })
    ).toBeVisible();
  });

  async function openNotesDocx(page: Page) {
    await openDocxConfirmingLoss(page, NOTES_DOCX);
    await expect(doc.pages.first()).toBeVisible();
  }

  test('D-TOC-06 a table of contents in a .docx shows its entries with the numbers it was saved with, and opening does not change them @smoke', async ({
    page,
  }) => {
    await openNotesDocx(page);
    // Entries with a dotted leader; the file's cached numbers (the file was paginated to four pages, the editor's layout has fewer).
    await expect(entry('Overview')).toHaveText('Overview1');
    await expect(entry('Method')).toHaveText('Method2');
    await expect(entry('Recommendations')).toHaveText('Recommendations4');
    await expect(doc.pages.locator('.tab-leader').first()).toBeVisible();
  });

  test('D-TOC-07 Update rewrites the numbers of an imported table of contents from the editor’s pages, and the download has the field Word can refresh', async ({
    page,
  }) => {
    await openNotesDocx(page);
    await expect(entry('Recommendations')).toHaveText('Recommendations4');
    await menuItem(page, /Update table of contents/);
    await expect(entry('Recommendations')).not.toHaveText('Recommendations4');
    await expect(entry('Overview')).toHaveText(/^Overview[123]$/);

    const file = await captureDownload(page, () =>
      page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
    );
    const xml =
      (await (await JSZip.loadAsync(file.bytes)).file('word/document.xml')?.async('string')) ?? '';
    expect(xml).toMatch(/<w:instrText[^>]*> TOC \\o &quot;1-3&quot; \\h \\z \\u <\/w:instrText>/);
    expect(
      xml.match(/<w:instrText[^>]*> PAGEREF _Toc\d+ \\h <\/w:instrText>/g)!.length
    ).toBeGreaterThanOrEqual(7);
    expect(xml).toMatch(/<w:bookmarkStart [^>]*w:name="_Toc9000"/);
    expect(xml).not.toContain('updateFields');
  });
});
