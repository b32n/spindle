import { test, expect, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { DocsPage } from '../support/docs';
import { captureDownload, collectPageErrors } from '../support/file-io';

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

  test('D-TOC-04 the table of contents downloads as a Word table of contents block, with bookmarks on the headings', async ({
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
    expect(xml).toContain('<w:docPartGallery w:val="Table of Contents"/>');
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
});
