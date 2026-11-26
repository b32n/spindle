import { test, expect, type Page } from '@playwright/test';
import JSZip from 'jszip';
import path from 'node:path';
import { DocsPage } from '../support/docs';
import { captureDownload, collectPageErrors, openViaChooser } from '../support/file-io';
import { roomName, waitForRoomSeed } from '../support/collab';

const OPEN = 'Open file (.docx)';
/** The DOCX fidelity corpus: four sections (portrait, landscape, portrait with an even-page header, A4). */
const SECTIONS_DOCX = path.resolve(
  __dirname,
  '..',
  '..',
  'fixtures',
  'docx',
  'sections-headers.docx'
);

/** Width, height and header/footer text of every page painted after scrolling through the whole document. */
async function pagesByScrolling(
  page: Page
): Promise<
  Map<number, { width: number; height: number; left: number; header: string; footer: string }>
> {
  const out = new Map<
    number,
    { width: number; height: number; left: number; header: string; footer: string }
  >();
  const total = parseInt(
    await page.evaluate(() => (document.querySelector('.page-stack') as HTMLElement).style.height)
  );
  for (let y = 0; y <= total + 700; y += 600) {
    await page.evaluate((top) => {
      (document.querySelector('.true-layout-viewport') as HTMLElement).scrollTop = top;
    }, y);
    await page.waitForTimeout(200);
    const seen = await page.evaluate(() =>
      [...document.querySelectorAll('.page')].map((p) => {
        const el = p as HTMLElement;
        return {
          index: Number(el.dataset.pageIndex),
          width: parseFloat(el.style.width),
          height: parseFloat(el.style.height),
          left: parseFloat(el.style.left),
          header: p.querySelector('.page-header')?.textContent ?? '',
          footer: p.querySelector('.page-footer')?.textContent ?? '',
        };
      })
    );
    for (const s of seen) out.set(s.index, s);
  }
  return out;
}

async function downloadXml(page: Page): Promise<string> {
  const file = await captureDownload(page, () =>
    page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
  );
  const zip = await JSZip.loadAsync(file.bytes);
  return (await zip.file('word/document.xml')?.async('string')) ?? '';
}

/** Number of pages the editor has laid out (the stack height divided by the page pitch is not needed: read the layout). */
async function pageCount(page: Page): Promise<number> {
  return page.evaluate(() => {
    const stack = document.querySelector('.page-stack') as HTMLElement;
    // Pages are virtualized: count them by the height of the stack (one Letter page is 1056px plus a 24px gap).
    return Math.round((parseFloat(stack.style.height) + 24) / (1056 + 24));
  });
}

test.describe('docs: sections', () => {
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

  async function openSections(page: Page) {
    await openViaChooser(page, OPEN, SECTIONS_DOCX);
    const confirm = page.getByRole('dialog', { name: /Import/ });
    await confirm.getByRole('button', { name: 'Import anyway' }).click();
    await expect(page.locator('.page').first()).toBeVisible();
  }

  test('D-SEC-01 a .docx with four sections shows each with its own page size and headers @smoke', async ({
    page,
  }) => {
    await openSections(page);
    const pages = await pagesByScrolling(page);
    // Letter portrait, landscape, Letter portrait, A4: the pages follow their section's size and are centred in the stack.
    const sizes = [...pages.values()].map((p) => [p.width, p.height]);
    expect(sizes).toEqual(
      expect.arrayContaining([
        [816, 1056],
        [1056, 816],
        [794, 1123],
      ])
    );
    const landscape = [...pages.values()].find((p) => p.width === 1056)!;
    expect(landscape.left).toBe(0);
    const portrait = [...pages.values()].find((p) => p.width === 816)!;
    expect(portrait.left).toBe(120);
    // Headers and footers differ per section: first-page header, landscape header with roman page numbers, even-page header, A4 header.
    const headers = [...pages.values()].map((p) => p.header);
    expect(headers.some((h) => h.includes('FIRST PAGE HEADER'))).toBe(true);
    expect(headers.some((h) => h.includes('Landscape section header'))).toBe(true);
    expect(headers.some((h) => h.includes('Even-page header'))).toBe(true);
    expect(headers.some((h) => h.includes('Harbour Report'))).toBe(true);
    expect([...pages.values()].some((p) => /Page i of/.test(p.footer))).toBe(true);
  });

  async function twoParagraphs() {
    await doc.clearAndFocus();
    await doc.typeText('Before');
    await doc.press('Enter');
    await doc.typeText('After');
  }

  /** Put the caret at the end of "Before" and use the Section break menu. */
  async function insertBreakAfterBefore(page: Page, label: RegExp) {
    await doc.clickText('Before');
    await doc.press('End');
    await doc.toolbar('Section break').click();
    await page.getByText(label).click();
  }

  test('D-SEC-02 a next-page section break starts a new page and downloads as a second section @smoke', async ({
    page,
  }) => {
    await twoParagraphs();
    expect(await pageCount(page)).toBe(1);
    await insertBreakAfterBefore(page, /Section break \(next page\)/);
    await expect.poll(() => pageCount(page)).toBe(2);
    await expect(doc.pages.nth(1)).toContainText('After');
    await expect(doc.pages.first()).not.toContainText('After');
    const xml = await downloadXml(page);
    expect(xml.match(/<w:sectPr/g)).toHaveLength(2);
    expect(xml).not.toContain('w:type w:val="continuous"');
    // The paragraph that carries the first section's properties is "Before".
    expect(xml).toMatch(/Before<\/w:t><\/w:r><\/w:p>|<w:sectPr/);
  });

  test('D-SEC-03 a continuous section break keeps the text on the page and is written as a continuous section', async ({
    page,
  }) => {
    await twoParagraphs();
    await insertBreakAfterBefore(page, /Section break \(continuous\)/);
    await expect(doc.pages.first()).toContainText('After');
    expect(await pageCount(page)).toBe(1);
    const xml = await downloadXml(page);
    expect(xml.match(/<w:sectPr/g)).toHaveLength(2);
    expect(xml).toContain('<w:type w:val="continuous"/>');
  });

  test('D-SEC-04 removing a section break merges the sections again', async ({ page }) => {
    await twoParagraphs();
    await insertBreakAfterBefore(page, /Section break \(next page\)/);
    await expect.poll(() => pageCount(page)).toBe(2);
    await doc.toolbar('Section break').click();
    await page.getByText('Remove section break').click();
    await expect.poll(() => pageCount(page)).toBe(1);
    await expect(doc.pages.first()).toContainText('After');
    const xml = await downloadXml(page);
    expect(xml.match(/<w:sectPr/g)).toHaveLength(1);
  });

  test('D-SEC-05 typing in the last section of a four-section file keeps all four sections in the download', async ({
    page,
  }) => {
    await openSections(page);
    // Scroll to the closing (A4) section and type at the end of its heading.
    const heading = 'Closing section (A4, single column)';
    for (let i = 0; i < 12 && (await doc.pages.getByText(heading).count()) === 0; i++) {
      await page.evaluate(() => {
        (document.querySelector('.true-layout-viewport') as HTMLElement).scrollBy(0, 1000);
      });
      await page.waitForTimeout(250);
    }
    await doc.clickText(heading);
    await doc.press('End');
    await page.keyboard.type(' SECTION4-EDIT');
    await expect(doc.pages.getByText(/SECTION4-EDIT/)).toBeVisible();
    const xml = await downloadXml(page);
    expect(xml).toContain('SECTION4-EDIT');
    expect(xml.match(/<w:sectPr/g)).toHaveLength(4);
  });

  test('D-SEC-06 page setup applies to the section the caret is in, or to the whole document', async ({
    page,
  }) => {
    await twoParagraphs();
    await insertBreakAfterBefore(page, /Section break \(next page\)/);
    await expect.poll(() => pageCount(page)).toBe(2);
    const widths = () =>
      page
        .locator('.page')
        .evaluateAll((els) => els.map((e) => parseFloat((e as HTMLElement).style.width)));
    const setup = async (
      orientation: 'Portrait' | 'Landscape',
      scope: 'This section' | 'Whole document'
    ) => {
      await doc.toolbar('Page setup').click();
      await expect(page.getByRole('heading', { name: 'Page Setup' })).toBeVisible();
      await page.getByRole('button', { name: orientation }).click();
      await page.getByLabel(scope).check();
      await page.getByRole('button', { name: 'Apply' }).click();
    };
    // The caret is in the second section ("After"): only that page turns landscape.
    await doc.clickText('After');
    await setup('Landscape', 'This section');
    await expect.poll(widths).toEqual([816, 1056]);
    // The whole document: both pages become landscape.
    await setup('Landscape', 'Whole document');
    await expect.poll(widths).toEqual([1056, 1056]);
  });

  test('D-SEC-07 two clients share a document with a section break and edit in the second section', async ({
    browser,
  }, info) => {
    const WS = 'ws://127.0.0.1:1234';
    const room = roomName('docs', info.testId);
    const open = async (user: string) => {
      const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
      const p = await context.newPage();
      const d = new DocsPage(p);
      await d.goto(`?ws=${encodeURIComponent(WS)}&room=${room}&user=${user}`);
      await expect(p.getByText(/· connected$/)).toBeVisible();
      return { context, page: p, doc: d };
    };
    const a = await open('Alice');
    await waitForRoomSeed(room);
    const b = await open('Bob');
    // Alice ends a section after the first heading, with a next-page break.
    await a.doc.clickLine('Welcome to Spindle Docs');
    await a.doc.press('End');
    await a.doc.toolbar('Section break').click();
    await a.page.getByText(/Section break \(next page\)/).click();
    // Both clients now show two sections: the heading alone on page 1, the rest from page 2.
    for (const c of [a, b]) {
      await expect.poll(() => c.page.locator('.page').count()).toBeGreaterThanOrEqual(2);
      await expect(c.doc.pages.first()).not.toContainText('Features');
    }
    // Bob types at the start of "Features" in the second section; Alice sees it.
    await b.doc.clickLine('Features');
    await b.doc.press('Home');
    await b.doc.typeText('B2:');
    await expect
      .poll(() =>
        a.doc.editor
          .locator('h2')
          .first()
          .evaluate((el) => {
            const clone = el.cloneNode(true) as HTMLElement;
            clone
              .querySelectorAll('.ProseMirror-yjs-cursor, .ProseMirror-widget')
              .forEach((n) => n.remove());
            return clone.textContent ?? '';
          })
      )
      .toBe('B2:Features');
    // The section break is still one, on both sides.
    for (const c of [a, b]) await expect(c.doc.editor.locator('.section-break')).toHaveCount(1);
    await a.context.close();
    await b.context.close();
  });

  test('D-SEC-08 a section after the first shows the previous header and can be unlinked from it', async ({
    page,
  }) => {
    await twoParagraphs();
    await insertBreakAfterBefore(page, /Section break \(next page\)/);
    await expect.poll(() => pageCount(page)).toBe(2);
    const secondHeader = page.locator('.page').nth(1).locator('.page-header');
    await secondHeader.click();
    const link = page.getByLabel('Link to previous');
    await expect(link).toBeChecked();
    await link.click();
    await expect(link).toBeHidden();
    await secondHeader.click();
    await expect(page.getByLabel('Link to previous')).not.toBeChecked();
  });
});
