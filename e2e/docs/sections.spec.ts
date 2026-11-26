import { test, expect, type Page } from '@playwright/test';
import path from 'node:path';
import { DocsPage } from '../support/docs';
import { collectPageErrors, openViaChooser } from '../support/file-io';

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
});
