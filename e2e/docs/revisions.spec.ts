import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { DocsPage, MOD } from '../support/docs';
import path from 'node:path';
import { captureDownload, collectPageErrors, openDocxConfirmingLoss } from '../support/file-io';

const REVIEW_DOCX = path.resolve(
  __dirname,
  '..',
  '..',
  'fixtures',
  'docx',
  'revisions-authors.docx'
);

test.describe('docs: tracked changes in a .docx', () => {
  let doc: DocsPage;
  let pageErrors: string[];
  test.beforeEach(async ({ page }) => {
    await page.route(/fonts\.b32n\.io/, (r) => r.abort());
    pageErrors = collectPageErrors(page);
    doc = new DocsPage(page);
    await doc.goto();
    await openDocxConfirmingLoss(page, REVIEW_DOCX);
    await expect(doc.pages.first()).toBeVisible();
  });
  test.afterEach(() => {
    expect(pageErrors, 'uncaught page errors').toEqual([]);
  });

  test('D-REV-01 a .docx with tracked changes shows insertions underlined and deletions struck through in each author’s colour, with a legend of the authors @smoke', async ({
    page,
  }) => {
    const legend = page.getByTestId('revision-legend');
    await expect(legend).toBeVisible();
    await expect(legend.getByRole('listitem')).toHaveCount(4);
    await expect(legend).toContainText('Alice Archer');

    // "2024" was replaced by "2025": both are in the page, the old one struck through and the new one underlined.
    const deleted = doc.pages.locator('[data-revision-deletion]', { hasText: '2024' }).first();
    const inserted = doc.pages.locator('[data-revision-insertion]', { hasText: '2025' }).first();
    await expect(deleted).toBeVisible();
    await expect(inserted).toBeVisible();
    await expect(deleted).toHaveCSS('text-decoration-line', 'line-through');
    await expect(inserted).toHaveCSS('text-decoration-line', 'underline');
    // Both are in Alice's colour, which is the colour her entry in the legend has.
    const swatch = legend.locator('[data-revision-author="Alice Archer"] span[aria-hidden]');
    const colour = await swatch.evaluate((el) => getComputedStyle(el).backgroundColor);
    await expect(deleted).toHaveCSS('color', colour);
    await expect(inserted).toHaveCSS('color', colour);
    // A different author is in a different colour.
    const other = doc.pages
      .locator('[data-revision-author]:not([data-revision-author="Alice Archer"])')
      .first();
    await expect(other).not.toHaveCSS('color', colour);
  });

  test('D-REV-02 the tracked changes are kept when the document is downloaded', async ({
    page,
  }) => {
    const file = await captureDownload(page, () =>
      page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
    );
    const xml =
      (await (await JSZip.loadAsync(file.bytes)).file('word/document.xml')?.async('string')) ?? '';
    expect(xml.match(/<w:ins /g)!.length).toBeGreaterThanOrEqual(10);
    expect(xml.match(/<w:del /g)!.length).toBeGreaterThanOrEqual(8);
    expect(xml).toMatch(
      /<w:del [^>]*w:author="Alice Archer"[^>]*><w:r>[\s\S]*?<w:delText[^>]*>2024<\/w:delText>/
    );
    expect(xml).toMatch(/<w:ins [^>]*w:author="Alice Archer"[^>]*w:date="2026-01-05T10:00:00Z"/);
  });

  test('D-REV-08 typing in the middle of a deleted run is live text, not more deleted text', async ({
    page,
  }) => {
    const deleted = doc.pages.locator('[data-revision-deletion]', { hasText: '2024' }).first();
    const box = (await deleted.boundingBox())!;
    await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
    await expect(doc.editor).toBeFocused();
    await doc.settle();
    await page.keyboard.type('Z');
    await expect(doc.pages.first()).toContainText('Z');
    // The typed letter is not inside any deleted run.
    await expect(doc.pages.locator('[data-revision-deletion]', { hasText: 'Z' })).toHaveCount(0);
    const file = await captureDownload(page, () =>
      page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
    );
    const xml =
      (await (await JSZip.loadAsync(file.bytes)).file('word/document.xml')?.async('string')) ?? '';
    expect(xml).not.toMatch(/<w:delText[^>]*>[^<]*Z[^<]*<\/w:delText>/);
  });

  test('D-REV-09 pasting a copied tracked range never repeats a change id', async ({
    page,
    context,
    browserName,
  }) => {
    test.skip(browserName !== 'chromium', 'clipboard permissions are Chromium-only');
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    const ids = async () => {
      const file = await captureDownload(page, () =>
        page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
      );
      const xml =
        (await (await JSZip.loadAsync(file.bytes)).file('word/document.xml')?.async('string')) ??
        '';
      const all = [...xml.matchAll(/<w:(?:ins|del) [^>]*w:id="(\d+)"/g)].map((m) => m[1]);
      return { all, distinct: new Set(all).size };
    };
    const before = await ids();
    // Select the deleted "2024" and the inserted "2025" after it, copy, and paste at the start of the document.
    const deleted = doc.pages.locator('[data-revision-deletion]', { hasText: '2024' }).first();
    const box = (await deleted.boundingBox())!;
    await page.mouse.click(box.x + 1, box.y + box.height / 2);
    await expect(doc.editor).toBeFocused();
    await doc.settle();
    for (let i = 0; i < 8; i++) await page.keyboard.press('Shift+ArrowRight');
    await doc.settle();
    await page.keyboard.press(`${MOD}+c`);
    await page.keyboard.press(`${MOD}+Home`);
    await doc.settle();
    await page.keyboard.press(`${MOD}+v`);
    await doc.settle();
    const after = await ids();
    // The pasted text has landed, and every tracked element it added (the clipboard may carry only its text) has an id of
    // its own: none repeats the id of the change it was copied from.
    await expect(doc.pages.first()).toContainText('20242025');
    expect(after.distinct - before.distinct).toBe(after.all.length - before.all.length);
  });
});
