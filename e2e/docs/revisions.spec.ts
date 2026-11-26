import { test, expect, type Page } from '@playwright/test';
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

  test('D-REV-03 moves, formatting changes, paragraph changes and inserted or deleted rows are marked, and kept in the download', async ({
    page,
  }) => {
    // Both ends of a move have a double line and share the name of the move.
    const moves = doc.pages.locator('[data-revision-move]');
    await expect(moves.first()).toBeVisible();
    await expect(moves.first()).toHaveCSS('text-decoration-style', 'double');
    // Each name has one end that is a deletion (moved from) and one that is an insertion (moved to).
    const ends = await moves.evaluateAll((els) =>
      els.map((el) => ({
        name: (el as HTMLElement).dataset.revisionMove ?? '',
        deletion: el.hasAttribute('data-revision-deletion'),
        insertion: el.hasAttribute('data-revision-insertion'),
      }))
    );
    const names = [...new Set(ends.map((end) => end.name))];
    expect(names.length).toBeGreaterThan(0);
    for (const name of names) {
      const named = ends.filter((end) => end.name === name);
      expect(
        named.some((end) => end.deletion),
        `${name}: moved from`
      ).toBe(true);
      expect(
        named.some((end) => end.insertion),
        `${name}: moved to`
      ).toBe(true);
    }
    // A change of formatting alone is dotted and keeps the text's colour.
    const formatted = doc.pages
      .locator(
        '[data-revision-format]:not([data-revision-insertion]):not([data-revision-deletion])'
      )
      .first();
    await expect(formatted).toHaveCSS('text-decoration-style', 'dotted');
    // A paragraph with a change has a bar in the margin; a row inserted or deleted is tinted.
    await expect(page.locator('.change-bar').first()).toBeVisible();
    await expect(page.locator('.change-bar').first()).not.toHaveCSS(
      'background-color',
      'rgba(0, 0, 0, 0)'
    );
    await expect(doc.pages.locator('p[data-revision-block]').first()).toBeVisible();
    await expect(doc.pages.locator('td[data-revision-row="insertion"]').first()).toBeVisible();
    // The tint is a gradient over the cell's own background.
    await expect(doc.pages.locator('td[data-revision-row="insertion"]').first()).not.toHaveCSS(
      'background-image',
      'none'
    );
    await expect(doc.pages.locator('td[data-revision-row="deletion"]').first()).toBeVisible();

    const file = await captureDownload(page, () =>
      page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
    );
    const xml =
      (await (await JSZip.loadAsync(file.bytes)).file('word/document.xml')?.async('string')) ?? '';
    for (const element of [
      '<w:moveFrom ',
      '<w:moveTo ',
      '<w:rPrChange ',
      '<w:pPrChange ',
      '<w:trPr><w:ins ',
      '<w:trPr><w:del ',
    ]) {
      expect(xml, element).toContain(element);
    }
    expect(xml).toMatch(/<w:moveFromRangeStart [^>]*w:name="move-1"/);
    // The records carry their author and date, and the formatting before the change.
    expect(xml).toMatch(/<w:rPrChange [^>]*w:author="[^"]+"[^>]*w:date="[^"]+"/);
    expect(xml).toMatch(/<w:rPrChange [^>]*>\s*<w:rPr[ >]/);
    expect(xml).toMatch(/<w:pPrChange [^>]*>\s*<w:pPr[ >]/);
    expect(xml).toMatch(/<w:pPrChange [^>]*w:author="[^"]+"/);
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
    // Select the deleted "2024" and the inserted "2025" after it, copy, and paste at the start of that line.
    const deleted = doc.pages.locator('[data-revision-deletion]', { hasText: '2024' }).first();
    const box = (await deleted.boundingBox())!;
    await page.mouse.click(box.x + 1, box.y + box.height / 2);
    await expect(doc.editor).toBeFocused();
    await doc.settle();
    for (let i = 0; i < 8; i++) await page.keyboard.press('Shift+ArrowRight');
    await doc.settle();
    await page.keyboard.press(`${MOD}+c`);
    // The caret to the start of the line (Left collapses the selection to its start), so the paste adds to the document.
    await page.keyboard.press('ArrowLeft');
    await page.keyboard.press('Home');
    await doc.settle();
    await page.keyboard.press(`${MOD}+v`);
    await doc.settle();
    const after = await ids();
    // The pasted text has landed with its tracked changes (the clipboard carried the markup, not only the text), and every
    // tracked element it added has an id of its own: none repeats the id of the change it was copied from.
    await expect(doc.pages.first()).toContainText('20242025');
    expect(after.all.length).toBeGreaterThan(before.all.length);
    expect(after.distinct - before.distinct).toBe(after.all.length - before.all.length);
  });

  const panel = (page: Page) => page.getByTestId('review-panel');
  const openPanel = async (page: Page) => {
    await doc.toolbar('Review changes').click();
    await expect(panel(page)).toBeVisible();
  };
  const markup = (doc: DocsPage) =>
    doc.pages.locator('[data-revision-insertion], [data-revision-deletion]');

  test('D-REV-04 the review panel lists the changes; accept an insertion, reject a deletion, accept all, and check the text @smoke', async ({
    page,
  }) => {
    await openPanel(page);
    const changes = panel(page).getByTestId('review-change');
    expect(await changes.count()).toBeGreaterThan(15);

    // "2024" was replaced by "2025" (Alice): accepting the insertion keeps "2025" as plain text, rejecting the deletion keeps "2024".
    const inserted = changes.filter({ hasText: /Inserted/ }).filter({ hasText: '“2025”' });
    await inserted.getByRole('button', { name: 'Accept change' }).click();
    await expect(doc.pages.locator('[data-revision-insertion]', { hasText: '2025' })).toHaveCount(
      0
    );
    await expect(doc.pages.first()).toContainText('2025');
    const deleted = changes.filter({ hasText: /Deleted/ }).filter({ hasText: '“2024”' });
    await deleted.getByRole('button', { name: 'Reject change' }).click();
    await expect(doc.pages.locator('[data-revision-deletion]', { hasText: '2024' })).toHaveCount(0);
    await expect(doc.pages.first()).toContainText('20242025');

    // Rejecting an insertion takes its text away ("eleven" was inserted by Bram).
    await expect(doc.pages.first()).toContainText('eleven');
    await changes
      .filter({ hasText: '“eleven”' })
      .getByRole('button', { name: 'Reject change' })
      .click();
    await expect(doc.pages.first()).not.toContainText('eleven');

    // Accept all: no markup is left, the text of every deletion is gone, the panel and the legend are empty.
    await panel(page).getByRole('button', { name: 'Accept all', exact: true }).click();
    await expect(markup(doc)).toHaveCount(0);
    await expect(panel(page)).toContainText('No tracked changes.');
    await expect(page.getByTestId('revision-legend')).toHaveCount(0);
    await expect(doc.pages.first()).not.toContainText('This sentence was judged redundant');
    await expect(doc.pages.first()).toContainText(
      'A replacement sentence was supplied by the fourth reviewer'
    );

    // The download has no tracked changes left.
    const file = await captureDownload(page, () =>
      page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
    );
    const xml =
      (await (await JSZip.loadAsync(file.bytes)).file('word/document.xml')?.async('string')) ?? '';
    expect(xml).not.toMatch(/<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange)\b/);
  });

  test('D-REV-05 the Final and Original views show the text as if every change were accepted or rejected, and filtering by author limits Accept all', async ({
    page,
  }) => {
    await openPanel(page);
    const view = panel(page).getByRole('group', { name: 'Show changes as' });
    // Bram replaced "eight" by "eleven".
    await view.getByRole('button', { name: 'Final' }).click();
    await expect(doc.pages.first()).toContainText('eleven');
    await expect(doc.pages.first()).not.toContainText('eight');
    await expect(markup(doc)).toHaveCount(0);
    await view.getByRole('button', { name: 'Original' }).click();
    await expect(doc.pages.first()).toContainText('eight');
    await expect(doc.pages.first()).not.toContainText('eleven');
    await expect(markup(doc)).toHaveCount(0);
    await view.getByRole('button', { name: 'All markup' }).click();
    await expect(doc.pages.first()).toContainText('eight');
    await expect(doc.pages.first()).toContainText('eleven');
    await expect(markup(doc).first()).toBeVisible();

    // Reject all by Bram: his insertion is gone, his deletion is plain text again, the other authors' changes stay.
    await panel(page).getByLabel('Author').selectOption('Bram Keller');
    await expect(panel(page).getByTestId('review-change')).not.toHaveCount(0);
    expect(
      await panel(page)
        .getByTestId('review-change')
        .evaluateAll((els) => els.map((e) => (e as HTMLElement).dataset.author))
    ).toEqual(expect.arrayContaining(['Bram Keller']));
    await panel(page).getByRole('button', { name: 'Reject all by Bram Keller' }).click();
    await expect(doc.pages.first()).toContainText('eight');
    await expect(doc.pages.first()).not.toContainText('eleven');
    await expect(page.getByTestId('revision-legend')).not.toContainText('Bram Keller');
    await expect(page.getByTestId('revision-legend')).toContainText('Alice Archer');
  });

  test('D-REV-06 next and previous go through the changes, scrolling to each', async ({ page }) => {
    await openPanel(page);
    const items = panel(page).getByTestId('review-change');
    await panel(page).getByRole('button', { name: 'Next change' }).click();
    await expect(items.first()).toHaveAttribute('aria-current', 'true');
    await panel(page).getByRole('button', { name: 'Next change' }).click();
    await expect(items.nth(1)).toHaveAttribute('aria-current', 'true');
    await panel(page).getByRole('button', { name: 'Previous change' }).click();
    await expect(items.first()).toHaveAttribute('aria-current', 'true');
    // The last change is on the second page; going there scrolls it into view.
    await panel(page).getByRole('button', { name: 'Previous change' }).click();
    await expect(items.last()).toHaveAttribute('aria-current', 'true');
    await expect(doc.pages.locator('td[data-revision-row="deletion"]').first()).toBeInViewport();
  });
  test('D-REV-07 the Final and Original views are read-only, with a banner that switches back to All markup', async ({
    page,
  }) => {
    await openPanel(page);
    const view = panel(page).getByRole('group', { name: 'Show changes as' });
    const banner = page.getByTestId('revision-view-banner');
    // A read-only editor takes no focus (so `clickText` cannot be used): click the text and send the keys anyway.
    const tryTyping = async () => {
      await doc.pages.getByText('Total tonnage handled in').first().click();
      await page.keyboard.press('Home');
      await page.keyboard.type('ZZ');
    };
    await expect(banner).toHaveCount(0);

    // Final: banner, and typing changes nothing.
    await view.getByRole('button', { name: 'Final' }).click();
    await expect(banner).toBeVisible();
    await expect(banner).toContainText('Showing final text: switch to All markup to edit');
    await tryTyping();
    await expect(doc.pages.first()).not.toContainText('ZZ');
    await expect(markup(doc)).toHaveCount(0);

    // Original: the same.
    await view.getByRole('button', { name: 'Original' }).click();
    await expect(banner).toContainText('Showing original text: switch to All markup to edit');
    await tryTyping();
    await expect(doc.pages.first()).not.toContainText('ZZ');

    // Accepting a change from the panel is an explicit command and still works in these views.
    const items = panel(page).getByTestId('review-change');
    const before = await items.count();
    await panel(page).getByRole('button', { name: 'Accept change' }).first().click();
    await expect(items).toHaveCount(before - 1);

    // One click on the banner goes back to All markup, where the document can be edited again.
    await banner.getByRole('button', { name: 'Switch to All markup' }).click();
    await expect(banner).toHaveCount(0);
    await expect(view.getByRole('button', { name: 'All markup' })).toHaveAttribute(
      'aria-pressed',
      'true'
    );
    await expect(markup(doc).first()).toBeVisible();
    await doc.clickText('Total tonnage handled in');
    await doc.press('Home');
    await doc.typeText('ZZ');
    await expect(doc.pages.first()).toContainText('ZZTotal tonnage');
  });

  test('D-REV-10 edits and accepted changes made before a read-only view are still there after it', async ({
    page,
  }) => {
    // An edit in All markup, and a change accepted from the panel.
    await doc.clickText('Total tonnage handled in');
    await doc.press('Home');
    await doc.typeText('QQ');
    await openPanel(page);
    const items = panel(page).getByTestId('review-change');
    const before = await items.count();
    await panel(page).getByRole('button', { name: 'Accept change' }).first().click();
    await expect(items).toHaveCount(before - 1);

    // To the Final view and back: the editor is not rebuilt from the document it was mounted with.
    const view = panel(page).getByRole('group', { name: 'Show changes as' });
    await view.getByRole('button', { name: 'Final' }).click();
    await expect(page.getByTestId('revision-view-banner')).toBeVisible();
    await expect(doc.pages.first()).toContainText('QQTotal tonnage');
    await page
      .getByTestId('revision-view-banner')
      .getByRole('button', { name: 'Switch to All markup' })
      .click();
    await expect(doc.pages.first()).toContainText('QQTotal tonnage');
    await expect(items).toHaveCount(before - 1);

    // The next keystroke does not bring the old document back.
    await doc.clickText('QQTotal tonnage handled in');
    await doc.press('Home');
    await doc.typeText('W');
    await expect(doc.pages.first()).toContainText('WQQTotal tonnage');
    await expect(items).toHaveCount(before - 1);
  });
});
