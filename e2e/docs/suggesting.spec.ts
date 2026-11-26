import { test, expect, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { DocsPage } from '../support/docs';
import { asUpload, captureDownload, collectPageErrors, openViaChooser } from '../support/file-io';

const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

test.describe('docs: Suggesting mode', () => {
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

  const modes = (page: Page) => page.getByRole('group', { name: 'Editing mode' });
  const setMode = (page: Page, name: 'Editing' | 'Suggesting' | 'Viewing') =>
    modes(page).getByRole('button', { name }).click();
  const insertions = () => doc.pages.locator('[data-revision-insertion]');
  const deletions = () => doc.pages.locator('[data-revision-deletion]');
  const formats = () => doc.pages.locator('[data-revision-format]');

  /** "Hello world" typed in Editing mode, then Suggesting mode on with the caret at the start. */
  async function startSuggesting(page: Page) {
    await doc.clearAndFocus();
    await doc.typeText('Hello world');
    await setMode(page, 'Suggesting');
    await doc.clickText('Hello world');
    await doc.press('End');
  }

  test('D-SUG-01 in Suggesting mode, typing, deleting and formatting become tracked changes that survive Download and Open, and can be accepted @smoke', async ({
    page,
  }) => {
    await startSuggesting(page);
    // Typing: " big" after "Hello" is an insertion by the current user.
    for (let i = 0; i < 6; i++) await doc.press('ArrowLeft');
    await doc.typeText(' big');
    await expect(insertions().filter({ hasText: 'big' })).toHaveCount(1);
    await expect(insertions().first()).toHaveAttribute('data-revision-author', 'Alice');
    await expect(insertions().first()).toHaveCSS('text-decoration-line', 'underline');

    // Deleting: backspacing over "world" leaves it, struck through.
    await doc.press('End');
    for (let i = 0; i < 5; i++) await page.keyboard.press('Backspace');
    await doc.settle();
    await expect(deletions().filter({ hasText: 'world' })).toHaveCount(1);
    await expect(deletions().first()).toHaveCSS('text-decoration-line', 'line-through');
    await expect(doc.pages.first()).toContainText('Hello big world');

    // Typing and deleting one's own insertion leaves no trace.
    await doc.typeText('ZZ');
    await page.keyboard.press('Backspace');
    await page.keyboard.press('Backspace');
    await doc.settle();
    await expect(insertions().filter({ hasText: 'ZZ' })).toHaveCount(0);
    await expect(doc.pages.first()).not.toContainText('ZZ');

    // Formatting: bold on "Hello" records a formatting change.
    await doc.press('Home');
    for (let i = 0; i < 5; i++) await page.keyboard.press('Shift+ArrowRight');
    await doc.toolbar('Bold').click();
    await expect(formats().filter({ hasText: 'Hello' })).toHaveCount(1);
    // The selection is still "Hello": a second format reaches the same text.
    await doc.toolbar('Italic').click();
    await expect(formats().filter({ hasText: 'Hello' })).toHaveCount(1);
    await expect(formats().filter({ hasText: 'Hello' })).toHaveCSS('font-style', 'italic');
    await expect(formats().filter({ hasText: 'Hello' })).toHaveCSS('font-weight', '700');
    await expect(page.getByTestId('revision-legend')).toContainText('Alice');

    // Download keeps them as Word's tracked changes by the user.
    const file = await captureDownload(page, () =>
      page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
    );
    const xml =
      (await (await JSZip.loadAsync(file.bytes)).file('word/document.xml')?.async('string')) ?? '';
    expect(xml).toMatch(/<w:ins [^>]*w:author="Alice"/);
    expect(xml).toMatch(
      /<w:del [^>]*w:author="Alice"[^>]*><w:r>[\s\S]*?<w:delText[^>]*>world<\/w:delText>/
    );
    expect(xml).toContain('<w:rPrChange ');

    // Opening that file brings the changes back; accepting all leaves what was suggested.
    await openViaChooser(page, 'Open file (.docx)', asUpload(file, DOCX));
    await expect(deletions().filter({ hasText: 'world' })).toHaveCount(1);
    await expect(insertions().filter({ hasText: 'big' })).toHaveCount(1);
    await doc.toolbar('Review changes').click();
    await page
      .getByTestId('review-panel')
      .getByRole('button', { name: 'Accept all', exact: true })
      .click();
    await expect(doc.pages.first()).toContainText('Hello big');
    await expect(doc.pages.first()).not.toContainText('world');
    await expect(doc.pages.locator('[data-revision-author]')).toHaveCount(0);
  });

  test('D-SUG-02 Enter and Backspace in Suggesting mode record paragraph marks, and rejecting puts the paragraph back together', async ({
    page,
  }) => {
    await startSuggesting(page);
    for (let i = 0; i < 6; i++) await doc.press('ArrowLeft');
    await page.keyboard.press('Enter');
    await doc.settle();
    // The paragraph that ends at the new break has a bar in the margin.
    await expect(page.locator('.change-bar')).toHaveCount(1);
    await expect(doc.pages.first().locator('p[data-revision-block="insertion"]')).toHaveCount(1);

    await doc.toolbar('Review changes').click();
    const panel = page.getByTestId('review-panel');
    await expect(panel.getByTestId('review-change')).toHaveCount(1);
    await expect(panel.getByTestId('review-change')).toHaveAttribute(
      'data-kind',
      'paragraphInserted'
    );
    await panel.getByRole('button', { name: 'Reject change' }).click();
    await expect(page.locator('.change-bar')).toHaveCount(0);
    await expect(doc.pages.first().locator('p')).toHaveCount(1);
    await expect(doc.pages.first()).toContainText('Hello world');
  });

  test('D-SUG-03 undo takes a suggestion back in one step', async ({ page }) => {
    await startSuggesting(page);
    // Typing within half a second of earlier edits joins their undo step (the editor's history groups them).
    await page.waitForTimeout(700);
    await doc.typeText('XYZ');
    await expect(insertions().filter({ hasText: 'XYZ' })).toHaveCount(1);
    await doc.toolbar('Undo').click();
    await expect(insertions()).toHaveCount(0);
    await expect(doc.pages.first()).not.toContainText('XYZ');
    await expect(doc.pages.first()).toContainText('Hello world');
  });

  test('D-SUG-04 Viewing mode cannot change the document, Editing changes it directly again', async ({
    page,
  }) => {
    await doc.clearAndFocus();
    await doc.typeText('Hello world');
    await setMode(page, 'Viewing');
    await expect(doc.toolbar('Bold')).toHaveCount(0);
    await doc.pages.first().getByText('Hello world').click();
    await page.keyboard.type('NOPE');
    await expect(doc.pages.first()).not.toContainText('NOPE');
    await setMode(page, 'Editing');
    await expect(doc.toolbar('Bold')).toHaveCount(1);
    await doc.clickText('Hello world');
    await doc.press('End');
    await doc.typeText('!');
    await expect(doc.pages.first()).toContainText('Hello world!');
    await expect(doc.pages.locator('[data-revision-author]')).toHaveCount(0);
  });

  test('D-SUG-05 what cannot be suggested is done directly, with a notice', async ({ page }) => {
    await startSuggesting(page);
    await doc.toolbar('Bullet list').click();
    await expect(page.getByTestId('suggesting-notice')).toContainText('cannot be suggested');
    // The paragraph became a list item all the same.
    await expect(doc.pages.first().locator('.list-item, li').first()).toBeVisible();
    // Switching mode clears the notice.
    await setMode(page, 'Editing');
    await expect(page.getByTestId('suggesting-notice')).toHaveCount(0);
  });

  test('D-SUG-06 text typed in the middle of a deletion is an insertion of its own, not part of the deletion', async ({
    page,
  }) => {
    await startSuggesting(page);
    for (let i = 0; i < 5; i++) await page.keyboard.press('Backspace');
    await doc.settle();
    await expect(deletions().first()).toContainText('world');
    // The caret is before the deleted "world": move two characters into it and type.
    await doc.press('ArrowRight');
    await doc.press('ArrowRight');
    await doc.typeText('ZZ');
    await doc.settle();
    await expect(insertions().filter({ hasText: 'ZZ' })).toHaveCount(1);
    expect((await deletions().allTextContents()).join('')).toBe('world');
    await expect(doc.pages.first()).toContainText('woZZrld');
  });

  test('D-SUG-07 in the Final view nothing can be typed, also in Suggesting mode, and typing is a suggestion again in All markup', async ({
    page,
  }) => {
    await startSuggesting(page);
    await doc.typeText('!');
    await expect(insertions().filter({ hasText: '!' })).toHaveCount(1);
    await doc.toolbar('Review changes').click();
    const view = page.getByTestId('review-panel').getByRole('group', { name: 'Show changes as' });
    await view.getByRole('button', { name: 'Final' }).click();
    const banner = page.getByTestId('revision-view-banner');
    await expect(banner).toBeVisible();
    // A read-only editor takes no focus: click the text and send the keys anyway.
    await doc.pages.getByText('Hello world').first().click();
    await page.keyboard.press('Home');
    await page.keyboard.type('ZZ');
    await expect(doc.pages.first()).not.toContainText('ZZ');
    await banner.getByRole('button', { name: 'Switch to All markup' }).click();
    await expect(banner).toHaveCount(0);
    await doc.clickText('Hello world');
    await doc.press('Home');
    await doc.typeText('ZZ');
    await expect(insertions().filter({ hasText: 'ZZ' })).toHaveCount(1);
  });
});
