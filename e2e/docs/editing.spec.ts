import { test, expect } from '@playwright/test';
import { DocsPage, MOD } from '../support/docs';

test.describe('docs: typing, selection & formatting', () => {
  let doc: DocsPage;
  test.beforeEach(async ({ page }) => {
    doc = new DocsPage(page);
    await doc.goto();
  });

  test('D-TXT-01 demo document renders with header/footer placeholders @smoke', async ({
    page,
  }) => {
    await expect(page.getByRole('heading', { name: /Spindle Docs Demo/ })).toBeVisible();
    await expect(doc.pages.first()).toContainText('Welcome to Spindle Docs');
    await expect(page.locator('.page-footer').first()).toContainText('Page 1 of');
  });

  test('D-TXT-02 typing inserts text at the caret @smoke', async ({ page }) => {
    await doc.clickText('Welcome to Spindle Docs');
    await doc.press('End');
    await page.keyboard.type(' XYZ');
    await expect(doc.editor.locator('h1')).toHaveText('Welcome to Spindle Docs XYZ');
    await expect(doc.pages.first()).toContainText('Welcome to Spindle Docs XYZ');
  });

  test('D-TXT-03 Enter splits a paragraph; Backspace merges it back', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('onetwo');
    await doc.press('ArrowLeft');
    await doc.press('ArrowLeft');
    await doc.press('ArrowLeft');
    await page.keyboard.press('Enter');
    await expect(doc.editor.locator('p')).toHaveText(['one', 'two']);
    await page.keyboard.press('Backspace');
    await expect(doc.editor.locator('p')).toHaveText(['onetwo']);
  });

  test('D-SEL-01 Ctrl/Cmd+A then Delete empties the document', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('x');
    await expect(doc.editor).toContainText('x');
    await expect(doc.editor).not.toContainText('Features');
  });

  test('D-SEL-02 double-click selects a word; typing replaces it', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('alpha beta gamma');
    await doc.pages.getByText('alpha beta gamma').dblclick();
    await doc.settle();
    await page.keyboard.type('X');
    await expect(doc.editor).toContainText('X');
    const t = await doc.text();
    expect(t).toMatch(/^(X beta gamma|alpha X gamma|alpha beta X)$/);
  });

  test('D-SEL-03 Shift+Arrow extends the selection', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('abcdef');
    await doc.press('Home');
    await doc.press('Shift+ArrowRight');
    await doc.press('Shift+ArrowRight');
    await page.keyboard.press(`${MOD}+b`);
    await expect(doc.editor.locator('strong')).toHaveText('ab');
  });

  test('D-FMT-01 Bold / Italic / Underline / Strike via shortcuts', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('word');
    await doc.press(`${MOD}+a`);
    await page.keyboard.press(`${MOD}+b`);
    await page.keyboard.press(`${MOD}+i`);
    await page.keyboard.press(`${MOD}+u`);
    await expect(doc.editor.locator('strong')).toContainText('word');
    await expect(doc.editor.locator('em')).toContainText('word');
    await expect(doc.editor.locator('u')).toContainText('word');
  });

  // https://github.com/b32n/spindle/issues/71: Word / Google Docs add the mark to a partly-marked selection.
  test('D-FMT-12 Ctrl+I on a partly-italic selection italicises all of it', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('aa bb');
    await page.keyboard.press('Shift+ArrowLeft');
    await page.keyboard.press('Shift+ArrowLeft');
    await doc.settle();
    await page.keyboard.press(`${MOD}+i`);
    await expect(doc.editor.locator('em')).toHaveText('bb');
    await doc.press(`${MOD}+a`);
    await page.keyboard.press(`${MOD}+i`);
    await expect(doc.editor.locator('em')).toHaveText('aa bb');
  });

  test('D-FMT-02 toolbar Bold toggles on and off @smoke', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('word');
    await doc.press(`${MOD}+a`);
    await doc.toolbar('Bold').click();
    await expect(doc.editor.locator('strong')).toHaveText('word');
    await doc.toolbar('Bold').click();
    await expect(doc.editor.locator('strong')).toHaveCount(0);
  });

  test('D-FMT-03 toolbar Italic, Underline, Strikethrough', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('word');
    await doc.press(`${MOD}+a`);
    await doc.toolbar('Italic').click();
    await doc.toolbar('Underline').click();
    await doc.toolbar('Strikethrough').click();
    await expect(doc.editor.locator('em')).toHaveText('word');
    await expect(doc.editor.locator('u')).toHaveText('word');
    await expect(doc.editor.locator('s, del, strike')).toHaveText('word');
  });

  test('D-FMT-04 heading styles from the Paragraph menu', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('Title');
    await doc.toolbar('Paragraph').click();
    await page.getByText('Heading 2', { exact: true }).click();
    await expect(doc.editor.locator('h2')).toHaveText('Title');
    await doc.toolbar(/Heading 2/).click();
    await page.getByText('Paragraph', { exact: true }).click();
    await expect(doc.editor.locator('h2')).toHaveCount(0);
  });

  test('D-FMT-05 text alignment buttons', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('aligned');
    await doc.toolbar('Align center').click();
    await expect(doc.editor.locator('p').first()).toHaveCSS('text-align', 'center');
    await doc.toolbar('Align right').click();
    await expect(doc.editor.locator('p').first()).toHaveCSS('text-align', 'right');
    await doc.toolbar('Justify').click();
    await expect(doc.editor.locator('p').first()).toHaveCSS('text-align', 'justify');
    await doc.toolbar('Align left').click();
    await expect(doc.editor.locator('p').first()).toHaveCSS('text-align', 'left');
  });

  test('D-FMT-06 bullet and numbered lists', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('one');
    await doc.toolbar('Bullet list').click();
    await expect(doc.editor.locator('ul > li')).toHaveCount(1);
    await page.keyboard.press('Enter');
    await expect(doc.editor.locator('ul > li')).toHaveCount(2);
    await doc.typeText('two');
    await doc.press(`${MOD}+a`);
    await doc.toolbar('Numbered list').click();
    await expect(doc.editor.locator('ol > li')).toHaveCount(2);
    await expect(doc.editor.locator('ul')).toHaveCount(0);
  });

  test('D-FMT-07 increase / decrease indent in a list', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('one');
    await doc.toolbar('Bullet list').click();
    await expect(doc.editor.locator('ul > li')).toHaveCount(1);
    await page.keyboard.press('Enter');
    await expect(doc.editor.locator('ul > li')).toHaveCount(2);
    await doc.typeText('two');
    await doc.toolbar('Increase indent').click();
    await expect(doc.editor.locator('li li')).toHaveCount(1);
    await doc.toolbar('Decrease indent').click();
    await expect(doc.editor.locator('li li')).toHaveCount(0);
  });

  test('D-FMT-08 font size increase applies to the selection', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('big');
    await doc.press(`${MOD}+a`);
    await doc.toolbar('Increase size').click();
    await expect(doc.editor.locator('span[style*="font-size"]')).toHaveCount(1);
  });

  test('D-FMT-09 undo and redo via toolbar and shortcuts @smoke', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('hello');
    await expect(doc.editor).toContainText('hello');
    await doc.toolbar('Undo').click();
    await expect(doc.editor).not.toContainText('hello');
    await doc.toolbar('Redo').click();
    await expect(doc.editor).toContainText('hello');
    await page.keyboard.press(`${MOD}+z`);
    await expect(doc.editor).not.toContainText('hello');
  });

  test('D-FMT-10 horizontal line and page break from the More tools menu', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('before');
    await doc.toolbar('More tools').click();
    await page.getByRole('menuitem', { name: 'Horizontal line' }).click();
    await expect(doc.editor.locator('hr')).toHaveCount(1);
  });

  test('D-FMT-11 superscript from the More tools menu', async ({ page }) => {
    await doc.clearAndFocus();
    await doc.typeText('x2');
    await doc.press(`${MOD}+a`);
    await doc.toolbar('More tools').click();
    await page.getByRole('menuitem', { name: 'Superscript' }).click();
    await expect(doc.editor.locator('sup')).toHaveText('x2');
  });
});
