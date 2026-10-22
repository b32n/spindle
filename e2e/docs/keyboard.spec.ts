import { test, expect, type Locator, type Page } from '@playwright/test';
import { DocsPage, MOD } from '../support/docs';

// The toolbar collapses into "More tools" on narrow windows; keep every button in the tab order.
test.use({ viewport: { width: 1900, height: 900 } });

/** Press Shift+Tab until `target` holds focus. */
async function tabUntilBackwards(page: Page, target: Locator, max = 60): Promise<number> {
  for (let i = 1; i <= max; i++) {
    if (await target.evaluate((el) => el === document.activeElement).catch(() => false))
      return i - 1;
    await page.keyboard.press('Shift+Tab');
  }
  throw new Error(`Shift+Tab never reached the target in ${max} presses`);
}

/** Press Tab until `target` holds focus; returns how many presses it took. */
async function tabUntil(page: Page, target: Locator, max = 60): Promise<number> {
  for (let i = 1; i <= max; i++) {
    await page.keyboard.press('Tab');
    if (await target.evaluate((el) => el === document.activeElement).catch(() => false)) return i;
  }
  throw new Error(`Tab never reached the target in ${max} presses`);
}

test.describe('docs: keyboard-only flows', () => {
  let doc: DocsPage;
  test.beforeEach(async ({ page }) => {
    doc = new DocsPage(page);
    await doc.goto();
  });

  test('D-KBD-01 Tab reaches the toolbar and then the editor; the focused control has a visible indicator', async ({
    page,
  }) => {
    await page.keyboard.press('Tab');
    const first = page.locator(':focus');
    await expect(first).toBeVisible();
    const toolbarBtn = doc.toolbar(/^Bold/);
    await tabUntil(page, toolbarBtn);
    await expect(toolbarBtn).toBeFocused();
    const ring = await toolbarBtn.evaluate((el) => {
      const cs = getComputedStyle(el);
      return { outline: cs.outlineStyle, width: parseFloat(cs.outlineWidth), shadow: cs.boxShadow };
    });
    expect((ring.outline !== 'none' && ring.width > 0) || ring.shadow !== 'none').toBe(true);
    await tabUntil(page, doc.editor);
    await expect(doc.editor).toBeFocused();
  });

  /** Focus a toolbar button with Tab/Shift+Tab (from wherever focus is) and activate it with `key`. */
  async function pressToolbar(page: Page, name: string | RegExp, key = 'Enter') {
    const button = doc.toolbar(name);
    await page.keyboard.press('Shift+Tab'); // back out of the editor onto the toolbar
    await tabUntilBackwards(page, button);
    await page.keyboard.press(key);
  }

  test('D-KBD-02 write text, a bullet list and a link using only the keyboard', async ({
    page,
  }) => {
    await tabUntil(page, doc.editor);
    await doc.settle();
    await doc.press(`${MOD}+a`);
    await page.keyboard.press('Delete');
    await expect(doc.editor).not.toContainText('Features');
    await doc.typeText('Quarterly report');
    await doc.press('Enter');
    await doc.typeText('first');
    // Bullet list via the toolbar button (Shift+Tab out of the editor, Enter on the button), then
    // back into the editor with Tab to keep typing.
    await pressToolbar(page, 'Bullet list');
    await expect(doc.editor.locator('ul > li')).toHaveText(['first']);
    await tabUntil(page, doc.editor);
    await doc.settle();
    await page.keyboard.press('End');
    await doc.press('Enter');
    await doc.typeText('second');
    await expect(doc.editor.locator('ul > li')).toHaveText(['first', 'second']);

    // Select the word "second", bold it with the shortcut, open the link dialog with Mod+K and fill it in.
    await doc.press('Shift+Home');
    await page.keyboard.press(`${MOD}+b`);
    await expect(doc.editor.locator('strong')).toHaveText('second');
    await page.keyboard.press(`${MOD}+k`);
    const url = page.getByPlaceholder('https://example.com');
    await expect(url).toBeFocused();
    await page.keyboard.type('https://example.org');
    await page.keyboard.press('Enter');
    await expect(url).toBeHidden();
    await expect(doc.editor.locator('a[href="https://example.org"]')).toHaveText('second');
  });

  // Bug: no keyboard-only way to give a paragraph a heading style: the Paragraph dropdown items are
  // plain divs outside the tab order, Escape does not close it, and there is no shortcut; markdown
  // input rules ("# ") are implemented in docs-core but not installed by the editor.
  // https://github.com/b32n/spindle/issues/91
  test.fixme('D-KBD-03 a heading style can be applied from the keyboard', async ({ page }) => {
    await tabUntil(page, doc.editor);
    await doc.settle();
    await doc.press(`${MOD}+a`);
    await page.keyboard.press('Delete');
    await doc.typeText('Title');
    await page.keyboard.press('Shift+Tab');
    await tabUntilBackwards(page, doc.toolbar('Paragraph'));
    await page.keyboard.press('Enter');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('Enter');
    await expect(doc.editor.locator('h1, h2, h3')).toHaveText('Title');
  });

  test('D-KBD-04 Escape closes the link dialog and keeps the typed text', async ({ page }) => {
    await tabUntil(page, doc.editor);
    await doc.settle();
    await doc.press(`${MOD}+a`);
    await page.keyboard.press('Delete');
    await doc.typeText('plain');
    await page.keyboard.press(`${MOD}+k`);
    const url = page.getByPlaceholder('https://example.com');
    await expect(url).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(url).toBeHidden();
    await expect(doc.editor).toContainText('plain');
    await expect(doc.editor.locator('a')).toHaveCount(0);
  });
});
