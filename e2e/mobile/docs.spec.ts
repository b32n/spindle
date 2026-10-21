import { test, expect } from '@playwright/test';
import { DocsPage } from '../support/docs';

test.describe('mobile: docs touch smoke', () => {
  // The docs editor is not supported on WebKit yet (see playwright.config.ts).
  test.skip(({ browserName }) => browserName === 'webkit', 'docs on WebKit is a known gap');

  test('M-DOCS-01 tapping a line places the caret and typing inserts text @mobile', async ({
    page,
  }) => {
    const doc = new DocsPage(page);
    await doc.goto();
    const line = doc.pages
      .locator('.line')
      .filter({ hasText: /^Welcome to Spindle Docs/ })
      .first();
    const box = (await line.boundingBox())!;
    await page.touchscreen.tap(box.x + box.width - 5, box.y + box.height / 2);
    await expect(doc.editor).toBeFocused();
    await doc.waitForCaretIn('Welcome to Spindle Docs');
    await doc.settle();
    await doc.press('End');
    await doc.typeText(' via touch');
    await expect(doc.editor.locator('h1').first()).toContainText(
      'Welcome to Spindle Docs via touch'
    );
  });
});
