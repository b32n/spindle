import { test, expect } from '@playwright/test';
import { DocsPage } from '../support/docs';
import { trackExternalRequests, FONT_LINKS } from '../support/fonts';

test.describe('docs: font source', () => {
  test('D-FNT-01 default: web fonts come from the Spindle mirror, never Google', async ({
    page,
  }) => {
    const external = await trackExternalRequests(page, { cssHosts: ['fonts.b32n.io'] });
    await new DocsPage(page).goto();
    await expect
      .poll(() => external.some((u) => u.startsWith('https://fonts.b32n.io/css2?family=')))
      .toBe(true);
    expect(external.filter((u) => !u.startsWith('https://fonts.b32n.io/'))).toEqual([]);
  });

  test("D-FNT-02 fontSource {kind:'local'}: zero font requests", async ({ page }) => {
    const external = await trackExternalRequests(page);
    await new DocsPage(page).goto('?fontSource=local');
    await expect(page.locator(FONT_LINKS)).toHaveCount(0);
    expect(external).toEqual([]);
  });

  test("D-FNT-03 fontSource {kind:'css', url}: fonts load from the host's own endpoint", async ({
    page,
  }) => {
    const external = await trackExternalRequests(page, { cssHosts: ['fonts.example.test'] });
    await new DocsPage(page).goto(
      `?fontSource=${encodeURIComponent('css:https://fonts.example.test/css2')}`
    );
    await expect
      .poll(() => external.some((u) => u.startsWith('https://fonts.example.test/css2?family=')))
      .toBe(true);
    expect(external.filter((u) => !u.startsWith('https://fonts.example.test/'))).toEqual([]);
  });

  test('D-FNT-04 unreachable font host: the editor still loads and types', async ({ page }) => {
    await trackExternalRequests(page, { fail: true });
    const doc = new DocsPage(page);
    await doc.goto();
    await expect(page.locator(FONT_LINKS)).toHaveCount(0);
    await doc.clearAndFocus();
    await doc.typeText('still works');
    await expect(doc.editor).toContainText('still works');
  });
});
