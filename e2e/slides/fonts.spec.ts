import { test, expect } from '@playwright/test';
import { SlidesPage } from '../support/slides';
import { trackExternalRequests, FONT_LINKS } from '../support/fonts';

test.describe('slides: font source', () => {
  test('L-FNT-01 default: theme fonts come from the Spindle mirror, never Google', async ({
    page,
  }) => {
    const external = await trackExternalRequests(page, { cssHosts: ['fonts.b32n.io'] });
    await new SlidesPage(page).goto();
    await expect
      .poll(() => external.some((u) => u.startsWith('https://fonts.b32n.io/css2?family=Inter')))
      .toBe(true);
    expect(external.filter((u) => !u.startsWith('https://fonts.b32n.io/'))).toEqual([]);
  });

  test("L-FNT-02 fontSource {kind:'local'}: zero font requests", async ({ page }) => {
    const external = await trackExternalRequests(page);
    await new SlidesPage(page).goto('?fontSource=local');
    await expect(page.locator(FONT_LINKS)).toHaveCount(0);
    expect(external).toEqual([]);
  });

  test("L-FNT-03 fontSource {kind:'css', url}: fonts load from the host's own endpoint", async ({
    page,
  }) => {
    const external = await trackExternalRequests(page, { cssHosts: ['fonts.example.test'] });
    await new SlidesPage(page).goto(
      `?fontSource=${encodeURIComponent('css:https://fonts.example.test/css2')}`
    );
    await expect
      .poll(() =>
        external.some((u) => u.startsWith('https://fonts.example.test/css2?family=Inter'))
      )
      .toBe(true);
    expect(external.filter((u) => !u.startsWith('https://fonts.example.test/'))).toEqual([]);
  });

  test('L-FNT-04 unreachable font host: the deck still renders', async ({ page }) => {
    await trackExternalRequests(page, { fail: true });
    const s = new SlidesPage(page);
    await s.goto();
    await expect(page.locator(FONT_LINKS)).toHaveCount(0);
    await expect(s.slide).toBeVisible();
  });
});
