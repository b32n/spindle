import { test, expect, type Page } from '@playwright/test';
import { SheetsPage } from '../support/sheets';
import { DocsPage } from '../support/docs';
import { SlidesPage } from '../support/slides';
import { trackExternalRequests } from '../support/fonts';

// With default settings Spindle may talk to exactly one non-local host: the font mirror it
// names as its default (fonts.b32n.io, see documentation/fonts.md). Anything else is a leak.
// Every external request is aborted (fonts get an empty stylesheet), so the suite stays offline.
const DEFAULT_MIRROR = 'https://fonts.b32n.io/';

async function expectOnlyDefaultMirror(external: string[]): Promise<void> {
  expect(external.filter((u) => !u.startsWith(DEFAULT_MIRROR))).toEqual([]);
}

/** Give any late request (images, fonts, beacons) a full render cycle to show up. */
async function settle(page: Page): Promise<void> {
  await page.evaluate(
    () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))
  );
  await page.waitForLoadState('networkidle');
}

test.describe('privacy: no third-party network with default settings', () => {
  test('P-NET-01 sheets: loading, the font picker and dialogs only ever reach the default font mirror', async ({
    page,
  }) => {
    const external = await trackExternalRequests(page, { cssHosts: ['fonts.b32n.io'] });
    const sheet = new SheetsPage(page);
    await sheet.goto();
    await sheet.select(0, 0);
    await page.getByRole('button', { name: /^Arial/ }).click();
    await expect(page.getByPlaceholder('Search fonts…')).toBeVisible();
    await page.keyboard.press('Escape');
    await sheet.enter(5, 0, '=SUM(1,2)');
    await sheet.expectValue(5, 0, 3);
    await settle(page);
    await expectOnlyDefaultMirror(external);
  });

  test('P-NET-02 docs: loading, typing and the image/link dialogs only ever reach the default font mirror', async ({
    page,
  }) => {
    const external = await trackExternalRequests(page, { cssHosts: ['fonts.b32n.io'] });
    const doc = new DocsPage(page);
    await doc.goto();
    await doc.clearAndFocus();
    await doc.typeText('hello');
    await doc.toolbar('Insert image').click();
    await expect(page.getByRole('button', { name: 'Upload' })).toBeVisible();
    await page.keyboard.press('Escape');
    await settle(page);
    await expectOnlyDefaultMirror(external);
  });

  test('P-NET-03 slides: loading, navigating slides and the insert menu only ever reach the default font mirror', async ({
    page,
  }) => {
    const external = await trackExternalRequests(page, { cssHosts: ['fonts.b32n.io'] });
    const s = new SlidesPage(page);
    await s.goto();
    await s.gotoSlide(2);
    await expect(s.toolbar('Image (upload)')).toBeVisible();
    await settle(page);
    await expectOnlyDefaultMirror(external);
  });

  for (const [id, app] of [
    ['P-NET-04', 'sheets'],
    ['P-NET-05', 'docs'],
    ['P-NET-06', 'slides'],
  ] as const) {
    test(`${id} ${app}: fontSource local makes no external request at all`, async ({ page }) => {
      const external = await trackExternalRequests(page);
      if (app === 'sheets') await new SheetsPage(page).goto('?fontSource=local');
      else if (app === 'docs') await new DocsPage(page).goto('?fontSource=local');
      else await new SlidesPage(page).goto('?fontSource=local');
      await settle(page);
      expect(external).toEqual([]);
    });
  }
});
