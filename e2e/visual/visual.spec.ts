import { test, expect } from '@playwright/test';
import { SheetsPage, MOD } from '../support/sheets';
import { DocsPage } from '../support/docs';
import { SlidesPage } from '../support/slides';
import { prepareVisual, repaint } from '../support/visual';

// Chromium-only visual regression (the `visual` project in playwright.config.ts, opt-in via PW_VISUAL=1). Baselines are
// generated on the CI image (Linux) and committed; fonts are embedded (support/visual.ts) so the
// text does not depend on the machine's fonts. See docs/testing for how to update them.
test.describe('visual', () => {
  test.beforeEach(async ({ page }) => {
    await prepareVisual(page);
  });

  test('V-SHEETS-01 toolbar', async ({ page }) => {
    const s = new SheetsPage(page);
    await s.goto();
    await repaint(page);
    await expect(page).toHaveScreenshot('sheets-toolbar.png', {
      clip: { x: 0, y: 130, width: 1280, height: 60 },
    });
  });

  test('V-SHEETS-02 a formatted region', async ({ page }) => {
    const s = new SheetsPage(page);
    await s.goto();
    await s.enter(5, 0, 'Item');
    await s.enter(5, 1, 'Price');
    await s.enter(5, 2, 'Share');
    await s.enter(6, 0, 'Apples');
    await s.enter(6, 1, '1234.5');
    await s.enter(6, 2, '0.25');
    await s.enter(7, 0, 'Pears');
    await s.enter(7, 1, '99');
    await s.enter(7, 2, '0.75');
    await s.enter(8, 1, '=SUM(B7:B8)');
    // Header row: bold + centred. Price column: currency. Share column: percentage.
    await s.select(5, 0);
    await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press(`${MOD}+b`);
    await s.toolbar('Align center').click();
    await s.select(6, 1);
    await page.keyboard.press('Shift+ArrowDown');
    await page.keyboard.press('Shift+ArrowDown');
    await s.toolbar('Currency format').click();
    await s.select(6, 2);
    await page.keyboard.press('Shift+ArrowDown');
    await s.toolbar('Percentage format').click();
    await s.select(9, 4); // park the selection on an empty cell
    await repaint(page);
    await expect(page).toHaveScreenshot('sheets-formatted-region.png', {
      clip: { x: 0, y: 190, width: 700, height: 360 },
    });
  });

  test('V-DOCS-01 toolbar', async ({ page }) => {
    const d = new DocsPage(page);
    await d.goto();
    await repaint(page);
    await expect(page).toHaveScreenshot('docs-toolbar.png', {
      clip: { x: 0, y: 85, width: 1280, height: 70 },
    });
  });

  test('V-DOCS-02 a paginated page', async ({ page }) => {
    const d = new DocsPage(page);
    await d.goto();
    await repaint(page);
    await expect(d.pages.first()).toBeVisible();
    expect(await d.pages.count(), 'the sample document spans several pages').toBeGreaterThan(1);
    await expect(page.locator('.page').first()).toHaveScreenshot('docs-page-1.png');
  });

  test('V-SLIDES-01 toolbar', async ({ page }) => {
    const sl = new SlidesPage(page);
    await sl.goto();
    await repaint(page);
    await expect(page).toHaveScreenshot('slides-toolbar.png', {
      clip: { x: 0, y: 40, width: 1280, height: 56 },
    });
  });

  test('V-SLIDES-02 a slide with shapes', async ({ page }) => {
    const sl = new SlidesPage(page);
    await sl.goto();
    await sl.gotoSlide(2);
    await repaint(page);
    await expect(sl.slide).toHaveScreenshot('slides-shapes.png');
  });
});
