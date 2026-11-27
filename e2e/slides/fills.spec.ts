import { test, expect } from '@playwright/test';
import { SlidesPage } from '../support/slides';
import { closeToolbarMenu } from '../support/toolbar';

// The contextual shape format bar sits past the toolbar's overflow point at any width, so its fill controls are
// reached through the "More tools" menu when it overflows (browser-dependent); a wide window keeps the Background label inline.
test.use({ viewport: { width: 1900, height: 900 } });

test.describe('slides: gradient and picture fills', () => {
  let s: SlidesPage;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
    await s.gotoSlide(2);
  });

  test('SL-FIL-01 a gradient preset fills the selected shape and the slide background', async ({
    page,
  }) => {
    const before = (await s.els()).map((e) => e.id);
    await s.toolbar('Shapes').click();
    await page.getByRole('button', { name: 'star', exact: false }).first().click();
    await expect.poll(async () => (await s.els()).length).toBe(before.length + 1);
    const created = (await s.els()).find((e) => !before.includes(e.id))!;

    // The overflow trigger only exists when something collapsed (Firefox's font metrics fit the whole bar inline).
    const more = page.getByRole('button', { name: 'More tools', exact: true });
    if ((await more.count()) > 0) await more.click();
    await page.getByTestId('shape-fill-gradient').selectOption('horizontal');
    await closeToolbarMenu(page);
    const el = s.stage.locator(`[data-element-id="${created.id}"]`);
    // The shape paints through an SVG paint server, so the DOM carries the gradient and the path references it.
    const fill = await el.locator('svg path, svg ellipse').first().getAttribute('fill');
    expect(fill).toMatch(/^url\(#fill-/);
    await expect(el.locator('linearGradient')).toHaveCount(1);
    expect(
      await page.evaluate((i) => (window as any).__spindle.deck.getElement(i).fill.kind, created.id)
    ).toBe('gradient');

    // Slide background: computed style of the slide box.
    await page.getByTitle('Slide background').click({ button: 'right' });
    await page.getByTestId('slide-bg-gradient').selectOption('radial');
    await expect
      .poll(() => s.slide.evaluate((n) => getComputedStyle(n).backgroundImage))
      .toContain('radial-gradient');
  });
});
