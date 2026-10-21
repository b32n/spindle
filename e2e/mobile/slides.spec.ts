import { test, expect, type Page } from '@playwright/test';
import { SlidesPage } from '../support/slides';

/** On a phone the filmstrip is a drawer behind the "Slides" button. */
async function openSlide2(page: Page, s: SlidesPage) {
  await page.getByRole('button', { name: 'Slides panel' }).tap();
  await page.getByRole('button', { name: 'Go to slide 2' }).tap();
  await expect(s.slide).toHaveAttribute('data-slide-id', (await s.slideIds())[1]);
  // Close the drawer if the thumbnail tap left it open over the stage.
  const slidesBtn = page.getByRole('button', { name: 'Slides panel' });
  if (await page.getByRole('button', { name: 'Go to slide 2' }).isVisible()) await slidesBtn.tap();
}

test.describe('mobile: slides touch smoke', () => {
  test('M-SLIDES-01 tapping an element selects it, tapping the background clears it @mobile', async ({
    page,
  }) => {
    const s = new SlidesPage(page);
    await s.goto();
    await openSlide2(page, s);
    const shape = await s.firstOfType('shape');
    const c = await s.elementCenter(shape.id);
    await page.touchscreen.tap(c.x, c.y);
    await s.expectSelection([shape.id]);
    const empty = await s.toPage(60, 650);
    await page.touchscreen.tap(empty.x, empty.y);
    await s.expectSelection([]);
  });

  test('M-SLIDES-02 dragging a selected element with a finger moves it @mobile', async ({
    page,
    browserName,
  }) => {
    test.skip(browserName !== 'chromium', 'raw touch events need CDP (Chromium only)');
    const s = new SlidesPage(page);
    await s.goto();
    await openSlide2(page, s);
    const shape = await s.firstOfType('shape');
    const c = await s.elementCenter(shape.id);
    await page.touchscreen.tap(c.x, c.y);
    await s.expectSelection([shape.id]);

    const cdp = await page.context().newCDPSession(page);
    const touch = (type: 'touchStart' | 'touchMove' | 'touchEnd', x: number, y: number) =>
      cdp.send('Input.dispatchTouchEvent', {
        type,
        touchPoints: type === 'touchEnd' ? [] : [{ x, y }],
      });
    const dx = 80;
    const dy = 40;
    await touch('touchStart', c.x, c.y);
    for (let i = 1; i <= 8; i++) await touch('touchMove', c.x + (dx * i) / 8, c.y + (dy * i) / 8);
    await touch('touchEnd', 0, 0);

    const { scale } = await s.toPage(0, 0);
    await expect
      .poll(async () => (await s.el(shape.id)).x)
      .toBeGreaterThan(shape.x + (dx / scale) * 0.6);
    const moved = await s.el(shape.id);
    expect(moved.y).toBeGreaterThan(shape.y + (dy / scale) * 0.6);
  });
});
