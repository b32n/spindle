import { test, expect } from '@playwright/test';
import { SlidesPage } from '../support/slides';
import { fixture } from '../support/file-io';
import { closeToolbarMenu, toolbarButton } from '../support/toolbar';

test.describe('slides: image crop', () => {
  let s: SlidesPage;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
    await s.gotoSlide(2);
  });

  test('SL-IMG-15 cropping an image clips it in the DOM and resetting restores the plain image', async ({
    page,
  }) => {
    const before = (await s.els()).map((e) => e.id);
    await page.locator('input[type="file"]').setInputFiles(fixture('logo.png'));
    await expect.poll(async () => (await s.els()).length).toBe(before.length + 1);
    const id = (await s.els()).find((e) => !before.includes(e.id))!.id;
    await s.expectSelection([id]);
    const el = s.stage.locator(`[data-element-id="${id}"]`);
    await expect(el.locator('img')).toHaveCount(1);
    await expect(el.locator('[data-image-crop]')).toHaveCount(0);

    // The crop UI is an inline popover, or a modal opened from the "More tools" overflow menu.
    await toolbarButton(page, 'Crop image').click();
    await page.getByTestId('image-crop-l').fill('10');
    await page.getByTestId('image-crop-t').fill('20');
    await page.getByTestId('image-crop-r').fill('30');

    const clip = el.locator('[data-image-crop]');
    await expect(clip).toHaveAttribute('data-image-crop', '10,20,30,0');
    expect(await clip.evaluate((n) => getComputedStyle(n).overflow)).toBe('hidden');
    // The whole image is scaled up by 1 / visible share and offset by the trimmed share of itself.
    const img = clip.locator('img');
    const box = await img.evaluate((n) => ({
      w: parseFloat((n as HTMLElement).style.width),
      h: parseFloat((n as HTMLElement).style.height),
      l: parseFloat((n as HTMLElement).style.left),
      t: parseFloat((n as HTMLElement).style.top),
    }));
    expect(box.w).toBeCloseTo(100 / 0.6, 1);
    expect(box.h).toBeCloseTo(125, 1);
    expect(box.l).toBeCloseTo(-10 / 0.6, 1);
    expect(box.t).toBeCloseTo(-25, 1);
    expect(
      await page.evaluate((i) => (window as any).__spindle.deck.getElement(i).crop, id)
    ).toEqual({ l: 10, t: 20, r: 30, b: 0 });

    // Reset through the popover: an all-zero crop renders the plain image again.
    await page.getByTestId('image-crop-reset').click();
    await expect(el.locator('[data-image-crop]')).toHaveCount(0);
    await expect(el.locator('img')).toHaveCount(1);
    // Overflow path: the fields live in a modal that must be dismissed before the menu.
    const dialog = page.getByRole('dialog', { name: 'Crop image' });
    if (await dialog.isVisible()) {
      await dialog.getByRole('button', { name: 'Close' }).click();
      await dialog.waitFor({ state: 'hidden' });
    }
    await closeToolbarMenu(page);
  });
});
