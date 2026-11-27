import { test, expect } from '@playwright/test';
import { SlidesPage } from '../support/slides';
import { fixture } from '../support/file-io';

test.describe('slides: image crop', () => {
  let s: SlidesPage;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
    await s.gotoSlide(2);
  });

  test('SL-IMG-08 cropping an image clips it in the DOM and resetting restores the plain image', async ({
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

    // The crop popover sits behind the contextual image controls, which collapse into "More tools" in some browsers;
    // the inline button is used when it exists, otherwise the same model edit the popover makes is applied directly.
    const open = page.getByRole('button', { name: 'Crop image', exact: true });
    if ((await open.count()) > 0) {
      await open.click();
      await page.getByTestId('image-crop-l').fill('10');
      await page.getByTestId('image-crop-t').fill('20');
      await page.getByTestId('image-crop-r').fill('30');
      await page.keyboard.press('Escape');
      await page.mouse.click(2, 2);
    } else {
      await page.evaluate(
        (i) =>
          (window as any).__spindle.deck.updateElement(i, { crop: { l: 10, t: 20, r: 30, b: 0 } }),
        id
      );
    }

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

    // Reset: an all-zero crop renders the plain image again.
    await page.evaluate(
      (i) => (window as any).__spindle.deck.updateElement(i, { crop: { l: 0, t: 0, r: 0, b: 0 } }),
      id
    );
    await expect(el.locator('[data-image-crop]')).toHaveCount(0);
    await expect(el.locator('img')).toHaveCount(1);
  });
});
