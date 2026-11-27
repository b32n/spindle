import { test, expect } from '@playwright/test';
import { SlidesPage } from '../support/slides';

test.describe('slides: shape adjustment handles', () => {
  let s: SlidesPage;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
    await s.gotoSlide(2);
  });

  async function insert(page: import('@playwright/test').Page, label: string): Promise<string> {
    const before = (await s.els()).map((e) => e.id);
    await s.toolbar('Shapes').click();
    await page.getByRole('button', { name: label, exact: true }).click();
    await expect.poll(async () => (await s.els()).length).toBe(before.length + 1);
    return (await s.els()).find((e) => !before.includes(e.id))!.id;
  }
  const adjustments = (page: import('@playwright/test').Page, id: string) =>
    page.evaluate((i) => (window as any).__spindle.deck.getElement(i).adjustments ?? null, id);

  test('SL-ADJ-01 dragging the handle of a rounded rectangle changes its corner radius and undo restores it', async ({
    page,
  }) => {
    const id = await insert(page, 'Rounded rectangle');
    await s.expectSelection([id]);
    expect(await adjustments(page, id)).toBeNull();
    const handle = s.stage.locator('[data-adjust="0"]');
    await expect(handle).toHaveCount(1);
    await expect(handle).toHaveAttribute('aria-label', 'Adjust shape');
    const hb = (await handle.boundingBox())!;
    const from = { x: hb.x + hb.width / 2, y: hb.y + hb.height / 2 };
    const pathBefore = await s.stage
      .locator(`[data-element-id="${id}"] svg path`)
      .first()
      .getAttribute('d');
    await s.drag(from, { x: from.x + 30, y: from.y });
    await expect
      .poll(async () => ((await adjustments(page, id)) as number[] | null)?.[0] ?? 0)
      .toBeGreaterThan(0.15);
    const radius = ((await adjustments(page, id)) as number[])[0];
    expect(radius).toBeLessThanOrEqual(0.5);
    await expect
      .poll(() => s.stage.locator(`[data-element-id="${id}"] svg path`).first().getAttribute('d'))
      .not.toBe(pathBefore);
    // one undo entry restores the default
    await s.slide.click({ position: { x: 5, y: 5 } });
    await page.evaluate(() => (window as any).__spindle.deck.undo());
    await expect.poll(() => adjustments(page, id)).toBeNull();
  });

  test('SL-ADJ-02 a callout shows a handle for its tail and a plain rectangle shows none', async ({
    page,
  }) => {
    const callout = await insert(page, 'Rectangular callout');
    await s.expectSelection([callout]);
    await expect(s.stage.locator('[data-adjust]')).toHaveCount(1);
    const hb = (await s.stage.locator('[data-adjust="0"]').boundingBox())!;
    const from = { x: hb.x + hb.width / 2, y: hb.y + hb.height / 2 };
    await s.drag(from, { x: from.x + 40, y: from.y + 40 });
    await expect
      .poll(async () => ((await adjustments(page, callout)) as number[] | null)?.length ?? 0)
      .toBe(2);
    // a selected shape adds a format bar that can push the toolbar's Shapes button into the overflow menu (browser-dependent), so add the plain rectangle through the model
    const rect = await page.evaluate(() => {
      const d = (window as any).__spindle.deck;
      const el = d.addElement(d.getActiveSlideId(), {
        type: 'shape',
        shape: 'rect',
        x: 100,
        y: 100,
        w: 200,
        h: 100,
      });
      d.setSelection({ slideId: d.getActiveSlideId(), elementIds: [el.id] });
      return el.id as string;
    });
    await s.expectSelection([rect]);
    await expect(s.stage.locator('[data-adjust]')).toHaveCount(0);
  });
});
