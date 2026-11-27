import { test, expect, type Page } from '@playwright/test';
import { SlidesPage, MOD } from '../support/slides';
import { closeToolbarMenu, toolbarButton } from '../support/toolbar';

const LONG =
  'Shrink to fit keeps every word of a long paragraph inside its box by scaling the text down until it fits, so no ' +
  'text is ever hidden or clipped. '.repeat(3);

/** Open the "Text box options" panel: an inline popover, or a modal opened from the "More tools" overflow menu. */
async function openTextOptions(page: Page): Promise<void> {
  await toolbarButton(page, 'Text box options').click();
  await expect(page.getByTestId('text-autofit')).toBeVisible();
}

/** Close the options panel (modal Close button, or the popover's backdrop) and any overflow menu left open. */
async function closeTextOptions(page: Page): Promise<void> {
  const dialog = page.getByRole('dialog', { name: 'Text box options' });
  if (await dialog.isVisible()) {
    await dialog.getByRole('button', { name: 'Close' }).click();
    await dialog.waitFor({ state: 'hidden' });
  } else {
    await page.mouse.click(2, 2);
    await expect(page.getByTestId('text-autofit')).toBeHidden();
  }
  await closeToolbarMenu(page);
}

test.describe('slides: text body fidelity (autofit, insets, direction)', () => {
  let s: SlidesPage;
  let id: string;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
    // A fresh empty text box (300x80) so the scenarios do not depend on the demo deck's copy.
    id = await page.evaluate(() => {
      const deck = (window as any).__spindle.deck;
      const slideId = deck.getActiveSlideId();
      const el = deck.addElement(slideId, { type: 'text', x: 100, y: 500, w: 300, h: 80 });
      deck.setSelection({ slideId, elementIds: [el.id] });
      return el.id as string;
    });
    await s.expectSelection([id]);
  });

  const body = (page: Page) => s.stage.locator(`[data-element-id="${id}"] [data-text-body]`);
  const model = (page: Page) =>
    page.evaluate((i) => (window as any).__spindle.deck.getElement(i), id) as Promise<{
      h: number;
      bodyStyle?: any;
    }>;

  async function typeText(page: Page, text: string): Promise<void> {
    const c = await s.elementCenter(id);
    await page.mouse.dblclick(c.x, c.y);
    await expect(page.locator('.ProseMirror[contenteditable="true"]')).toBeFocused();
    await page.keyboard.press(`${MOD}+a`);
    await page.keyboard.insertText(text);
    await page.keyboard.press('Escape');
    await expect(page.locator('.ProseMirror[contenteditable="true"]')).toHaveCount(0);
  }

  test('SL-TXT-07 shrink on overflow keeps a long text inside its box and stores the scale', async ({
    page,
  }) => {
    await openTextOptions(page);
    await page.getByTestId('text-autofit').selectOption('shrink');
    await closeTextOptions(page);
    await typeText(page, LONG);

    const box = body(page);
    await expect(box).toContainText('Shrink to fit keeps every word');
    // Nothing overflows the box: the text was scaled down, not clipped.
    await expect.poll(() => box.evaluate((n) => n.scrollHeight <= n.clientHeight + 1)).toBe(true);
    const { bodyStyle } = await model(page);
    expect(bodyStyle.autofit.mode).toBe('shrink');
    expect(bodyStyle.autofit.fontScale).toBeLessThan(1);
    expect(bodyStyle.autofit.fontScale).toBeGreaterThanOrEqual(0.25);
    // The stored scale is applied as stored (a reload of the same data renders the same way).
    const fs = await box.evaluate((n) => (n as HTMLElement).style.getPropertyValue('--spindle-fs'));
    expect(Number(fs)).toBeCloseTo(bodyStyle.autofit.fontScale, 5);
  });

  test('SL-TXT-08 resize shape to fit text grows the box; insets and direction apply', async ({
    page,
  }) => {
    const before = (await model(page)).h;
    await openTextOptions(page);
    await page.getByTestId('text-autofit').selectOption('resize');
    await page.getByTestId('text-inset-l').fill('30');
    await page.getByTestId('text-vert').selectOption('horz');
    await closeTextOptions(page);
    await expect.poll(async () => (await model(page)).bodyStyle?.insets?.l).toBe(30);
    await expect(body(page)).toHaveCSS('padding-left', '30px');

    await typeText(page, LONG);
    await expect.poll(async () => (await model(page)).h).toBeGreaterThan(before);
    // The box now holds the text: no overflow and nothing clipped.
    const box = body(page);
    await expect.poll(() => box.evaluate((n) => n.scrollHeight <= n.clientHeight + 1)).toBe(true);
    expect(await box.evaluate((n) => getComputedStyle(n.parentElement!).overflow)).toBe('visible');

    // Text direction: rotated text switches the body's writing mode.
    await openTextOptions(page);
    await page.getByTestId('text-vert').selectOption('vert');
    await closeTextOptions(page);
    await expect(body(page)).toHaveCSS('writing-mode', 'vertical-rl');
  });
});
