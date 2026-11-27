import { test, expect, type Page } from '@playwright/test';
import { SlidesPage } from '../support/slides';
import { toolbarButton } from '../support/toolbar';

const deckSize = (page: Page) =>
  page.evaluate(() => (window as any).__spindle.deck.getSlideSize() as { w: number; h: number });

/**
 * The theme popover, which also holds the slide size section. Inline it opens from the theme button (named after the
 * current theme); in the overflow menu it is the "Theme" row, which opens a modal (browser-dependent toolbar overflow).
 */
async function openThemePanel(page: Page) {
  const themeName = await page.evaluate(
    () => (window as any).__spindle.deck.getTheme().name as string
  );
  const inline = page.getByRole('button', { name: themeName, exact: true });
  if ((await inline.count()) > 0) {
    await inline.click();
  } else {
    await page.getByRole('button', { name: 'More tools', exact: true }).click();
    const menu = page.getByRole('menu');
    await menu
      .getByRole('button', { name: /^Theme/ })
      .or(menu.getByRole('menuitem', { name: /^Theme/ }))
      .first()
      .click();
  }
  await expect(page.getByTestId('slide-size-section')).toBeVisible();
}

test.describe('slides: layouts, slide size and hidden slides', () => {
  let s: SlidesPage;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
  });

  test('SL-HID-01 a slide hidden from the filmstrip menu is marked, skipped in present mode, and shown again', async ({
    page,
  }) => {
    const ids = await s.slideIds();
    await s.thumbs.nth(1).click({ button: 'right' });
    await page.getByText('Hide slide', { exact: true }).click();
    await expect(page.getByTestId('slide-hidden-badge')).toHaveCount(1);
    await expect(page.getByRole('button', { name: 'Go to slide 2 (hidden)' })).toBeVisible();
    expect(
      await page.evaluate((id) => (window as any).__spindle.deck.getSlide(id).hidden, ids[1])
    ).toBe(true);

    // Present from slide 1: only two slides play and the next one is slide 3.
    await s.gotoSlide(1);
    await s.toolbar('Present').click();
    await expect(page.getByRole('button', { name: 'Exit (Esc)' })).toBeVisible();
    await expect(page.getByText(/^\d+ \/ \d+$/)).toHaveText('1 / 2');
    await page.keyboard.press('ArrowRight');
    await expect(page.getByText(/^\d+ \/ \d+$/)).toHaveText('2 / 2');
    await expect.poll(() => s.activeSlide()).toBe(ids[2]);
    await page.getByRole('button', { name: 'Exit (Esc)' }).click();

    await s.thumbs.nth(1).click({ button: 'right' });
    await page.getByText('Show slide', { exact: true }).click();
    await expect(page.getByTestId('slide-hidden-badge')).toHaveCount(0);
    expect(
      await page.evaluate((id) => (window as any).__spindle.deck.getSlide(id).hidden, ids[1])
    ).toBeFalsy();
  });

  test('SL-SIZ-01 changing the slide size to 4:3 keeps elements at their relative positions and is undoable', async ({
    page,
  }) => {
    await s.gotoSlide(2);
    const before = await s.els();
    const first = before[0];
    expect(await deckSize(page)).toEqual({ w: 1280, h: 720 });

    await openThemePanel(page);
    await page.getByLabel('Slide size', { exact: true }).selectOption('4:3');
    await expect.poll(() => deckSize(page)).toEqual({ w: 960, h: 720 });
    const after = (await s.els()).find((e) => e.id === first.id)!;
    expect(after.x).toBeCloseTo(first.x * 0.75, 1);
    expect(after.w).toBeCloseTo(first.w * 0.75, 1);
    expect(after.y).toBeCloseTo(first.y, 1);
    await expect(s.slide).toHaveCSS('width', '960px');

    // The popover (or modal) stays open after picking a size: dismiss it by clicking outside, then undo.
    await page.mouse.click(2, 2);
    await expect(page.getByTestId('slide-size-section')).toBeHidden();
    await toolbarButton(page, /^Undo/).click();
    await expect.poll(() => deckSize(page)).toEqual({ w: 1280, h: 720 });
    expect((await s.els()).find((e) => e.id === first.id)!.x).toBeCloseTo(first.x, 1);
  });

  test('SL-SIZ-02 a custom slide size is applied from the width and height fields', async ({
    page,
  }) => {
    await openThemePanel(page);
    await page.getByLabel('Slide width', { exact: true }).fill('1000');
    await page.getByLabel('Slide height', { exact: true }).fill('600');
    await page.getByRole('button', { name: 'Apply slide size' }).click();
    await expect.poll(() => deckSize(page)).toEqual({ w: 1000, h: 600 });
    await expect(page.getByLabel('Slide size', { exact: true })).toHaveValue('custom');
  });

  test('SL-LAY-02 static layout content is drawn on slides that use the layout and is not selectable', async ({
    page,
  }) => {
    const slideId = await page.evaluate(() => {
      const d = (window as any).__spindle.deck;
      const layout = {
        id: 'branded',
        name: 'Branded',
        masterName: 'Corporate',
        placeholders: [],
        elements: [
          {
            id: 'logo',
            containerId: '',
            type: 'shape',
            index: 'a0',
            x: 1100,
            y: 640,
            w: 120,
            h: 40,
            rotation: 0,
            shape: 'rect',
            fill: { kind: 'solid', color: { kind: 'rgb', hex: '#ff0000' } },
          },
        ],
      };
      d.setData({ ...d.getData(), layouts: [...d.getLayouts(), layout] });
      return d.addSlide({ layoutId: 'branded' }).id as string;
    });
    await expect.poll(async () => (await s.slideIds()).includes(slideId)).toBe(true);
    await page.evaluate((id) => (window as any).__spindle.deck.setActiveSlide(id), slideId);
    const logo = s.stage.locator('[data-layout-element-id="logo"]');
    await expect(logo).toHaveCount(1);
    expect(await logo.evaluate((n) => getComputedStyle(n).pointerEvents)).toBe('none');
    expect(await s.selection()).toEqual([]);
  });
});
