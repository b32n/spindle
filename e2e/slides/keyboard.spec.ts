import { test, expect } from '@playwright/test';
import { SlidesPage } from '../support/slides';

// Slides keyboard-only flows. Toolbar buttons are focused programmatically (the keyboard equivalent
// of tabbing to them) because Tab cannot reach the toolbar at all (SL-KBD-03); everything after
// that is keys only.
test.use({ viewport: { width: 1900, height: 900 } });

test.describe('slides: keyboard-only flows', () => {
  let s: SlidesPage;
  const slideIds = (page: import('@playwright/test').Page) =>
    page.evaluate(() => (window as any).__spindle.deck.getSlideIds() as string[]);
  const activeId = (page: import('@playwright/test').Page) =>
    page.evaluate(() => (window as any).__spindle.deck.getActiveSlideId() as string);

  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
  });

  test('SL-KBD-01 add a slide, navigate the filmstrip, present and exit using only keys', async ({
    page,
  }) => {
    const before = await slideIds(page);
    expect(before).toHaveLength(3);
    await s.toolbar('New slide').focus();
    await page.keyboard.press('Enter');
    await expect.poll(async () => (await slideIds(page)).length).toBe(4);
    const ids = await slideIds(page);

    // Filmstrip: arrows move the active slide.
    // The new slide is inserted after the current one and becomes active.
    await expect.poll(() => activeId(page)).toBe(ids[1]);
    await s.thumbs.nth(1).focus();
    await page.keyboard.press('ArrowDown');
    await expect.poll(() => activeId(page)).toBe(ids[2]);
    await page.keyboard.press('ArrowDown');
    await expect.poll(() => activeId(page)).toBe(ids[3]);
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('ArrowUp');
    await expect.poll(() => activeId(page)).toBe(ids[0]);
    await expect(s.thumbs.first()).toBeFocused();

    // Present from the keyboard, step through, and leave with Escape.
    await s.toolbar('Present').focus();
    await page.keyboard.press('Enter');
    const counter = page.getByText(/^\d+ \/ \d+$/);
    await expect(counter).toHaveText('1 / 4');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Space');
    await expect(counter).toHaveText('3 / 4');
    await page.keyboard.press('End');
    await expect(counter).toHaveText('4 / 4');
    await page.keyboard.press('Escape');
    await expect(page.getByRole('button', { name: 'Exit (Esc)' })).toBeHidden();
    await expect(s.stage).toBeVisible();
  });

  test('SL-KBD-02 select and nudge an element, then delete it, with the keyboard', async ({
    page,
  }) => {
    // The first Tab lands on the editor root, which handles Tab itself: it selects the first element.
    await page.keyboard.press('Tab');
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).__spindle.deck.getSelection().elementIds.length)
      )
      .toBe(1);
    const sel = await page.evaluate(
      () => (window as any).__spindle.deck.getSelection().elementIds[0] as string
    );
    const start = (await s.els()).find((e) => e.id === sel)!;
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Shift+ArrowDown');
    await expect
      .poll(async () => {
        const e = (await s.els()).find((x) => x.id === sel)!;
        return [e.x - start.x, e.y - start.y];
      })
      .toEqual([1, 10]);
    await page.keyboard.press('Delete');
    await expect.poll(async () => (await s.els()).some((x) => x.id === sel)).toBe(false);
  });

  // Bug: the editor root swallows Tab to cycle the slide's elements, so the toolbar can never be
  // reached with Tab/Shift+Tab (a keyboard trap). https://github.com/b32n/spindle/issues/89
  test.fixme('SL-KBD-03 Tab reaches the toolbar buttons', async ({ page }) => {
    const present = s.toolbar('Present');
    let reached = false;
    for (let i = 0; i < 40 && !reached; i++) {
      await page.keyboard.press('Tab');
      reached = await present.evaluate((el) => el === document.activeElement);
    }
    expect(reached).toBe(true);
  });

  // Bug: leaving present mode drops focus to <body>, so the next keystroke goes nowhere.
  // https://github.com/b32n/spindle/issues/90
  test.fixme('SL-KBD-04 focus returns to the editor after leaving present mode', async ({
    page,
  }) => {
    await s.toolbar('Present').focus();
    await page.keyboard.press('Enter');
    await expect(page.getByRole('button', { name: 'Exit (Esc)' })).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.getByRole('button', { name: 'Exit (Esc)' })).toBeHidden();
    const inApp = await page.evaluate(() => document.activeElement !== document.body);
    expect(inApp).toBe(true);
  });
});
