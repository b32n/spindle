import { test, expect } from '@playwright/test';
import { SlidesPage, MOD } from '../support/slides';

test.describe('slides: present mode', () => {
  let s: SlidesPage;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
  });

  const counter = (page: import('@playwright/test').Page) => page.getByText(/^\d+ \/ \d+$/);

  test('SL-PRS-01 Present opens full-screen present mode on the active slide @smoke', async ({
    page,
  }) => {
    await s.toolbar('Present').click();
    await expect(page.getByRole('button', { name: 'Exit (Esc)' })).toBeVisible();
    await expect(counter(page)).toHaveText('1 / 3');
    await expect(page.getByText('Press S for presenter panel')).toBeVisible();
  });

  test('SL-PRS-02 arrow keys, space, PageDown/Up, Home and End navigate', async ({ page }) => {
    await s.toolbar('Present').click();
    await expect(counter(page)).toHaveText('1 / 3');
    await page.keyboard.press('ArrowRight');
    await expect(counter(page)).toHaveText('2 / 3');
    await page.keyboard.press('Space');
    await expect(counter(page)).toHaveText('3 / 3');
    await page.keyboard.press('ArrowRight'); // clamps at the last slide
    await expect(counter(page)).toHaveText('3 / 3');
    await page.keyboard.press('PageUp');
    await expect(counter(page)).toHaveText('2 / 3');
    await page.keyboard.press('Home');
    await expect(counter(page)).toHaveText('1 / 3');
    await page.keyboard.press('End');
    await expect(counter(page)).toHaveText('3 / 3');
    await page.keyboard.press('ArrowLeft');
    await expect(counter(page)).toHaveText('2 / 3');
  });

  test('SL-PRS-03 clicking the stage advances; typing a number then Enter jumps', async ({
    page,
  }) => {
    await s.toolbar('Present').click();
    await expect(counter(page)).toHaveText('1 / 3');
    await page.mouse.click(640, 300);
    await expect(counter(page)).toHaveText('2 / 3');
    await page.keyboard.press('3');
    await page.keyboard.press('Enter');
    await expect(counter(page)).toHaveText('3 / 3');
    await page.keyboard.press('1');
    await page.keyboard.press('Enter');
    await expect(counter(page)).toHaveText('1 / 3');
  });

  test('SL-PRS-04 Escape and the Exit button leave present mode', async ({ page }) => {
    await s.toolbar('Present').click();
    await expect(page.getByRole('button', { name: 'Exit (Esc)' })).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.getByRole('button', { name: 'Exit (Esc)' })).toBeHidden();
    await s.toolbar('Present').click();
    await page.getByRole('button', { name: 'Exit (Esc)' }).click();
    await expect(page.getByRole('button', { name: 'Exit (Esc)' })).toBeHidden();
    await expect(s.slide).toBeVisible();
  });

  test('SL-PRS-05 presenting starts from the currently selected slide', async ({ page }) => {
    await s.gotoSlide(2);
    await s.toolbar('Present').click();
    await expect(counter(page)).toHaveText('2 / 3');
  });

  test('SL-PRS-06 S toggles the presenter notes panel', async ({ page }) => {
    await s.gotoSlide(1);
    const notes = page.getByRole('textbox', { name: 'Add notes for the presenter…' });
    await notes.click();
    await page.keyboard.type('Say hello');
    await s.toolbar('Present').click();
    await page.keyboard.press('s');
    await expect(page.getByText(/Speaker notes · slide 1/)).toBeVisible();
    await expect(page.getByText('Say hello').last()).toBeVisible();
    await page.keyboard.press('s');
    await expect(page.getByText(/Speaker notes · slide 1/)).toBeHidden();
  });
});

test.describe('slides: comments, clipboard, read-only', () => {
  let s: SlidesPage;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
  });

  test('SL-CMT-01 the Comments button toggles the comments panel', async ({ page }) => {
    await s.toolbar('Comments').click();
    await expect(page.getByText('Select an element to add a comment.')).toBeVisible();
    await expect(page.getByText('No comments yet.')).toBeVisible();
    await page.getByRole('button', { name: 'Comments' }).first().click();
    await expect(page.getByText('No comments yet.')).toBeHidden();
  });

  test('SL-CMT-02 select an element and add a comment on it', async ({ page }) => {
    await s.toolbar('Comments').click();
    const [first] = await s.els();
    await s.click(first.id);
    const box = page.getByPlaceholder(/comment/i).first();
    await box.fill('Please tighten the title');
    await page
      .getByRole('button', { name: /^(Comment|Post|Add|Send|Reply)/ })
      .first()
      .click();
    await expect(page.getByText('Please tighten the title')).toBeVisible();
  });

  test('SL-CLP-01 copy and paste an element (keyboard)', async ({ page, browserName }) => {
    test.skip(
      browserName === 'webkit',
      'WebKit headless does not expose the OS clipboard to keyboard copy/paste'
    );
    await s.gotoSlide(2);
    const shape = await s.firstOfType('shape');
    const n = (await s.els()).length;
    await s.click(shape.id);
    await page.keyboard.press(`${MOD}+c`);
    await page.keyboard.press(`${MOD}+v`);
    await expect.poll(async () => (await s.els()).length).toBe(n + 1);
  });

  test('SL-CLP-02 cut removes the element and paste brings it back', async ({
    page,
    browserName,
  }) => {
    test.skip(
      browserName === 'webkit',
      'WebKit headless does not expose the OS clipboard to keyboard copy/paste'
    );
    await s.gotoSlide(2);
    const shape = await s.firstOfType('shape');
    const n = (await s.els()).length;
    await s.click(shape.id);
    await page.keyboard.press(`${MOD}+x`);
    await expect.poll(async () => (await s.els()).length).toBe(n - 1);
    await page.keyboard.press(`${MOD}+v`);
    await expect.poll(async () => (await s.els()).length).toBe(n);
  });

  test('SL-RO-01 read-only mode hides editing chrome and blocks edits @smoke', async ({ page }) => {
    await s.goto('?readonly=1');
    await expect(s.toolbar('Text box')).toBeHidden();
    await expect(page.getByRole('button', { name: 'New slide' })).toBeHidden();
    const els = await s.els();
    const c = await s.elementCenter(els[0].id);
    await page.mouse.click(c.x, c.y);
    await page.keyboard.press('Delete');
    await page.keyboard.press(`${MOD}+a`);
    await page.keyboard.press('Backspace');
    expect((await s.els()).length).toBe(els.length);
  });

  test('SL-RO-02 read-only can still change slides in the filmstrip', async ({ page }) => {
    await s.goto('?readonly=1');
    await page.getByRole('button', { name: 'Go to slide 2' }).click();
    await expect(s.slide).toContainText("What's in the box");
  });
});
