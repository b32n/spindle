import { test, expect } from '@playwright/test';
import { SlidesPage, MOD } from '../support/slides';

// 1x1 transparent PNG
const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';

test.describe('slides: insert tools, text editing, zoom, notes, background', () => {
  let s: SlidesPage;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
    await s.gotoSlide(2); // slide with several shapes; leaves slide 1 placeholders alone
  });

  const count = async () => (await s.els()).length;

  /** Ids present now that were not in `before`. */
  const added = async (before: string[]) => (await s.els()).filter((e) => !before.includes(e.id));
  const ids = async () => (await s.els()).map((e) => e.id);

  // The toolbar's insert buttons add the element straight away at a default
  // position (no click/drag placement) and select it.
  test('SL-INS-01 Text box button inserts and selects a text box @smoke', async () => {
    const before = await ids();
    await s.toolbar('Text box').click();
    await expect.poll(async () => (await added(before)).length).toBe(1);
    const [created] = await added(before);
    expect(created.type).toBe('text');
    await s.expectSelection([created.id]);
    await expect(s.stage.locator(`[data-element-id="${created.id}"]`)).toBeVisible();
  });

  test('SL-INS-02 Enter on a freshly inserted text box starts editing; typing stores the text', async ({ page }) => {
    const before = await ids();
    await s.toolbar('Text box').click();
    await expect.poll(async () => (await added(before)).length).toBe(1);
    const [created] = await added(before);
    // Insertion selects the box; Enter starts editing its text.
    await page.keyboard.press('Enter');
    await expect(page.locator('.ProseMirror[contenteditable="true"]')).toBeFocused();
    await page.keyboard.type('Hello slides');
    await page.keyboard.press('Escape');
    await expect(s.stage.locator(`[data-element-id="${created.id}"]`)).toContainText('Hello slides');
  });

  test('SL-INS-03 Shapes picker inserts the chosen shape preset @smoke', async ({ page }) => {
    const before = await ids();
    await s.toolbar('Shapes').click();
    await page.getByRole('button', { name: 'star', exact: false }).first().click();
    await expect.poll(async () => (await added(before)).length).toBe(1);
    const [created] = await added(before);
    expect(created.type).toBe('shape');
  });

  test('SL-INS-04 Line and Arrow buttons insert line elements', async ({ page }) => {
    const before = await ids();
    await s.toolbar('Line').click();
    await expect.poll(async () => (await added(before)).length).toBe(1);
    await page.keyboard.press('Escape');
    await s.toolbar('Arrow').click();
    await expect.poll(async () => (await added(before)).length).toBe(2);
    expect((await added(before)).every((e) => e.type === 'line')).toBe(true);
  });

  test('SL-INS-05 Table (3x3) button inserts a table', async () => {
    await s.toolbar('Table (3×3)').click();
    await expect.poll(async () => (await s.els()).some((e) => e.type === 'table')).toBe(true);
  });

  test('SL-INS-06 Image from URL prompts for a URL and inserts an image', async ({ page }) => {
    const before = await ids();
    page.once('dialog', (d) => d.accept(PNG));
    await s.toolbar('Image from URL').click();
    await expect.poll(async () => (await added(before)).length).toBe(1);
    expect((await added(before))[0].type).toBe('image');
  });

  test('SL-INS-07 Image upload inserts an image from a chosen file', async ({ page }) => {
    const before = await ids();
    await page.locator('input[type="file"]').setInputFiles({
      name: 'pixel.png',
      mimeType: 'image/png',
      buffer: Buffer.from(PNG.split(',')[1], 'base64'),
    });
    await expect.poll(async () => (await added(before)).length).toBe(1);
    expect((await added(before))[0].type).toBe('image');
  });

  test('SL-INS-08 cancelling the image URL prompt inserts nothing', async ({ page }) => {
    const before = await ids();
    page.once('dialog', (d) => d.dismiss());
    await s.toolbar('Image from URL').click();
    expect((await s.els()).length).toBe(before.length);
  });

  test('SL-INS-09 insertion is undoable and redoable', async ({ page }) => {
    const before = await ids();
    await s.toolbar('Table (3×3)').click();
    await expect.poll(async () => (await s.els()).length).toBe(before.length + 1);
    await page.keyboard.press(`${MOD}+z`);
    await expect.poll(async () => (await s.els()).length).toBe(before.length);
    await page.keyboard.press(`${MOD}+Shift+z`);
    await expect.poll(async () => (await s.els()).length).toBe(before.length + 1);
  });

  test('SL-TXT-01 double-clicking a text element enters editing; typing edits it', async ({ page }) => {
    const text = await s.firstOfType('text');
    const c = await s.elementCenter(text.id);
    await page.mouse.dblclick(c.x, c.y);
    await expect(page.locator('.ProseMirror[contenteditable="true"]')).toBeFocused();
    await page.keyboard.press(`${MOD}+a`);
    await page.keyboard.type('Replaced title');
    await page.keyboard.press('Escape');
    await expect(s.stage.locator(`[data-element-id="${text.id}"]`)).toContainText('Replaced title');
  });

  test('SL-TXT-02 Enter on a selected text element starts editing', async ({ page }) => {
    const text = await s.firstOfType('text');
    await s.click(text.id);
    await page.keyboard.press('Enter');
    await expect(page.locator('.ProseMirror[contenteditable="true"]')).toBeFocused();
  });

  test('SL-TXT-03 bold via shortcut inside the text editor', async ({ page }) => {
    const text = (await s.els()).filter((e) => e.type === 'text')[1]; // body text (title is bold by theme)
    const c = await s.elementCenter(text.id);
    await page.mouse.dblclick(c.x, c.y);
    await expect(page.locator('.ProseMirror[contenteditable="true"]')).toBeFocused();
    await page.keyboard.press(`${MOD}+a`);
    await page.keyboard.press(`${MOD}+b`);
    await page.keyboard.press('Escape');
    await expect
      .poll(() => page.evaluate((id) => JSON.stringify((window as any).__spindle.deck.getElement(id).richText), text.id))
      .toMatch(/bold|strong/);
  });

  test('SL-ZOOM-01 zoom buttons rescale the slide', async ({ page }) => {
    const width = async () => (await s.slide.boundingBox())!.width;
    await s.toolbar('100%').click();
    await expect.poll(width).toBeCloseTo(1280, -1);
    await s.toolbar('50%').click();
    await expect.poll(width).toBeCloseTo(640, -1);
    await s.toolbar('200%').click();
    await expect.poll(width).toBeCloseTo(2560, -1);
    await s.toolbar('Fit').click();
    await expect.poll(width).toBeLessThan(1100);
  });

  test('SL-BGD-01 background colour input recolours the slide', async ({ page }) => {
    const bg = page.getByRole('textbox', { name: 'Background' });
    await bg.fill('#ff0000');
    await bg.press('Enter');
    await expect
      .poll(() => page.evaluate(() => {
        const d = (window as any).__spindle.deck;
        return JSON.stringify(d.getSlide(d.getActiveSlideId()).background ?? null);
      }))
      .toContain('ff0000');
  });

  test('SL-NTS-01 speaker notes are stored per slide', async ({ page }) => {
    const notes = page.getByRole('textbox', { name: 'Add notes for the presenter…' });
    await notes.click();
    await page.keyboard.type('Remember the demo');
    await expect(notes).toContainText('Remember the demo');
    await s.gotoSlide(1);
    await expect(notes).not.toContainText('Remember the demo');
    await s.gotoSlide(2);
    await expect(notes).toContainText('Remember the demo');
  });
});
