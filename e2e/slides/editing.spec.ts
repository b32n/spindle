import { test, expect } from '@playwright/test';
import { SlidesPage, MOD } from '../support/slides';

test.describe('slides: filmstrip, selection and element manipulation', () => {
  let s: SlidesPage;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
  });

  test('SL-LOAD-01 demo deck loads with three slides @smoke', async ({ page }) => {
    await expect(page.getByText('Spindle Slides demo')).toBeVisible();
    await expect(s.thumbs).toHaveCount(3);
    await expect(s.slide).toContainText('Spindle Slides');
    await expect(s.elements).toHaveCount(3);
  });

  test('SL-FLM-01 clicking a filmstrip thumbnail switches slides @smoke', async ({ page }) => {
    const ids = await s.slideIds();
    await page.getByRole('button', { name: 'Go to slide 2' }).click();
    await expect.poll(() => s.activeSlide()).toBe(ids[1]);
    await expect(s.slide).toContainText("What's in the box");
    await page.getByRole('button', { name: 'Go to slide 3' }).click();
    await expect(s.slide).toContainText('Shapes, images & lines');
  });

  test('SL-FLM-02 arrow keys navigate slides when nothing is selected', async ({ page }) => {
    const ids = await s.slideIds();
    await page.getByRole('button', { name: 'Go to slide 1' }).click();
    await page.keyboard.press('ArrowDown');
    await expect.poll(() => s.activeSlide()).toBe(ids[1]);
    await page.keyboard.press('ArrowUp');
    await expect.poll(() => s.activeSlide()).toBe(ids[0]);
  });

  test('SL-FLM-03 New slide adds a slide and makes it active', async ({ page }) => {
    await page.getByRole('button', { name: 'New slide' }).click();
    await expect(s.thumbs).toHaveCount(4);
    // The new slide is inserted after the active (first) slide.
    const ids = await s.slideIds();
    await expect.poll(() => s.activeSlide()).toBe(ids[1]);
  });

  test('SL-FLM-04 slide context menu duplicates and deletes a slide', async ({ page }) => {
    await s.thumbs.nth(1).click({ button: 'right' });
    await page.getByText('Duplicate slide', { exact: false }).first().click();
    await expect(s.thumbs).toHaveCount(4);
    await s.thumbs.nth(1).click({ button: 'right' });
    await page.getByText('Delete slide', { exact: false }).first().click();
    await expect(s.thumbs).toHaveCount(3);
  });

  test('SL-SEL-01 clicking an element selects it and shows handles @smoke', async ({ page }) => {
    const shape = await s.firstOfType('shape');
    await s.click(shape.id);
    await s.expectSelection([shape.id]);
    await expect(page.locator('[data-handle]')).toHaveCount(8);
    await expect(page.locator('[data-rotate]')).toHaveCount(1);
  });

  test('SL-SEL-02 clicking empty space or pressing Escape clears the selection', async ({
    page,
  }) => {
    const [first] = await s.els();
    await s.click(first.id);
    await s.expectSelection([first.id]);
    await page.keyboard.press('Escape');
    await s.expectSelection([]);
    await s.click(first.id);
    const p = await s.toPage(1240, 20);
    await page.mouse.click(p.x, p.y);
    await s.expectSelection([]);
  });

  test('SL-SEL-03 shift-click adds and removes elements from the selection', async ({ page }) => {
    // Slide 2 has several well-separated shapes.
    await s.gotoSlide(2);
    const shapes = (await s.els()).filter((e) => e.type === 'shape');
    await s.click(shapes[0].id);
    await s.click(shapes[1].id, ['Shift']);
    await s.expectSelection([shapes[0].id, shapes[1].id]);
    await s.click(shapes[2].id, ['Shift']);
    await s.expectSelection([shapes[0].id, shapes[1].id, shapes[2].id]);
    // (Not the element just clicked: two quick clicks on one element count as a double-click.)
    await s.click(shapes[1].id, ['Shift']);
    await s.expectSelection([shapes[0].id, shapes[2].id]);
  });

  test('SL-SEL-04 Ctrl/Cmd+A selects every element on the slide', async ({ page }) => {
    const els = await s.els();
    await s.click(els[0].id);
    await page.keyboard.press(`${MOD}+a`);
    await s.expectSelection(els.map((e) => e.id));
  });

  test('SL-SEL-05 dragging a marquee selects the elements it touches', async ({ page }) => {
    const els = await s.els();
    const text = els.find((e) => e.type === 'text')!;
    const a = await s.toPage(text.x - 20, text.y - 20);
    const b = await s.toPage(text.x + text.w + 20, text.y + text.h + 20);
    await s.drag({ x: a.x, y: a.y }, { x: b.x, y: b.y });
    expect((await s.selection()).includes(text.id)).toBe(true);
  });

  test('SL-MOV-01 dragging an element moves it by the pointer delta @smoke', async ({ page }) => {
    const shape = await s.firstOfType('shape');
    const c = await s.elementCenter(shape.id);
    const { scale } = await s.toPage(0, 0);
    await s.drag(c, { x: c.x + 100 * scale, y: c.y + 40 * scale });
    await expect
      .poll(async () => Math.round((await s.el(shape.id)).x - shape.x))
      .toBeGreaterThanOrEqual(90);
    const moved = await s.el(shape.id);
    expect(moved.x - shape.x).toBeCloseTo(100, -1);
    expect(moved.y - shape.y).toBeCloseTo(40, -1);
  });

  test('SL-MOV-02 arrow keys nudge by 1 and Shift+arrow by 10', async ({ page }) => {
    const shape = await s.firstOfType('shape');
    await s.click(shape.id);
    await page.keyboard.press('ArrowRight');
    await expect.poll(async () => (await s.el(shape.id)).x).toBe(shape.x + 1);
    await page.keyboard.press('Shift+ArrowDown');
    await expect.poll(async () => (await s.el(shape.id)).y).toBe(shape.y + 10);
  });

  test('SL-RSZ-01 dragging the SE handle resizes the element', async ({ page }) => {
    const shape = await s.firstOfType('shape');
    await s.click(shape.id);
    const handle = page.locator('[data-handle="se"]');
    const hb = (await handle.boundingBox())!;
    const { scale } = await s.toPage(0, 0);
    await s.drag(
      { x: hb.x + hb.width / 2, y: hb.y + hb.height / 2 },
      { x: hb.x + hb.width / 2 + 60 * scale, y: hb.y + hb.height / 2 + 30 * scale }
    );
    await expect
      .poll(async () => Math.round((await s.el(shape.id)).w - shape.w))
      .toBeGreaterThan(40);
    const after = await s.el(shape.id);
    expect(after.x).toBeCloseTo(shape.x, 0);
    expect(after.h).toBeGreaterThan(shape.h);
  });

  test('SL-RSZ-02 the rotate handle rotates the element', async ({ page }) => {
    const shape = await s.firstOfType('shape');
    await s.click(shape.id);
    const rot = (await page.locator('[data-rotate]').boundingBox())!;
    const c = await s.elementCenter(shape.id);
    // Drag the handle a quarter turn clockwise around the element centre.
    const r = Math.hypot(rot.x + rot.width / 2 - c.x, rot.y + rot.height / 2 - c.y);
    await s.drag({ x: rot.x + rot.width / 2, y: rot.y + rot.height / 2 }, { x: c.x + r, y: c.y });
    await expect.poll(async () => Math.abs((await s.el(shape.id)).rotation)).toBeGreaterThan(45);
  });

  test('SL-DEL-01 Delete removes the selection; undo restores it @smoke', async ({ page }) => {
    const els = await s.els();
    await s.click(els[1].id);
    await page.keyboard.press('Delete');
    await expect(s.elements).toHaveCount(els.length - 1);
    await page.keyboard.press(`${MOD}+z`);
    await expect(s.elements).toHaveCount(els.length);
    await page.keyboard.press(`${MOD}+Shift+z`);
    await expect(s.elements).toHaveCount(els.length - 1);
  });

  test('SL-DUP-01 Ctrl/Cmd+D duplicates the selection', async ({ page }) => {
    const shape = await s.firstOfType('shape');
    const n = (await s.els()).length;
    await s.click(shape.id);
    await page.keyboard.press(`${MOD}+d`);
    await expect(s.elements).toHaveCount(n + 1);
    expect((await s.selection())[0]).not.toBe(shape.id);
  });

  test('SL-GRP-01 group and ungroup with Ctrl/Cmd+G', async ({ page }) => {
    await s.gotoSlide(2);
    const els = (await s.els()).filter((e) => e.type === 'shape');
    await s.click(els[0].id);
    await s.click(els[1].id, ['Shift']);
    await page.keyboard.press(`${MOD}+g`);
    const grouped = await page.evaluate(() => {
      const d = (window as any).__spindle.deck;
      return new Set(
        d
          .getElementsForSlide(d.getActiveSlideId())
          .map((e: any) => e.groupId)
          .filter(Boolean)
      ).size;
    });
    expect(grouped).toBe(1);
    // Clicking one member selects the whole group.
    await page.keyboard.press('Escape');
    await s.click(els[0].id);
    await expect.poll(async () => (await s.selection()).length).toBeGreaterThanOrEqual(2);
    await page.keyboard.press(`${MOD}+Shift+g`);
    const after = await page.evaluate(() => {
      const d = (window as any).__spindle.deck;
      return d.getElementsForSlide(d.getActiveSlideId()).filter((e: any) => e.groupId).length;
    });
    expect(after).toBe(0);
  });

  test('SL-ORD-01 bring to front / send to back change z-order', async ({ page }) => {
    const order = () =>
      page.evaluate(() => {
        const d = (window as any).__spindle.deck;
        return d.getElementIdsForSlide(d.getActiveSlideId());
      });
    const before: string[] = await order();
    await s.click(before[0]);
    await page.keyboard.press(`${MOD}+Shift+]`);
    await expect.poll(async () => (await order()).at(-1)).toBe(before[0]);
    await page.keyboard.press(`${MOD}+Shift+[`);
    await expect.poll(async () => (await order())[0]).toBe(before[0]);
  });

  test('SL-TAB-01 Tab cycles element selection', async ({ page }) => {
    const els = await s.els();
    await s.click(els[0].id);
    await page.keyboard.press('Tab');
    await s.expectSelection([els[1].id]);
  });
});
