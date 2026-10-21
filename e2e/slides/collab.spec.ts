import { test, expect, type Browser } from '@playwright/test';
import { SlidesPage, MOD } from '../support/slides';
import { roomName, waitForRoomSeed } from '../support/collab';

const WS = 'ws://127.0.0.1:1234';

async function peer(browser: Browser, room: string, user: string) {
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  const page = await context.newPage();
  const s = new SlidesPage(page);
  await s.goto(`?ws=${encodeURIComponent(WS)}&room=${room}&user=${user}`);
  await expect(page.getByText(/· connected$/)).toBeVisible();
  return { context, page, s };
}

test.describe('slides: two-client live collaboration', () => {
  test('SL-COL-01 an element moved by one client moves for the other @smoke', async ({
    browser,
  }, info) => {
    const room = roomName('slides', info.testId);
    const a = await peer(browser, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, room, 'Bob');
    await a.s.gotoSlide(2);
    await b.s.gotoSlide(2);
    const shape = await a.s.firstOfType('shape');
    await a.s.click(shape.id);
    await a.page.keyboard.press('Shift+ArrowRight');
    await expect.poll(async () => (await b.s.el(shape.id)).x).toBe(shape.x + 10);
    await a.context.close();
    await b.context.close();
  });

  test('SL-COL-02 inserted elements and deletions propagate', async ({ browser }, info) => {
    const room = roomName('slides', info.testId);
    const a = await peer(browser, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, room, 'Bob');
    const n = (await b.s.els()).length;
    await a.s.toolbar('Table (3×3)').click();
    await expect.poll(async () => (await b.s.els()).length).toBe(n + 1);
    const table = (await b.s.els()).find((e) => e.type === 'table')!;
    await a.page.keyboard.press(`${MOD}+z`);
    await expect.poll(async () => (await b.s.els()).some((e) => e.id === table.id)).toBe(false);
    await a.context.close();
    await b.context.close();
  });

  test('SL-COL-03 new slides appear in both filmstrips', async ({ browser }, info) => {
    const room = roomName('slides', info.testId);
    const a = await peer(browser, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, room, 'Bob');
    await a.page.getByRole('button', { name: 'New slide' }).click();
    await expect(a.s.thumbs).toHaveCount(4);
    await expect(b.s.thumbs).toHaveCount(4);
    await a.context.close();
    await b.context.close();
  });

  test('SL-COL-04 text edits propagate', async ({ browser }, info) => {
    const room = roomName('slides', info.testId);
    const a = await peer(browser, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, room, 'Bob');
    const text = await a.s.firstOfType('text');
    const c = await a.s.elementCenter(text.id);
    await a.page.mouse.dblclick(c.x, c.y);
    await expect(a.page.locator('.ProseMirror[contenteditable="true"]')).toBeFocused();
    await a.page.keyboard.press(`${MOD}+a`);
    await a.page.keyboard.type('Edited by Alice');
    await a.page.keyboard.press('Escape');
    await expect(b.s.stage.locator(`[data-element-id="${text.id}"]`)).toContainText(
      'Edited by Alice'
    );
    await a.context.close();
    await b.context.close();
  });

  test('SL-COL-05 a late joiner receives the current deck without duplicated slides', async ({
    browser,
  }, info) => {
    const room = roomName('slides', info.testId);
    const a = await peer(browser, room, 'Alice');
    await waitForRoomSeed(room);
    await a.page.getByRole('button', { name: 'New slide' }).click();
    await expect(a.s.thumbs).toHaveCount(4);
    const witness = await peer(browser, room, 'Bob');
    await expect(witness.s.thumbs).toHaveCount(4);
    const late = await peer(browser, room, 'Carol');
    await expect(late.s.thumbs).toHaveCount(4);
    await a.context.close();
    await witness.context.close();
    await late.context.close();
  });

  test('SL-COL-06 concurrent moves of different elements converge', async ({ browser }, info) => {
    const room = roomName('slides', info.testId);
    const a = await peer(browser, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, room, 'Bob');
    await a.s.gotoSlide(2);
    await b.s.gotoSlide(2);
    const shapes = (await a.s.els()).filter((e) => e.type === 'shape');
    await a.s.click(shapes[0].id);
    await b.s.click(shapes[1].id);
    await Promise.all([
      a.page.keyboard.press('Shift+ArrowLeft'),
      b.page.keyboard.press('Shift+ArrowDown'),
    ]);
    for (const p of [a, b]) {
      await expect.poll(async () => (await p.s.el(shapes[0].id)).x).toBe(shapes[0].x - 10);
      await expect.poll(async () => (await p.s.el(shapes[1].id)).y).toBe(shapes[1].y + 10);
    }
    await a.context.close();
    await b.context.close();
  });

  test('SL-COL-07 the in-process two-pane demo mounts two editors', async ({ page }) => {
    const s = new SlidesPage(page);
    await s.goto();
    await page.getByRole('button', { name: 'Two-pane collab →' }).click();
    await expect(page.locator('[data-slide-stage]')).toHaveCount(2);
  });
});
