import { test, expect, type Browser } from '@playwright/test';
import { SlidesPage, MOD } from '../support/slides';
import { roomName, waitForRoomSeed } from '../support/collab';
import { NetworkProxy, RelayProcess, killAllRelays } from '../support/relay';

const WS = 'ws://127.0.0.1:1234';

async function peer(browser: Browser, wsUrl: string, room: string, user: string) {
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  const page = await context.newPage();
  const s = new SlidesPage(page);
  await s.goto(`?ws=${encodeURIComponent(wsUrl)}&room=${room}&user=${user}`);
  await expect(page.getByText(/· connected$/)).toBeVisible();
  return { context, page, s };
}
type Peer = Awaited<ReturnType<typeof peer>>;

const down = (p: Peer) => expect(p.page.getByText(/· (offline|connecting)$/)).toBeVisible();
const up = (p: Peer) => expect(p.page.getByText(/· connected$/)).toBeVisible({ timeout: 15_000 });
const slideCount = async (p: Peer) => (await p.s.slideIds()).length;

/** An element of any slide by id (`SlidesPage.el` only looks at the active slide). */
const anyEl = (p: Peer, id: string) =>
  p.page.evaluate((elId) => {
    const d = (window as any).__spindle.deck;
    return d
      .getSlideIds()
      .flatMap((sid: string) => d.getElementsForSlide(sid))
      .find((e: any) => e.id === elId);
  }, id);

/** Select the element and nudge it by `presses` x 10px with Shift+arrow. */
async function nudge(p: Peer, id: string, key: string, presses = 1) {
  // Clicking an already-selected shape would enter its text-edit mode.
  const selected = await p.s.selection();
  if (selected.length !== 1 || selected[0] !== id) await p.s.click(id);
  await p.s.expectSelection([id]);
  for (let i = 0; i < presses; i++) await p.page.keyboard.press(`Shift+${key}`);
}

test.describe('slides: collaboration resilience', () => {
  test.afterEach(() => killAllRelays());

  test('SL-RES-01 a dropped connection reconnects and edits sync both ways without duplication', async ({
    browser,
  }, info) => {
    const proxy = new NetworkProxy();
    await proxy.start();
    const room = roomName('slides-res', info.testId);
    const a = await peer(browser, proxy.url, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, proxy.url, room, 'Bob');
    await a.s.gotoSlide(2);
    await b.s.gotoSlide(2);
    const [s0, s1] = (await a.s.els()).filter((e) => e.type === 'shape');
    await nudge(a, s0.id, 'ArrowRight');
    await expect.poll(async () => (await b.s.el(s0.id)).x).toBe(s0.x + 10);

    await proxy.stop();
    await down(a);
    await down(b);
    await proxy.start();
    await up(a);
    await up(b);

    await nudge(a, s0.id, 'ArrowRight');
    await expect.poll(async () => (await b.s.el(s0.id)).x).toBe(s0.x + 20);
    await nudge(b, s1.id, 'ArrowDown');
    await expect.poll(async () => (await a.s.el(s1.id)).y).toBe(s1.y + 10);
    for (const p of [a, b]) expect(await slideCount(p)).toBe(3);
    await a.context.close();
    await b.context.close();
    await proxy.stop();
  });

  test('SL-RES-02 edits made while offline on both sides merge on reconnect, nothing is lost', async ({
    browser,
  }, info) => {
    const proxy = new NetworkProxy();
    await proxy.start();
    const room = roomName('slides-res', info.testId);
    const a = await peer(browser, proxy.url, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, proxy.url, room, 'Bob');
    await a.s.gotoSlide(2);
    await b.s.gotoSlide(2);
    const [s0, s1] = (await a.s.els()).filter((e) => e.type === 'shape');
    const before = (await a.s.els()).length;

    await proxy.stop();
    await down(a);
    await down(b);
    await nudge(a, s0.id, 'ArrowLeft', 2);
    await a.page.keyboard.press('Escape');
    await a.s.toolbar('Table (3×3)').click();
    await nudge(b, s1.id, 'ArrowUp', 3);
    await b.page.getByRole('button', { name: 'New slide' }).click();
    await expect.poll(async () => (await a.s.el(s0.id)).x).toBe(s0.x - 20);
    expect((await anyEl(b, s0.id)).x).toBe(s0.x);
    expect(await slideCount(a)).toBe(3);

    await proxy.start();
    await up(a);
    await up(b);
    for (const p of [a, b]) {
      await expect.poll(() => slideCount(p)).toBe(4);
    }
    await a.s.gotoSlide(2);
    await b.s.gotoSlide(2);
    for (const p of [a, b]) {
      await expect.poll(async () => (await p.s.el(s0.id)).x).toBe(s0.x - 20);
      await expect.poll(async () => (await p.s.el(s1.id)).y).toBe(s1.y - 30);
      await expect.poll(async () => (await p.s.els()).length).toBe(before + 1);
    }
    await a.context.close();
    await b.context.close();
    await proxy.stop();
  });

  test('SL-RES-03 clients survive a relay restart and keep syncing with each other', async ({
    browser,
  }, info) => {
    test.slow();
    const relay = new RelayProcess();
    await relay.start();
    const room = roomName('slides-res', info.testId);
    const a = await peer(browser, relay.url, room, 'Alice');
    await waitForRoomSeed(room, 1000, 10_000, relay.url);
    const b = await peer(browser, relay.url, room, 'Bob');
    await a.s.gotoSlide(2);
    await b.s.gotoSlide(2);
    const [s0, s1] = (await a.s.els()).filter((e) => e.type === 'shape');
    await nudge(a, s0.id, 'ArrowRight');
    await expect.poll(async () => (await b.s.el(s0.id)).x).toBe(s0.x + 10);

    await relay.kill();
    await down(a);
    await down(b);
    await nudge(a, s0.id, 'ArrowRight');
    await expect.poll(async () => (await a.s.el(s0.id)).x).toBe(s0.x + 20);
    await relay.start(); // empty room: the old log is gone
    await up(a);
    await up(b);

    await nudge(b, s1.id, 'ArrowDown');
    await expect.poll(async () => (await a.s.el(s1.id)).y).toBe(s1.y + 10);
    await expect.poll(async () => (await b.s.el(s0.id)).x).toBe(s0.x + 20);
    for (const p of [a, b]) expect(await slideCount(p)).toBe(3);
    await a.context.close();
    await b.context.close();
    await relay.kill();
  });

  test('SL-RES-04 three clients editing at once converge', async ({ browser }, info) => {
    const room = roomName('slides-res', info.testId);
    const a = await peer(browser, WS, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, WS, room, 'Bob');
    const c = await peer(browser, WS, room, 'Carol');
    const all = [a, b, c];
    for (const p of all) await p.s.gotoSlide(2);
    const shapes = (await a.s.els()).filter((e) => e.type === 'shape');
    expect(shapes.length).toBeGreaterThanOrEqual(2);
    await a.s.click(shapes[0].id);
    await b.s.click(shapes[1].id);
    await Promise.all([
      a.page.keyboard.press('Shift+ArrowRight'),
      b.page.keyboard.press('Shift+ArrowDown'),
      c.page.getByRole('button', { name: 'New slide' }).click(),
    ]);
    for (const p of all) {
      await expect.poll(async () => (await anyEl(p, shapes[0].id)).x).toBe(shapes[0].x + 10);
      await expect.poll(async () => (await anyEl(p, shapes[1].id)).y).toBe(shapes[1].y + 10);
      await expect.poll(() => slideCount(p)).toBe(4);
    }
    for (const p of all) await p.context.close();
  });

  test('SL-RES-05 undo under collaboration reverts only your own edit', async ({
    browser,
  }, info) => {
    const room = roomName('slides-res', info.testId);
    const a = await peer(browser, WS, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, WS, room, 'Bob');
    await a.s.gotoSlide(2);
    await b.s.gotoSlide(2);
    const [s0, s1] = (await a.s.els()).filter((e) => e.type === 'shape');
    await nudge(a, s0.id, 'ArrowRight');
    await nudge(b, s1.id, 'ArrowDown');
    await expect.poll(async () => (await a.s.el(s1.id)).y).toBe(s1.y + 10);
    await expect.poll(async () => (await b.s.el(s0.id)).x).toBe(s0.x + 10);

    await a.s.click(s0.id);
    await a.page.keyboard.press(`${MOD}+z`);
    for (const p of [a, b]) {
      await expect.poll(async () => (await p.s.el(s0.id)).x).toBe(s0.x);
      expect((await p.s.el(s1.id)).y).toBe(s1.y + 10);
    }
    await a.page.keyboard.press(`${MOD}+Shift+z`);
    for (const p of [a, b]) {
      await expect.poll(async () => (await p.s.el(s0.id)).x).toBe(s0.x + 10);
      expect((await p.s.el(s1.id)).y).toBe(s1.y + 10);
    }
    await a.context.close();
    await b.context.close();
  });

  test('SL-RES-06 a fresh room opened alone is seeded exactly once (StrictMode attach/detach, #72)', async ({
    browser,
  }, info) => {
    const room = roomName('slides-res', info.testId);
    const a = await peer(browser, WS, room, 'Alice');
    expect(await slideCount(a)).toBe(3);
    const elsA = (await a.s.els()).length;
    await waitForRoomSeed(room);
    const b = await peer(browser, WS, room, 'Bob');
    expect(await slideCount(b)).toBe(3);
    expect((await b.s.els()).length).toBe(elsA);
    await a.context.close();
    await b.context.close();
  });
});
