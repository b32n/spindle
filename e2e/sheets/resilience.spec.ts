import { test, expect, type Browser } from '@playwright/test';
import { SheetsPage, MOD } from '../support/sheets';
import { roomName, waitForRoomSeed } from '../support/collab';
import { NetworkProxy, RelayProcess } from '../support/relay';

async function peer(browser: Browser, wsUrl: string, room: string, user: string) {
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  const page = await context.newPage();
  const sheet = new SheetsPage(page);
  await sheet.goto(`?ws=${encodeURIComponent(wsUrl)}&room=${room}&user=${user}`);
  await expect(page.getByText(/· connected$/)).toBeVisible();
  return { context, page, sheet };
}
type Peer = Awaited<ReturnType<typeof peer>>;

const down = (p: Peer) => expect(p.page.getByText(/· (offline|connecting)$/)).toBeVisible();
const up = (p: Peer) => expect(p.page.getByText(/· connected$/)).toBeVisible({ timeout: 15_000 });
const sheetCount = (p: Peer) =>
  p.page.evaluate(() => (window as any).__spindle.workbook.getData().sheets.length);

test.describe('sheets: collaboration resilience', () => {
  test('S-RES-01 a dropped connection reconnects and edits sync both ways without duplication', async ({
    browser,
  }, info) => {
    const proxy = new NetworkProxy();
    await proxy.start();
    const room = roomName('sheets-res', info.testId);
    const a = await peer(browser, proxy.url, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, proxy.url, room, 'Bob');
    await a.sheet.enter(6, 0, 'before drop');
    await b.sheet.expectValue(6, 0, 'before drop');

    await proxy.stop();
    await down(a);
    await down(b);
    await proxy.start();
    await up(a);
    await up(b);

    await a.sheet.enter(7, 0, 'after from alice');
    await b.sheet.expectValue(7, 0, 'after from alice');
    await b.sheet.enter(8, 0, 'after from bob');
    await a.sheet.expectValue(8, 0, 'after from bob');
    for (const p of [a, b]) {
      await p.sheet.expectValue(6, 0, 'before drop');
      expect(await sheetCount(p)).toBe(1);
      expect(await p.sheet.cellValue(1, 0)).toBe('Widget');
    }
    await a.context.close();
    await b.context.close();
    await proxy.stop();
  });

  test('S-RES-02 edits made while offline on both sides merge on reconnect, nothing is lost', async ({
    browser,
  }, info) => {
    const proxy = new NetworkProxy();
    await proxy.start();
    const room = roomName('sheets-res', info.testId);
    const a = await peer(browser, proxy.url, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, proxy.url, room, 'Bob');

    await proxy.stop();
    await down(a);
    await down(b);
    // Each side keeps working locally; neither sees the other yet.
    await a.sheet.enter(10, 0, 'alice offline 1');
    await a.sheet.enter(11, 0, 'alice offline 2');
    await a.sheet.enter(14, 3, '=1+2');
    await b.sheet.enter(15, 1, 'bob offline 1');
    await b.sheet.enter(12, 0, 'bob offline 2');
    // Same cell edited on both sides: they must still agree on one of the two values.
    await a.sheet.enter(13, 0, 'alice conflict');
    await b.sheet.enter(13, 0, 'bob conflict');
    await a.sheet.expectValue(11, 0, 'alice offline 2');
    expect(await b.sheet.cellValue(11, 0)).toBeNull();

    await proxy.start();
    await up(a);
    await up(b);
    for (const p of [a, b]) {
      await p.sheet.expectValue(10, 0, 'alice offline 1');
      await p.sheet.expectValue(11, 0, 'alice offline 2');
      await p.sheet.expectValue(14, 3, 3);
      await p.sheet.expectValue(15, 1, 'bob offline 1');
      await p.sheet.expectValue(12, 0, 'bob offline 2');
      expect(await sheetCount(p)).toBe(1);
    }
    await expect
      .poll(async () => {
        const [x, y] = [await a.sheet.cellValue(13, 0), await b.sheet.cellValue(13, 0)];
        return x === y && (x === 'alice conflict' || x === 'bob conflict');
      })
      .toBe(true);
    await a.context.close();
    await b.context.close();
    await proxy.stop();
  });

  // Bug: Rows are created lazily with a random id per peer, so two peers that first touch the same
  // empty row concurrently each create their own row id at that index; the order merge keeps only one
  // of them and the other peer's cell silently vanishes (https://github.com/b32n/spindle/issues/79).
  test.fixme('S-RES-02b concurrent first edits to different cells of the same empty row both survive', async ({
    browser,
  }, info) => {
    const proxy = new NetworkProxy();
    await proxy.start();
    const room = roomName('sheets-res', info.testId);
    const a = await peer(browser, proxy.url, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, proxy.url, room, 'Bob');
    await proxy.stop();
    await down(a);
    await down(b);
    await a.sheet.enter(10, 0, 'alice A11');
    await b.sheet.enter(10, 1, 'bob B11');
    await proxy.start();
    await up(a);
    await up(b);
    for (const p of [a, b]) {
      await p.sheet.expectValue(10, 0, 'alice A11');
      await p.sheet.expectValue(10, 1, 'bob B11');
    }
    await a.context.close();
    await b.context.close();
    await proxy.stop();
  });

  test('S-RES-03 clients survive a relay restart and keep syncing with each other', async ({
    browser,
  }, info) => {
    test.slow();
    const relay = new RelayProcess();
    await relay.start();
    const room = roomName('sheets-res', info.testId);
    const a = await peer(browser, relay.url, room, 'Alice');
    await waitForRoomSeed(room, 1000, 10_000, relay.url);
    const b = await peer(browser, relay.url, room, 'Bob');
    await a.sheet.enter(6, 0, 'pre-crash');
    await b.sheet.expectValue(6, 0, 'pre-crash');

    await relay.kill();
    await down(a);
    await down(b);
    await a.sheet.enter(7, 0, 'alice while down');
    await relay.start(); // empty room: the old log is gone
    await up(a);
    await up(b);

    await b.sheet.enter(8, 0, 'bob after restart');
    await a.sheet.expectValue(8, 0, 'bob after restart');
    await b.sheet.expectValue(7, 0, 'alice while down');
    for (const p of [a, b]) {
      await p.sheet.expectValue(6, 0, 'pre-crash');
      expect(await sheetCount(p)).toBe(1);
    }
    await a.context.close();
    await b.context.close();
    await relay.kill();
  });

  test('S-RES-04 three clients editing at once converge', async ({ browser }, info) => {
    const room = roomName('sheets-res', info.testId);
    const wsUrl = 'ws://127.0.0.1:1234';
    const a = await peer(browser, wsUrl, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, wsUrl, room, 'Bob');
    const c = await peer(browser, wsUrl, room, 'Carol');
    const all = [a, b, c];
    await Promise.all([
      a.sheet.enter(10, 0, 'from A'),
      b.sheet.enter(11, 0, 'from B'),
      c.sheet.enter(12, 0, 'from C'),
    ]);
    await Promise.all([
      a.sheet.enter(20, 5, 'same cell A'),
      b.sheet.enter(20, 5, 'same cell B'),
      c.sheet.enter(20, 5, 'same cell C'),
    ]);
    for (const p of all) {
      await p.sheet.expectValue(10, 0, 'from A');
      await p.sheet.expectValue(11, 0, 'from B');
      await p.sheet.expectValue(12, 0, 'from C');
      expect(await sheetCount(p)).toBe(1);
    }
    await expect
      .poll(async () => {
        const values = await Promise.all(all.map((p) => p.sheet.cellValue(20, 5)));
        return new Set(values).size === 1 && String(values[0]).startsWith('same cell ');
      })
      .toBe(true);
    for (const p of all) await p.context.close();
  });

  test('S-RES-05 undo under collaboration reverts only your own edit', async ({
    browser,
  }, info) => {
    const room = roomName('sheets-res', info.testId);
    const wsUrl = 'ws://127.0.0.1:1234';
    const a = await peer(browser, wsUrl, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, wsUrl, room, 'Bob');
    await a.sheet.enter(10, 0, 'alice edit');
    await b.sheet.enter(11, 0, 'bob edit');
    await a.sheet.expectValue(11, 0, 'bob edit');
    await b.sheet.expectValue(10, 0, 'alice edit');

    await a.sheet.focusGrid();
    await a.page.keyboard.press(`${MOD}+z`);
    for (const p of [a, b]) {
      await p.sheet.expectValue(10, 0, null);
      await p.sheet.expectValue(11, 0, 'bob edit');
    }
    // Redo brings Alice's edit back and still leaves Bob's alone.
    await a.page.keyboard.press(`${MOD}+Shift+z`);
    for (const p of [a, b]) {
      await p.sheet.expectValue(10, 0, 'alice edit');
      await p.sheet.expectValue(11, 0, 'bob edit');
    }
    await a.context.close();
    await b.context.close();
  });

  test('S-RES-06 a fresh room opened alone is seeded exactly once (StrictMode attach/detach, #72)', async ({
    browser,
  }, info) => {
    const room = roomName('sheets-res', info.testId);
    const a = await peer(browser, 'ws://127.0.0.1:1234', room, 'Alice');
    expect(await sheetCount(a)).toBe(1);
    expect(await a.sheet.cellValue(1, 0)).toBe('Widget');
    // A second client joining afterwards sees the same single copy.
    await waitForRoomSeed(room);
    const b = await peer(browser, 'ws://127.0.0.1:1234', room, 'Bob');
    expect(await sheetCount(b)).toBe(1);
    expect(await b.sheet.cellValue(1, 0)).toBe('Widget');
    expect(await b.sheet.cellValue(8, 0)).toBe(await a.sheet.cellValue(8, 0));
    await a.context.close();
    await b.context.close();
  });
});
