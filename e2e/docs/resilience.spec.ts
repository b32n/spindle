import { test, expect, type Browser } from '@playwright/test';
import { DocsPage, MOD } from '../support/docs';
import { roomName, waitForRoomSeed } from '../support/collab';
import { NetworkProxy, RelayProcess } from '../support/relay';

const WS = 'ws://127.0.0.1:1234';

async function peer(browser: Browser, wsUrl: string, room: string, user: string) {
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  const page = await context.newPage();
  const doc = new DocsPage(page);
  await doc.goto(`?ws=${encodeURIComponent(wsUrl)}&room=${room}&user=${user}`);
  await expect(page.getByText(/· connected$/)).toBeVisible();
  return { context, page, doc };
}
type Peer = Awaited<ReturnType<typeof peer>>;

const down = (p: Peer) => expect(p.page.getByText(/· (offline|connecting)$/)).toBeVisible();
const up = (p: Peer) => expect(p.page.getByText(/· connected$/)).toBeVisible({ timeout: 15_000 });

/** Text of every top-level block, without remote-cursor labels. */
function blocks(p: Peer): Promise<string[]> {
  return p.doc.editor.evaluate((ed) =>
    [...ed.children].map((el) => {
      const clone = el.cloneNode(true) as HTMLElement;
      clone
        .querySelectorAll('.ProseMirror-yjs-cursor, .ProseMirror-widget')
        .forEach((n) => n.remove());
      return clone.textContent ?? '';
    })
  );
}
const block = async (p: Peer, startsWith: string) =>
  (await blocks(p)).find((t) => t.startsWith(startsWith) || t.includes(startsWith)) ?? null;

async function typeAtEnd(p: Peer, line: string, text: string) {
  await p.doc.clickLine(line);
  await p.doc.press('End');
  await p.doc.typeText(text);
}

test.describe('docs: collaboration resilience', () => {
  test('D-RES-01 a dropped connection reconnects and edits sync both ways without duplication', async ({
    browser,
  }, info) => {
    const proxy = new NetworkProxy();
    await proxy.start();
    const room = roomName('docs-res', info.testId);
    const a = await peer(browser, proxy.url, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, proxy.url, room, 'Bob');
    await typeAtEnd(a, 'Welcome to Spindle Docs', ' one');
    await expect
      .poll(() => block(b, 'Welcome to Spindle Docs'))
      .toBe('Welcome to Spindle Docs one');

    await proxy.stop();
    await down(a);
    await down(b);
    await proxy.start();
    await up(a);
    await up(b);

    await typeAtEnd(a, 'Welcome to Spindle Docs', ' two');
    await expect
      .poll(() => block(b, 'Welcome to Spindle Docs'))
      .toBe('Welcome to Spindle Docs one two');
    await typeAtEnd(b, 'Features', ' bob');
    await expect.poll(() => block(a, 'Features')).toBe('Features bob');
    for (const p of [a, b]) {
      await expect(p.doc.editor.locator('h1')).toHaveCount(1);
    }
    expect(await blocks(a)).toEqual(await blocks(b));
    await a.context.close();
    await b.context.close();
    await proxy.stop();
  });

  test('D-RES-02 edits made while offline on both sides merge on reconnect, nothing is lost', async ({
    browser,
  }, info) => {
    const proxy = new NetworkProxy();
    await proxy.start();
    const room = roomName('docs-res', info.testId);
    const a = await peer(browser, proxy.url, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, proxy.url, room, 'Bob');
    const before = await blocks(a);

    await proxy.stop();
    await down(a);
    await down(b);
    await typeAtEnd(a, 'Welcome to Spindle Docs', ' (alice offline)');
    await typeAtEnd(a, 'Features', ' (alice 2)');
    await typeAtEnd(b, 'Page Configuration', ' (bob offline)');
    // Same paragraph, different positions: both survive.
    await b.doc.clickLine('The editor includes support for:');
    await b.doc.press('End');
    await b.doc.typeText('<B');
    await a.doc.clickLine('The editor includes support for:');
    await a.doc.press('Home');
    await a.doc.typeText('A>');
    expect(await block(b, 'Welcome to Spindle Docs')).toBe('Welcome to Spindle Docs');

    await proxy.start();
    await up(a);
    await up(b);
    for (const p of [a, b]) {
      await expect
        .poll(() => block(p, 'Welcome to Spindle Docs'))
        .toBe('Welcome to Spindle Docs (alice offline)');
      await expect.poll(() => block(p, 'Features')).toBe('Features (alice 2)');
      await expect
        .poll(() => block(p, 'Page Configuration'))
        .toBe('Page Configuration (bob offline)');
      await expect
        .poll(() => block(p, 'A>The editor'))
        .toBe('A>The editor includes support for:<B');
    }
    expect((await blocks(a)).length).toBe(before.length);
    expect(await blocks(a)).toEqual(await blocks(b));
    await a.context.close();
    await b.context.close();
    await proxy.stop();
  });

  test('D-RES-03 clients survive a relay restart and keep syncing with each other', async ({
    browser,
  }, info) => {
    test.slow();
    const relay = new RelayProcess();
    await relay.start();
    const room = roomName('docs-res', info.testId);
    const a = await peer(browser, relay.url, room, 'Alice');
    await waitForRoomSeed(room, 1000, 10_000, relay.url);
    const b = await peer(browser, relay.url, room, 'Bob');
    await typeAtEnd(a, 'Welcome to Spindle Docs', ' pre-crash');
    await expect
      .poll(() => block(b, 'Welcome to Spindle Docs'))
      .toBe('Welcome to Spindle Docs pre-crash');

    await relay.kill();
    await down(a);
    await down(b);
    await typeAtEnd(a, 'Features', ' while-down');
    await relay.start(); // empty room: the old log is gone
    await up(a);
    await up(b);

    await typeAtEnd(b, 'Page Configuration', ' after-restart');
    await expect
      .poll(() => block(a, 'Page Configuration'))
      .toBe('Page Configuration after-restart');
    await expect.poll(() => block(b, 'Features')).toBe('Features while-down');
    for (const p of [a, b]) {
      await expect(p.doc.editor.locator('h1')).toHaveCount(1);
      expect(await block(p, 'Welcome to Spindle Docs')).toBe('Welcome to Spindle Docs pre-crash');
    }
    expect(await blocks(a)).toEqual(await blocks(b));
    await a.context.close();
    await b.context.close();
    await relay.kill();
  });

  test('D-RES-04 three clients editing at once converge', async ({ browser }, info) => {
    const room = roomName('docs-res', info.testId);
    const a = await peer(browser, WS, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, WS, room, 'Bob');
    const c = await peer(browser, WS, room, 'Carol');
    await a.doc.clickLine('Welcome to Spindle Docs');
    await a.doc.press('End');
    await b.doc.clickLine('Features');
    await b.doc.press('End');
    await c.doc.clickLine('Page Configuration');
    await c.doc.press('End');
    await Promise.all([
      a.page.keyboard.type('-from-A'),
      b.page.keyboard.type('-from-B'),
      c.page.keyboard.type('-from-C'),
    ]);
    for (const p of [a, b, c]) {
      await expect
        .poll(() => block(p, 'Welcome to Spindle Docs'))
        .toBe('Welcome to Spindle Docs-from-A');
      await expect.poll(() => block(p, 'Features')).toBe('Features-from-B');
      await expect.poll(() => block(p, 'Page Configuration')).toBe('Page Configuration-from-C');
      await expect(p.doc.editor.locator('h1')).toHaveCount(1);
    }
    expect(await blocks(a)).toEqual(await blocks(b));
    expect(await blocks(b)).toEqual(await blocks(c));
    for (const p of [a, b, c]) await p.context.close();
  });

  // B5: undo is not scoped to the local user under collaboration (https://github.com/b32n/spindle/issues/80).
  test.fixme('D-RES-05 undo under collaboration reverts only your own typing (B5)', async ({
    browser,
  }, info) => {
    const room = roomName('docs-res', info.testId);
    const a = await peer(browser, WS, room, 'Alice');
    await waitForRoomSeed(room);
    const b = await peer(browser, WS, room, 'Bob');
    await typeAtEnd(a, 'Welcome to Spindle Docs', ' alice');
    await typeAtEnd(b, 'Features', ' bob');
    await expect.poll(() => block(a, 'Features')).toBe('Features bob');
    await expect
      .poll(() => block(b, 'Welcome to Spindle Docs'))
      .toBe('Welcome to Spindle Docs alice');

    await a.doc.clickLine('Welcome to Spindle Docs');
    await a.page.keyboard.press(`${MOD}+z`);
    for (const p of [a, b]) {
      await expect.poll(() => block(p, 'Welcome to Spindle Docs')).toBe('Welcome to Spindle Docs');
      expect(await block(p, 'Features')).toBe('Features bob');
    }
    await a.context.close();
    await b.context.close();
  });

  test('D-RES-06 a fresh room opened alone is seeded exactly once (StrictMode attach/detach, #72)', async ({
    browser,
  }, info) => {
    const room = roomName('docs-res', info.testId);
    const a = await peer(browser, WS, room, 'Alice');
    await expect(a.doc.editor.locator('h1')).toHaveCount(1);
    const blocksA = await blocks(a);
    await waitForRoomSeed(room);
    const b = await peer(browser, WS, room, 'Bob');
    await expect(b.doc.editor.locator('h1')).toHaveCount(1);
    expect(await blocks(b)).toEqual(blocksA);
    await a.context.close();
    await b.context.close();
  });
});
