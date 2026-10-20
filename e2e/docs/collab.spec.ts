import { test, expect, type Browser } from '@playwright/test';
import { DocsPage, MOD } from '../support/docs';

const WS = 'ws://127.0.0.1:1234';

/** Block text without the remote-cursor widgets (name labels) that peers render into the DOM. */
async function headingText(doc: DocsPage, selector = 'h1'): Promise<string> {
  return doc.editor.locator(selector).first().evaluate((el) => {
    const clone = el.cloneNode(true) as HTMLElement;
    clone.querySelectorAll('.ProseMirror-yjs-cursor, .ProseMirror-widget').forEach((n) => n.remove());
    return clone.textContent ?? '';
  });
}

async function peer(browser: Browser, room: string, user: string) {
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  const page = await context.newPage();
  const doc = new DocsPage(page);
  await doc.goto(`?ws=${encodeURIComponent(WS)}&room=${room}&user=${user}`);
  await expect(page.getByText(/· connected$/)).toBeVisible();
  return { context, page, doc };
}

test.describe('docs: two-client live collaboration', () => {
  test('D-COL-01 text typed by one client appears in the other @smoke', async ({ browser }, info) => {
    const room = `docs-${info.testId}`;
    const a = await peer(browser, room, 'Alice');
    const b = await peer(browser, room, 'Bob');
    await a.doc.clickText('Welcome to Spindle Docs');
    await a.doc.press('End');
    await a.doc.typeText(' from Alice');
    await expect.poll(() => headingText(b.doc)).toBe('Welcome to Spindle Docs from Alice');
    await expect(b.doc.pages.first()).toContainText('from Alice');
    await a.context.close();
    await b.context.close();
  });

  test('D-COL-02 both directions: concurrent edits in different paragraphs converge', async ({ browser }, info) => {
    const room = `docs-${info.testId}`;
    const a = await peer(browser, room, 'Alice');
    const b = await peer(browser, room, 'Bob');
    await a.doc.clickText('Page Configuration');
    await a.doc.press('Home');
    await b.doc.clickText('Features');
    await b.doc.press('End');
    await Promise.all([a.page.keyboard.type('A>'), b.page.keyboard.type('<B')]);
    for (const p of [a, b]) {
      await expect.poll(() => headingText(p.doc, 'h2:nth-of-type(2)')).toBe('A>Page Configuration');
      await expect.poll(() => headingText(p.doc, 'h2')).toBe('Features<B');
    }
    await a.context.close();
    await b.context.close();
  });

  test('D-COL-03 concurrent typing at the same spot converges to identical text', async ({ browser }, info) => {
    const room = `docs-${info.testId}`;
    const a = await peer(browser, room, 'Alice');
    const b = await peer(browser, room, 'Bob');
    await a.doc.clickText('Welcome to Spindle Docs');
    await a.doc.press('End');
    await b.doc.clickText('Welcome to Spindle Docs');
    await b.doc.press('End');
    await Promise.all([a.page.keyboard.type('-aaa'), b.page.keyboard.type('-bbb')]);
    await expect
      .poll(async () => {
        const [x, y] = [await headingText(a.doc), await headingText(b.doc)];
        return x === y && x.includes('aaa') && x.includes('bbb');
      })
      .toBe(true);
    await a.context.close();
    await b.context.close();
  });

  test('D-COL-04 formatting propagates', async ({ browser }, info) => {
    const room = `docs-${info.testId}`;
    const a = await peer(browser, room, 'Alice');
    const b = await peer(browser, room, 'Bob');
    await a.doc.clickText('Welcome to Spindle Docs');
    await a.doc.press(`${MOD}+a`);
    await a.page.keyboard.press(`${MOD}+i`);
    await expect(b.doc.editor.locator('h1 em').first()).toContainText('Welcome to Spindle Docs');
    await a.context.close();
    await b.context.close();
  });

  test('D-COL-05 a late joiner receives the current document', async ({ browser }, info) => {
    const room = `docs-${info.testId}`;
    const a = await peer(browser, room, 'Alice');
    await a.doc.clickText('Welcome to Spindle Docs');
    await a.doc.press('End');
    await a.doc.typeText(' (edited)');
    // A witness confirms the edit reached the relay before the late joiner connects.
    const witness = await peer(browser, room, 'Bob');
    await expect.poll(() => headingText(witness.doc)).toBe('Welcome to Spindle Docs (edited)');
    const late = await peer(browser, room, 'Carol');
    await expect.poll(() => headingText(late.doc)).toBe('Welcome to Spindle Docs (edited)');
    await a.context.close();
    await witness.context.close();
    await late.context.close();
  });

  // Bug: under CPU load (e.g. several workers) a late joiner occasionally ends up
  // with the demo seed content twice (the whole document repeated after the real
  // one) - observed ~2 in 10 runs of this scenario with 4 workers. Joiners are
  // meant to seed nothing (binding.ts "Connect + seed"); suspect the demo's
  // StrictMode double-mount racing two attachCollab() calls on the same room /
  // IndexedDB persistence key, or a connect() that resolves before the replay is
  // applied. Not reproducible serially.
  test.fixme('D-COL-05b a late joiner never ends up with duplicated seed content', async ({ browser }, info) => {
    const room = `docs-${info.testId}`;
    const a = await peer(browser, room, 'Alice');
    await a.doc.clickText('Welcome to Spindle Docs');
    await a.doc.press('End');
    await a.doc.typeText(' (edited)');
    const witness = await peer(browser, room, 'Bob');
    await expect.poll(() => headingText(witness.doc)).toBe('Welcome to Spindle Docs (edited)');
    const late = await peer(browser, room, 'Carol');
    await expect(late.doc.editor.locator('h1')).toHaveCount(1);
    await a.context.close();
    await witness.context.close();
    await late.context.close();
  });

  test('D-COL-06 list and structural edits propagate', async ({ browser }, info) => {
    const room = `docs-${info.testId}`;
    const a = await peer(browser, room, 'Alice');
    const b = await peer(browser, room, 'Bob');
    await a.doc.clickText('Multiple heading levels');
    await a.doc.press('End');
    await a.page.keyboard.press('Enter');
    await a.doc.typeText('Collab item');
    await expect(b.doc.editor.locator('li', { hasText: 'Collab item' })).toHaveCount(1);
    await a.context.close();
    await b.context.close();
  });

  test('D-COL-07 two-pane in-process demo mounts two editors', async ({ page }) => {
    const doc = new DocsPage(page);
    await doc.goto();
    await page.getByRole('button', { name: 'Collab demo' }).click();
    await expect(page.locator('.ProseMirror')).toHaveCount(2);
  });
});
