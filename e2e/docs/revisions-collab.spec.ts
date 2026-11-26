import { test, expect, type Browser } from '@playwright/test';
import { DocsPage } from '../support/docs';
import { roomName, waitForRoomSeed } from '../support/collab';
import { NetworkProxy, killAllRelays } from '../support/relay';

const WS = 'ws://127.0.0.1:1234';

/** A collaborator joining a room under a name, with the colour of their cursor. */
async function peer(browser: Browser, room: string, user: string, color: string, wsUrl = WS) {
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  const page = await context.newPage();
  const doc = new DocsPage(page);
  await doc.goto(
    `?ws=${encodeURIComponent(wsUrl)}&room=${room}&user=${user}&color=${encodeURIComponent(color)}`
  );
  await expect(page.getByText(/· connected$/)).toBeVisible();
  return { context, page, doc };
}
type Peer = Awaited<ReturnType<typeof peer>>;

const suggesting = (p: Peer) =>
  p.page
    .getByRole('group', { name: 'Editor mode' })
    .getByRole('button', { name: 'Suggesting' })
    .click();
const insertionsOf = (p: Peer) => p.doc.pages.locator('[data-revision-insertion]');
const deletionsOf = (p: Peer) => p.doc.pages.locator('[data-revision-deletion]');
/** The text of the first paragraph without the remote-cursor name labels peers render into it. */
const firstLine = (p: Peer) =>
  p.doc.editor
    .locator('h1')
    .first()
    .evaluate((el) => {
      const clone = el.cloneNode(true) as HTMLElement;
      clone
        .querySelectorAll('.ProseMirror-yjs-cursor, .ProseMirror-widget')
        .forEach((n) => n.remove());
      return clone.textContent ?? '';
    });

test.describe('docs: tracked changes between collaborators', () => {
  test.afterEach(() => killAllRelays());

  test('D-COL-10 two people suggest, a third accepts one person’s change and rejects the other’s, and all three end with the same document @smoke', async ({
    browser,
  }, info) => {
    const room = roomName('docs', info.testId);
    const alice = await peer(browser, room, 'Alice', '#d93025');
    await waitForRoomSeed(room);
    const bob = await peer(browser, room, 'Bob', '#1a73e8');
    const cy = await peer(browser, room, 'Cy', '#188038');

    // Alice suggests an addition to the heading; Bob, once it has reached him, suggests deleting its first letter.
    await suggesting(alice);
    await alice.doc.clickLine('Welcome to Spindle Docs');
    await alice.doc.press('End');
    await alice.doc.typeText(' (draft)');
    await expect.poll(() => insertionsOf(bob).count()).toBeGreaterThan(0);
    await suggesting(bob);
    await bob.doc.clickLine('Welcome to Spindle Docs');
    await bob.doc.press('Home');
    await bob.doc.press('Delete');
    await expect.poll(() => insertionsOf(cy).count()).toBeGreaterThan(0);
    await expect.poll(() => deletionsOf(cy).count()).toBeGreaterThan(0);

    // Everybody sees both changes, each in the colour of its author's cursor.
    for (const p of [alice, bob, cy]) {
      await expect(insertionsOf(p).first()).toHaveAttribute('data-revision-author', 'Alice');
      await expect(insertionsOf(p).first()).toHaveCSS('color', 'rgb(217, 48, 37)');
      await expect(deletionsOf(p).first()).toHaveAttribute('data-revision-author', 'Bob');
      await expect(deletionsOf(p).first()).toHaveCSS('color', 'rgb(26, 115, 232)');
      await expect(p.page.getByTestId('revision-legend')).toContainText('Alice');
      await expect(p.page.getByTestId('revision-legend')).toContainText('Bob');
    }

    // Cy accepts Alice's change and rejects Bob's.
    await cy.doc.toolbar('Review changes').click();
    const panel = cy.page.getByTestId('review-panel');
    await panel.getByLabel('Author').selectOption('Alice');
    await panel.getByRole('button', { name: 'Accept all by Alice' }).click();
    await panel.getByLabel('Author').selectOption('Bob');
    await panel.getByRole('button', { name: 'Reject all by Bob' }).click();

    // The same document everywhere: the addition is plain text, the deleted letter is back, no changes are left.
    for (const p of [alice, bob, cy]) {
      await expect.poll(() => firstLine(p)).toBe('Welcome to Spindle Docs (draft)');
      await expect(insertionsOf(p)).toHaveCount(0);
      await expect(deletionsOf(p)).toHaveCount(0);
      await expect(p.page.getByTestId('revision-legend')).toHaveCount(0);
    }
    await alice.context.close();
    await bob.context.close();
    await cy.context.close();
  });

  test('D-COL-11 a suggestion typed in Suggesting mode is one undo step for its author only', async ({
    browser,
  }, info) => {
    const room = roomName('docs', info.testId);
    const alice = await peer(browser, room, 'Alice', '#d93025');
    await waitForRoomSeed(room);
    const bob = await peer(browser, room, 'Bob', '#1a73e8');
    await suggesting(alice);
    await alice.doc.clickLine('Welcome to Spindle Docs');
    await alice.doc.press('End');
    // Typing within half a second of earlier edits joins their undo step (the editor's history groups them).
    await alice.page.waitForTimeout(700);
    await alice.doc.typeText(' ABC');
    await expect(insertionsOf(bob).first()).toContainText('ABC');
    await alice.doc.toolbar('Undo').click();
    await expect(insertionsOf(alice)).toHaveCount(0);
    await expect.poll(() => insertionsOf(bob).count()).toBe(0);
    await expect.poll(() => firstLine(bob)).toBe('Welcome to Spindle Docs');
    await alice.context.close();
    await bob.context.close();
  });
  test('D-COL-12 one person splits a paragraph while another marks its mark deleted, both offline: both suggestions survive and are settled one by one', async ({
    browser,
  }, info) => {
    const proxy = new NetworkProxy();
    await proxy.start();
    const room = roomName('docs', info.testId);
    const alice = await peer(browser, room, 'Alice', '#d93025', proxy.url);
    await waitForRoomSeed(room);
    const bob = await peer(browser, room, 'Bob', '#1a73e8', proxy.url);
    await suggesting(alice);
    await suggesting(bob);
    const line = 'Welcome to Spindle Docs';
    const lines = (p: Peer) => p.doc.pages.locator('.line').filter({ hasText: line });

    await proxy.stop();
    await expect(alice.page.getByText(/· (offline|connecting)$/)).toBeVisible();
    await expect(bob.page.getByText(/· (offline|connecting)$/)).toBeVisible();
    // Alice presses Enter at the end of the line; Bob, who knows nothing of it, deletes the paragraph mark there.
    await alice.doc.clickLine(line);
    await alice.doc.press('End');
    await alice.doc.press('Enter');
    await bob.doc.clickLine(line);
    await bob.doc.press('End');
    await bob.page.keyboard.press('Delete');
    await bob.doc.settle();

    await proxy.start();
    await expect(alice.page.getByText(/· connected$/)).toBeVisible({ timeout: 15_000 });
    await expect(bob.page.getByText(/· connected$/)).toBeVisible({ timeout: 15_000 });

    // Both suggestions survive, for both people.
    for (const p of [alice, bob]) {
      await p.doc.toolbar('Review changes').click();
      const panel = p.page.getByTestId('review-panel');
      await expect(
        panel.getByTestId('review-change').filter({ hasText: 'Inserted paragraph break' })
      ).toHaveCount(1, { timeout: 15_000 });
      await expect(
        panel.getByTestId('review-change').filter({ hasText: 'Deleted paragraph break' })
      ).toHaveCount(1);
      await expect(
        panel.getByTestId('review-change').filter({ hasText: 'Inserted paragraph break' })
      ).toHaveAttribute('data-author', 'Alice');
      await expect(
        panel.getByTestId('review-change').filter({ hasText: 'Deleted paragraph break' })
      ).toHaveAttribute('data-author', 'Bob');
    }

    // Alice rejects Bob's deletion alone: it goes away for both, Alice's split stays suggested.
    const panelA = alice.page.getByTestId('review-panel');
    await panelA.getByLabel('Author').selectOption('Bob');
    await panelA.getByRole('button', { name: 'Reject all by Bob' }).click();
    for (const p of [alice, bob]) {
      const panel = p.page.getByTestId('review-panel');
      await expect(panel.getByTestId('review-change')).toHaveCount(1);
      await expect(panel.getByTestId('review-change')).toContainText('Inserted paragraph break');
    }
    // Bob accepts Alice's split: nothing is left to review, and the line is still its own paragraph for both.
    await bob.page
      .getByTestId('review-panel')
      .getByRole('button', { name: 'Accept all', exact: true })
      .click();
    for (const p of [alice, bob]) {
      await expect(p.page.getByTestId('review-panel')).toContainText('No tracked changes.');
      await expect(lines(p)).toHaveCount(1);
      await expect(
        p.doc.pages.locator('[data-revision-insertion], [data-revision-deletion]')
      ).toHaveCount(0);
    }
    await alice.context.close();
    await bob.context.close();
    await proxy.stop();
  });
  test('D-COL-13 a collaborator’s edit lays the page out again in the view the reader chose, not in All markup', async ({
    browser,
  }, info) => {
    const room = roomName('docs', info.testId);
    const alice = await peer(browser, room, 'Alice', '#d93025');
    await waitForRoomSeed(room);
    const bob = await peer(browser, room, 'Bob', '#1a73e8');
    await suggesting(alice);
    await alice.doc.clickLine('Welcome to Spindle Docs');
    await alice.doc.press('End');
    await alice.doc.typeText(' one');
    await expect.poll(() => insertionsOf(bob).count()).toBeGreaterThan(0);

    // Bob looks at the final text: no markup.
    await bob.doc.toolbar('Review changes').click();
    await bob.page
      .getByTestId('review-panel')
      .getByRole('group', { name: 'Show changes as' })
      .getByRole('button', { name: 'Final' })
      .click();
    await expect(insertionsOf(bob)).toHaveCount(0);

    // Alice keeps typing; the page Bob sees is laid out again for her edit and is still the final text.
    await alice.doc.typeText(' two');
    await expect
      .poll(() => bob.doc.pages.first().textContent())
      .toContain('Welcome to Spindle Docs one two');
    await expect(insertionsOf(bob)).toHaveCount(0);
    await expect(insertionsOf(alice).first()).toBeVisible();
    await alice.context.close();
    await bob.context.close();
  });

  test('D-COL-14 an author’s changes are drawn in their cursor colour while they are in the room, and in the page’s own colour once they have left', async ({
    browser,
  }, info) => {
    // Bob's cursor state leaves the room's awareness after y-protocols' outdated timeout (30 s) once his page is closed, so the test
    // waits about that long: it is slow by nature, not by accident.
    test.slow();
    test.setTimeout(100_000);
    const room = roomName('docs', info.testId);
    const alice = await peer(browser, room, 'Alice', '#d93025');
    await waitForRoomSeed(room);
    const bob = await peer(browser, room, 'Bob', '#1a73e8');
    await suggesting(bob);
    await bob.doc.clickLine('Welcome to Spindle Docs');
    await bob.doc.press('End');
    await bob.doc.typeText(' (Bob)');
    await expect.poll(() => insertionsOf(alice).count()).toBeGreaterThan(0);
    const bobsChange = insertionsOf(alice).first();
    await expect(bobsChange).toHaveAttribute('data-revision-author', 'Bob');
    await expect(bobsChange).toHaveCSS('color', 'rgb(26, 115, 232)');
    const swatch = alice.page
      .getByTestId('revision-legend')
      .locator('[data-revision-author="Bob"] span[aria-hidden]');
    await expect(swatch).toHaveCSS('background-color', 'rgb(26, 115, 232)');

    // Bob leaves: his colour was the cursor's, so once the room forgets him (awareness drops a silent peer after 30 seconds)
    // his changes take the colour the page gives them.
    await bob.context.close();
    await expect(bobsChange).not.toHaveCSS('color', 'rgb(26, 115, 232)', { timeout: 50_000 });
    await expect(swatch).not.toHaveCSS('background-color', 'rgb(26, 115, 232)');
    // The legend and the text agree.
    const legendColour = await swatch.evaluate((el) => getComputedStyle(el).backgroundColor);
    await expect(bobsChange).toHaveCSS('color', legendColour);
    await alice.context.close();
  });
});
