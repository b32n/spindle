import { test, expect, type Browser } from '@playwright/test';
import { DocsPage } from '../support/docs';
import { roomName, waitForRoomSeed } from '../support/collab';

const WS = 'ws://127.0.0.1:1234';

/** A collaborator joining a room under a name, with the colour of their cursor. */
async function peer(browser: Browser, room: string, user: string, color: string) {
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  const page = await context.newPage();
  const doc = new DocsPage(page);
  await doc.goto(
    `?ws=${encodeURIComponent(WS)}&room=${room}&user=${user}&color=${encodeURIComponent(color)}`
  );
  await expect(page.getByText(/· connected$/)).toBeVisible();
  return { context, page, doc };
}
type Peer = Awaited<ReturnType<typeof peer>>;

const suggesting = (p: Peer) =>
  p.page
    .getByRole('group', { name: 'Editing mode' })
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
  test('D-COL-10 two people suggest, a third accepts one person’s change and rejects the other’s, and all three end with the same document @smoke', async ({
    browser,
  }, info) => {
    const room = roomName('docs', info.testId);
    const alice = await peer(browser, room, 'Alice', '#d93025');
    await waitForRoomSeed(room);
    const bob = await peer(browser, room, 'Bob', '#1a73e8');
    const cy = await peer(browser, room, 'Cy', '#188038');

    // Alice suggests an addition to the heading; Bob suggests deleting its last word.
    await suggesting(alice);
    await alice.doc.clickLine('Welcome to Spindle Docs');
    await alice.doc.press('End');
    await alice.doc.typeText(' (draft)');
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
});
