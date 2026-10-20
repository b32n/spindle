import { test, expect, type Page } from '@playwright/test';
import { DocsPage, MOD } from '../support/docs';

/** Dispatch a paste event carrying the given clipboard payloads at the editor. */
async function paste(page: Page, data: { text?: string; html?: string }) {
  await page.evaluate((d) => {
    const dt = new DataTransfer();
    if (d.text !== undefined) dt.setData('text/plain', d.text);
    if (d.html !== undefined) dt.setData('text/html', d.html);
    const ev = new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true });
    document.querySelector('.ProseMirror')!.dispatchEvent(ev);
  }, data);
}

test.describe('docs: clipboard', () => {
  let doc: DocsPage;
  test.beforeEach(async ({ page }) => {
    doc = new DocsPage(page);
    await doc.goto();
    await doc.clearAndFocus();
  });

  test('D-CLP-01 pasting plain text inserts it @smoke', async ({ page }) => {
    await paste(page, { text: 'pasted text' });
    await expect(doc.editor).toContainText('pasted text');
  });

  test('D-CLP-02 multi-line plain text becomes paragraphs', async ({ page }) => {
    await paste(page, { text: 'line one\n\nline two' });
    await expect(doc.editor.locator('p').filter({ hasText: 'line one' })).toHaveCount(1);
    await expect(doc.editor.locator('p').filter({ hasText: 'line two' })).toHaveCount(1);
  });

  test('D-CLP-03 pasting HTML keeps bold, links and lists', async ({ page }) => {
    await paste(page, {
      html: '<p>plain <strong>bold</strong> <a href="https://example.org">link</a></p><ul><li>one</li><li>two</li></ul>',
    });
    await expect(doc.editor.locator('strong')).toHaveText('bold');
    await expect(doc.editor.locator('a[href="https://example.org"]')).toHaveText('link');
    await expect(doc.editor.locator('ul > li')).toHaveCount(2);
  });

  test('D-CLP-04 pasted HTML is sanitised (scripts, handlers, javascript: URLs)', async ({
    page,
  }) => {
    await paste(page, {
      html:
        '<p onclick="window.__pwned=1">safe</p><script>window.__pwned=2</script>' +
        '<p><a href="javascript:window.__pwned=3">evil</a></p><img src="x" onerror="window.__pwned=4">',
    });
    await expect(doc.editor).toContainText('safe');
    await expect(doc.editor.locator('script')).toHaveCount(0);
    await expect(doc.editor.locator('[onclick], [onerror]')).toHaveCount(0);
    await expect(doc.editor.locator('a[href^="javascript:"]')).toHaveCount(0);
    expect(
      await page.evaluate(() => (window as unknown as { __pwned?: number }).__pwned)
    ).toBeUndefined();
  });

  test('D-CLP-05 copy and paste within the document (keyboard) @smoke', async ({
    page,
    browserName,
  }) => {
    test.skip(
      browserName === 'webkit',
      'WebKit headless does not expose the OS clipboard to keyboard copy/paste'
    );
    await doc.typeText('copy me');
    await doc.press(`${MOD}+a`);
    await doc.press(`${MOD}+c`);
    await doc.press('End');
    await page.keyboard.type(' ');
    await page.keyboard.press(`${MOD}+v`);
    await expect.poll(() => doc.text()).toMatch(/copy me\s*copy me/);
  });

  test('D-CLP-06 cut removes the selection and paste restores it', async ({
    page,
    browserName,
  }) => {
    test.skip(
      browserName === 'webkit',
      'WebKit headless does not expose the OS clipboard to keyboard copy/paste'
    );
    await doc.typeText('cut me');
    await doc.press(`${MOD}+a`);
    await doc.press(`${MOD}+x`);
    await expect(doc.editor).not.toContainText('cut me');
    await page.keyboard.press(`${MOD}+v`);
    await expect(doc.editor).toContainText('cut me');
  });
});

test.describe('docs: read-only mode', () => {
  test('D-RO-01 read-only blocks typing, formatting and undo @smoke', async ({ page }) => {
    const doc = new DocsPage(page);
    await doc.goto();
    await page.getByRole('button', { name: /Read-only: OFF/ }).click();
    await expect(page.getByRole('button', { name: /Read-only: ON/ })).toBeVisible();
    const before = await doc.text();
    await doc.pages.first().getByText('Features', { exact: true }).first().click();
    await page.keyboard.type('hacked');
    await page.keyboard.press(`${MOD}+a`);
    await page.keyboard.press(`${MOD}+b`);
    await page.keyboard.press('Delete');
    await expect(doc.editor).not.toContainText('hacked');
    expect(await doc.text()).toBe(before);
    await expect(doc.editor.locator('[contenteditable="true"]')).toHaveCount(0);
  });

  test('D-RO-02 the editing toolbar is hidden', async ({ page }) => {
    const doc = new DocsPage(page);
    await doc.goto();
    await page.getByRole('button', { name: /Read-only: OFF/ }).click();
    await expect(doc.editor).toHaveAttribute('contenteditable', 'false');
    // The editing toolbar is not offered at all in read-only mode.
    await expect(doc.toolbar('Bold')).toBeHidden();
    await expect(doc.toolbar('Bullet list')).toBeHidden();
  });
});
