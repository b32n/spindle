import { expect, type Locator, type Page } from '@playwright/test';

export const DOCS_URL = 'http://127.0.0.1:5176';
export const MOD = process.platform === 'darwin' ? 'Meta' : 'Control';

/** Page object for the paginated docs editor (hidden ProseMirror + rendered pages). */
export class DocsPage {
  /** Off-screen ProseMirror instance that owns the document model & input. */
  readonly editor: Locator;
  readonly pages: Locator;

  constructor(readonly page: Page) {
    this.editor = page.locator('.ProseMirror');
    this.pages = page.locator('.page-content');
  }

  async goto(query = ''): Promise<void> {
    await this.page.goto(`${DOCS_URL}/${query}`);
    await expect(this.pages.first()).toBeVisible();
    await this.page.waitForFunction(() => !!(window as unknown as { __spindle?: unknown }).__spindle);
  }

  toolbar(name: string | RegExp) {
    return this.page.getByRole('button', { name, exact: typeof name === 'string' });
  }

  /** Click the centre of a rendered text run. */
  async clickText(text: string | RegExp, opts: { last?: boolean } = {}) {
    const loc = this.pages.getByText(text, { exact: typeof text === 'string' });
    await (opts.last ? loc.last() : loc.first()).click();
    // The click hands focus to the hidden ProseMirror view asynchronously;
    // keystrokes sent before that land nowhere.
    await expect(this.editor).toBeFocused();
    if (typeof text === 'string') await this.waitForCaretIn(text);
    await this.settle();
  }

  /** Wait until the hidden editor's DOM selection (what native keys act on) sits in text containing `text`. */
  async waitForCaretIn(text: string) {
    await this.page.waitForFunction((t) => {
      const sel = window.getSelection();
      const ed = document.querySelector('.ProseMirror');
      return !!sel && !!ed && !!sel.anchorNode && ed.contains(sel.anchorNode) &&
        (sel.anchorNode.textContent ?? '').includes(t);
    }, text);
  }

  /**
   * Wait for the click's selection to be mirrored into the hidden editor's DOM.
   * ProseMirror applies it ~20ms after focus (a timer in prosemirror-view), and
   * native keys such as Home/End act on that DOM selection, so sending them
   * earlier lands at a stale position. We wait for the DOM selection to sit in
   * the editor, then for two animation frames.
   */
  async settle() {
    await this.page.waitForFunction(() => {
      const sel = window.getSelection();
      const ed = document.querySelector('.ProseMirror');
      return !!sel && !!ed && !!sel.anchorNode && ed.contains(sel.anchorNode);
    });
    await this.page.evaluate(
      () => new Promise<void>((r) => requestAnimationFrame(() => requestAnimationFrame(() => r()))),
    );
  }

  /** Replace the whole document with blank content and focus it. */
  async clearAndFocus() {
    await this.clickText('Welcome to Spindle Docs');
    await this.page.keyboard.press(`${MOD}+a`);
    await this.page.keyboard.press('Delete');
    await expect(this.editor).not.toContainText('Features');
  }

  /** Type and wait until the text has landed in the model (typing races relayout otherwise). */
  async typeText(text: string) {
    await this.page.keyboard.type(text);
    await expect(this.editor).toContainText(text);
    // Caret must have been mirrored to the end of what we typed before further native keys.
    await this.page.waitForFunction((t) => {
      const sel = window.getSelection();
      const n = sel?.anchorNode;
      return !!sel && !!n && sel.isCollapsed && (n.textContent ?? '').endsWith(t) &&
        sel.anchorOffset === (n.textContent ?? '').length;
    }, text);
    await this.settle();
  }

  /**
   * Press a navigation/selection key and let ProseMirror pick up the resulting
   * DOM selection (it observes `selectionchange` asynchronously; typing or
   * formatting sent before that acts on the previous selection).
   */
  async press(key: string) {
    await this.page.keyboard.press(key);
    await this.settle();
  }

  async text(): Promise<string> {
    return this.editor.innerText();
  }
}
