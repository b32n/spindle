import { test, expect } from '@playwright/test';
import { DocsPage, MOD } from '../support/docs';
import { collectPageErrors, openViaChooser } from '../support/file-io';
import { bigDocx } from '../support/generated';

const PARAGRAPHS = 1800; // about 55 US-Letter pages with the demo's default styling
const LOAD_BUDGET = 60_000;
const KEY_BUDGET = 20_000;

test.describe('docs: large documents', () => {
  test('D-BIG-01 a ~55-page .docx opens, paginates and typing at the end appears within budget', async ({
    page,
  }, info) => {
    test.setTimeout(180_000);
    const errors = collectPageErrors(page);
    const doc = new DocsPage(page);
    await doc.goto();

    const t0 = Date.now();
    await openViaChooser(page, 'Open file (.docx)', {
      name: 'big.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: await bigDocx(PARAGRAPHS),
    });
    await expect(doc.editor).toContainText(`Paragraph ${PARAGRAPHS}:`, { timeout: LOAD_BUDGET });
    const loadMs = Date.now() - t0;

    // Pagination: the scroller is as tall as >= 45 stacked pages (only a window of pages is
    // mounted in the DOM at a time).
    const pages = await page.evaluate(() => {
      const first = document.querySelector('.page') as HTMLElement;
      let el: HTMLElement | null = first;
      while (
        el &&
        !(el.scrollHeight > el.clientHeight + 50 && getComputedStyle(el).overflowY !== 'visible')
      )
        el = el.parentElement;
      return el ? el.scrollHeight / (first.getBoundingClientRect().height + 24) : 0;
    });
    expect(pages, 'approximate page count').toBeGreaterThan(45);

    // Jump to the end of the document and type there.
    await doc.clickLine('Paragraph 1:');
    await page.keyboard.press(`${MOD}+End`);
    await doc.settle();
    const typed = 'xyz';
    const keyMs: number[] = [];
    for (const ch of typed) {
      const t = Date.now();
      await page.keyboard.type(ch);
      await expect(doc.editor).toContainText(`${PARAGRAPHS}:`, { timeout: KEY_BUDGET });
      await expect
        .poll(() => doc.editor.evaluate((ed) => ed.lastElementChild?.textContent ?? ''), {
          timeout: KEY_BUDGET,
        })
        .toContain(typed.slice(0, keyMs.length + 1));
      keyMs.push(Date.now() - t);
    }
    // The text landed at the end of the last paragraph and nowhere else.
    const last = await doc.editor.evaluate((ed) => ed.lastElementChild?.textContent ?? '');
    expect(last.endsWith(`bank.${typed}`)).toBe(true);
    expect(
      await doc.editor.evaluate((ed, t) => (ed.textContent ?? '').split(t).length - 1, typed)
    ).toBe(1);

    console.log(
      `[D-BIG-01] load ${loadMs} ms, ~${pages.toFixed(1)} pages, per-key ${keyMs.join('/')} ms`
    );
    info.annotations.push({
      type: 'timings',
      description: `load ${loadMs} ms, per-key ${keyMs.join('/')} ms`,
    });
    expect(errors, 'uncaught page errors').toEqual([]);
  });
});
