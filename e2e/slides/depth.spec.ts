import { test, expect, type Page } from '@playwright/test';
import { SlidesPage, MOD, type ElSnapshot } from '../support/slides';

// The slides toolbar collapses into a "More tools" menu below ~1500px; the formatting bar and the
// align/distribute buttons are asserted inline, so use a wide window.
test.use({ viewport: { width: 1900, height: 900 } });

interface RichDoc {
  content: Array<{
    attrs: Record<string, unknown>;
    content?: Array<{ marks?: Array<{ type: string; attrs?: Record<string, unknown> }> }>;
  }>;
}
const richText = (page: Page, id: string): Promise<RichDoc> =>
  page.evaluate((i) => (window as any).__spindle.deck.getElement(i).richText, id);
/** Every text run of every paragraph must carry `mark` (optionally with matching attrs). */
async function everyRunHas(page: Page, id: string, mark: string, attrs?: Record<string, unknown>) {
  const doc = await richText(page, id);
  const runs = doc.content.flatMap((p) => p.content ?? []);
  expect(runs.length).toBeGreaterThan(1);
  return runs.every((r) =>
    (r.marks ?? []).some(
      (m) =>
        m.type === mark && Object.entries(attrs ?? {}).every(([k, v]) => (m.attrs ?? {})[k] === v)
    )
  );
}
const paragraphAttr = async (page: Page, id: string, key: string) =>
  (await richText(page, id)).content.map((p) => p.attrs[key]);

test.describe('slides: text formatting, themes and layouts', () => {
  let s: SlidesPage;
  let bodyId: string;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
    await s.gotoSlide(2);
    bodyId = (await s.els()).filter((e) => e.type === 'text')[1].id;
  });

  /** Enter edit mode on the body text and select all of it. */
  async function editBody(page: Page) {
    await expect(page.locator('.ProseMirror[contenteditable="true"]')).toHaveCount(0);
    const editor = page.locator('.ProseMirror[contenteditable="true"]');
    const c = await s.elementCenter(bodyId);
    await page.mouse.dblclick(c.x, c.y);
    await expect(editor).toHaveCount(1);
    // After a toolbar click Firefox can leave the focus on the button: put it back in the editor.
    await editor.focus();
    await expect(editor).toBeFocused();
    await page.keyboard.press(`${MOD}+a`);
  }
  const bar = (page: Page, name: string | RegExp) => page.getByRole('button', { name });

  test('SL-TXT-04 italic, underline and strikethrough from the format bar apply to every run and toggle off', async ({
    page,
  }) => {
    await editBody(page);
    for (const [name, mark] of [
      [/^Italic/, 'italic'],
      [/^Underline/, 'underline'],
      ['Strikethrough', 'strikethrough'],
    ] as const) {
      await bar(page, name).click();
      await page.keyboard.press('Escape');
      await expect.poll(() => everyRunHas(page, bodyId, mark)).toBe(true);
      await editBody(page);
      await bar(page, name).click(); // toggle off
      await page.keyboard.press('Escape');
      await expect.poll(() => everyRunHas(page, bodyId, mark)).toBe(false);
      await editBody(page);
    }
  });

  test('SL-TXT-05 font size and text colour apply to the selected text', async ({ page }) => {
    await editBody(page);
    await page.locator('select[title="Font size"]').selectOption('32');
    await page.locator('input[title="Text color"]').fill('#ff0000');
    await page.keyboard.press('Escape');
    await expect.poll(() => everyRunHas(page, bodyId, 'fontSize', { size: 32 })).toBe(true);
    const doc = await richText(page, bodyId);
    const colorMark = doc.content[0].content![0].marks!.find((m) => /color/i.test(m.type));
    expect(JSON.stringify(colorMark).toLowerCase()).toContain('#ff0000');
    // The rendered text really is red and 32px (at scale 1 the stage font follows the slide scale).
    const run = s.stage
      .locator(`[data-element-id="${bodyId}"]`)
      .getByText('Text, shapes, images and lines')
      .first();
    await expect(run).toHaveCSS('color', 'rgb(255, 0, 0)');
  });

  test('SL-TXT-06 numbered and bulleted lists, alignment and line spacing change the paragraphs', async ({
    page,
  }) => {
    // The rich text is written to the model when the editor closes, so each step is committed
    // (Escape) and read back, then the text is re-opened for the next one.
    const apply = async (name: string) => {
      await bar(page, name).click();
      await page.keyboard.press('Escape');
    };
    await editBody(page);
    await apply('Numbered list');
    await expect
      .poll(() => paragraphAttr(page, bodyId, 'listType'))
      .toEqual(Array(5).fill('number'));
    await editBody(page);
    await apply('Bulleted list');
    await expect
      .poll(() => paragraphAttr(page, bodyId, 'listType'))
      .toEqual(Array(5).fill('bullet'));

    await editBody(page);
    await apply('Align right');
    await expect.poll(() => paragraphAttr(page, bodyId, 'align')).toEqual(Array(5).fill('right'));
    await editBody(page);
    await apply('Align center');
    await expect.poll(() => paragraphAttr(page, bodyId, 'align')).toEqual(Array(5).fill('center'));

    await editBody(page);
    await page.locator('select[title="Line spacing"]').selectOption({ index: 3 });
    await page.keyboard.press('Escape');
    await expect
      .poll(async () => (await paragraphAttr(page, bodyId, 'lineHeight')).every((v) => !!v))
      .toBe(true);
  });

  test('SL-THM-01 choosing another theme re-colours the deck and is undoable', async ({ page }) => {
    const themeName = () => page.evaluate(() => (window as any).__spindle.deck.getTheme().name);
    const accent = () =>
      page.evaluate(() => (window as any).__spindle.deck.getTheme().colors.accent1);
    const first = await themeName();
    const firstAccent = await accent();
    const other = first === 'Coral' ? 'Forest' : 'Coral';
    await page.getByRole('button', { name: first, exact: true }).click();
    await page.getByText(other, { exact: true }).click();
    await expect.poll(themeName).toBe(other);
    expect(await accent()).not.toBe(firstAccent);
    await page.keyboard.press('Escape');
    await page.getByRole('button', { name: /^Undo/ }).click();
    await expect.poll(themeName).toBe(first);
  });

  test("SL-LAY-01 New slide with a chosen layout creates that layout's placeholders", async ({
    page,
  }) => {
    const before = (await s.slideIds()).length;
    await page.getByRole('button', { name: 'Slide', exact: true }).click();
    await page.getByText('Title and Content', { exact: true }).click();
    await expect.poll(async () => (await s.slideIds()).length).toBe(before + 1);
    const created = await page.evaluate(() => {
      const d = (window as any).__spindle.deck;
      const id = d.getActiveSlideId();
      const layout = d.getLayout(d.getSlide(id).layoutRef);
      return {
        layoutName: layout?.name,
        placeholders: d.getElementsForSlide(id).filter((e: any) => e.placeholder).length,
        expected: layout?.placeholders.length,
      };
    });
    expect(created.layoutName).toBe('Title and Content');
    expect(created.expected).toBeGreaterThan(1);
    expect(created.placeholders).toBe(created.expected);
  });
});

test.describe('slides: align, distribute and snapping', () => {
  let s: SlidesPage;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
    await s.gotoSlide(3);
  });
  const shapes = async (n: number) => (await s.els()).filter((e) => e.type === 'shape').slice(0, n);
  const select = async (page: Page, els: ElSnapshot[]) => {
    await s.click(els[0].id);
    for (const e of els.slice(1)) await s.click(e.id, ['Shift']);
    await s.expectSelection(els.map((e) => e.id));
  };
  const button = (page: Page, name: string) => page.getByRole('button', { name });

  test('SL-ALN-01 align left / right / top / bottom / centre / middle line the selection up', async ({
    page,
  }) => {
    const three = await shapes(3);
    expect(three.length).toBe(3);
    await select(page, three);
    const cur = async () => Promise.all(three.map((e) => s.el(e.id)));
    const near = (a: number, b: number) => Math.abs(a - b) < 0.6;

    await button(page, 'Align left').click();
    await expect
      .poll(async () => (await cur()).every((e, _, all) => near(e.x, all[0].x)))
      .toBe(true);
    expect(near((await cur())[0].x, Math.min(...three.map((e) => e.x)))).toBe(true);

    await button(page, 'Align right').click();
    await expect
      .poll(async () => (await cur()).every((e, _, all) => near(e.x + e.w, all[0].x + all[0].w)))
      .toBe(true);

    await button(page, 'Align top').click();
    await expect
      .poll(async () => (await cur()).every((e, _, all) => near(e.y, all[0].y)))
      .toBe(true);

    await button(page, 'Align bottom').click();
    await expect
      .poll(async () => (await cur()).every((e, _, all) => near(e.y + e.h, all[0].y + all[0].h)))
      .toBe(true);

    await button(page, 'Align centre').click();
    await expect
      .poll(async () =>
        (await cur()).every((e, _, all) => near(e.x + e.w / 2, all[0].x + all[0].w / 2))
      )
      .toBe(true);
    await button(page, 'Align middle').click();
    await expect
      .poll(async () =>
        (await cur()).every((e, _, all) => near(e.y + e.h / 2, all[0].y + all[0].h / 2))
      )
      .toBe(true);
  });

  test('SL-ALN-02 a single element aligns to the slide edges', async ({ page }) => {
    const [one] = await shapes(1);
    await s.click(one.id);
    await button(page, 'Align left').click();
    await expect.poll(async () => (await s.el(one.id)).x).toBe(0);
    await button(page, 'Align right').click();
    await expect.poll(async () => (await s.el(one.id)).x + (await s.el(one.id)).w).toBe(1280);
    await button(page, 'Align top').click();
    await expect.poll(async () => (await s.el(one.id)).y).toBe(0);
    await button(page, 'Align bottom').click();
    await expect.poll(async () => (await s.el(one.id)).y + (await s.el(one.id)).h).toBe(720);
  });

  test('SL-ALN-03 distribute horizontally and vertically equalises the gaps', async ({ page }) => {
    const three = await shapes(3);
    await select(page, three);
    const cur = async () =>
      (await Promise.all(three.map((e) => s.el(e.id)))).sort((a, b) => a.x - b.x);
    await button(page, 'Distribute horizontally').click();
    await expect
      .poll(async () => {
        const [a, b, c] = await cur();
        return Math.abs(b.x - (a.x + a.w) - (c.x - (b.x + b.w))) < 0.6;
      })
      .toBe(true);
    const curV = async () =>
      (await Promise.all(three.map((e) => s.el(e.id)))).sort((a, b) => a.y - b.y);
    await button(page, 'Distribute vertically').click();
    await expect
      .poll(async () => {
        const [a, b, c] = await curV();
        return Math.abs(b.y - (a.y + a.h) - (c.y - (b.y + b.h))) < 0.6;
      })
      .toBe(true);
    // distribution keeps the outermost elements where they were
    const xs = (await cur()).map((e) => e.x);
    expect(Math.min(...xs)).toBeCloseTo(Math.min(...three.map((e) => e.x)), 0);
  });

  test('SL-SNP-01 dragging near another element snaps to its edge and shows a guide', async ({
    page,
  }) => {
    const [a, b] = await shapes(2);
    const { scale } = await s.toPage(0, 0);
    // Drag `a` so its left edge ends 3 slide-px away from `b`'s left edge: it should snap to b.x.
    const start = await s.elementCenter(a.id);
    const dx = (b.x - a.x + 3) * scale;
    await page.mouse.move(start.x, start.y);
    await page.mouse.down();
    await page.mouse.move(start.x + dx / 2, start.y, { steps: 4 });
    await page.mouse.move(start.x + dx, start.y, { steps: 4 });
    const guideDuringDrag = await s.stage.locator('div[style*="rgb(245, 55, 95)"]').count();
    await page.mouse.up();
    await expect.poll(async () => Math.round((await s.el(a.id)).x)).toBe(Math.round(b.x));
    expect(guideDuringDrag).toBeGreaterThan(0);
    // The guide is only shown while dragging.
    await expect(s.stage.locator('div[style*="rgb(245, 55, 95)"]')).toHaveCount(0);
  });
});
