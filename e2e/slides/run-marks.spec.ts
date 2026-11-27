import { test, expect, type Page } from '@playwright/test';
import { SlidesPage } from '../support/slides';
import { closeToolbarMenu, toolbarButton } from '../support/toolbar';

/** Open the "Text box options" panel (it also holds the run fields): an inline popover, or a modal from "More tools". */
async function openTextOptions(page: Page): Promise<void> {
  await toolbarButton(page, 'Text box options').click();
  await expect(page.getByTestId('run-baseline')).toBeVisible();
}

/** Close the panel (modal Close button, or the popover's backdrop) and any overflow menu left open. */
async function closeTextOptions(page: Page): Promise<void> {
  const dialog = page.getByRole('dialog', { name: 'Text box options' });
  if (await dialog.isVisible()) {
    await dialog.getByRole('button', { name: 'Close' }).click();
    await dialog.waitFor({ state: 'hidden' });
  } else {
    await page.mouse.click(2, 2);
    await expect(page.getByTestId('run-baseline')).toBeHidden();
  }
  await closeToolbarMenu(page);
}

test.describe('slides: run marks and fields', () => {
  let s: SlidesPage;
  let id: string;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
    // A fresh text box so the scenario does not depend on the demo deck's copy. It holds a slide-number field.
    id = await page.evaluate(() => {
      const deck = (window as any).__spindle.deck;
      const slideId = deck.getActiveSlideId();
      const el = deck.addElement(slideId, {
        type: 'text',
        x: 100,
        y: 500,
        w: 600,
        h: 120,
        richText: {
          type: 'doc',
          content: [
            {
              type: 'paragraph',
              content: [
                { type: 'text', text: 'Page ' },
                {
                  type: 'field',
                  attrs: { kind: 'slidenum' },
                  content: [{ type: 'text', text: '‹#›' }],
                },
              ],
            },
          ],
        },
      });
      deck.setSelection({ slideId, elementIds: [el.id] });
      return el.id as string;
    });
    await s.expectSelection([id]);
  });

  const body = (page: Page) => s.stage.locator(`[data-element-id="${id}"] [data-text-body]`);
  const inlines = (page: Page) =>
    page.evaluate(
      (i) => (window as any).__spindle.deck.getElement(i).richText.content[0].content,
      id
    ) as Promise<Array<Record<string, any>>>;

  test('SL-TXT-10 baseline, caps, character spacing and language render and persist', async ({
    page,
  }) => {
    await openTextOptions(page);
    await page.getByTestId('run-baseline').selectOption('super');
    await page.getByTestId('run-caps').selectOption('all');
    await page.getByTestId('run-spacing').fill('3');
    await page.getByTestId('run-lang').fill('fr-FR');
    await page.getByTestId('run-lang').press('Enter');
    await closeTextOptions(page);

    const runs = await inlines(page);
    const types = (runs[0].marks as Array<{ type: string }>).map((m) => m.type).sort();
    expect(types).toEqual(['baseline', 'caps', 'charSpacing', 'lang']);
    expect(runs[0].marks.find((m: any) => m.type === 'baseline').attrs.pct).toBe(30);
    expect(runs[0].marks.find((m: any) => m.type === 'lang').attrs.lang).toBe('fr-FR');

    const run = body(page).locator('span[lang="fr-FR"]').first();
    await expect(run).toHaveCSS('text-transform', 'uppercase');
    await expect(run).toHaveCSS('letter-spacing', '3px');
    // The raised text is a smaller inner span, shifted up.
    const inner = run.locator('span').first();
    await expect(inner).toHaveCSS('position', 'relative');
    expect(parseFloat(await inner.evaluate((e) => getComputedStyle(e).top))).toBeLessThan(0);

    // Back to normal removes the marks again.
    await s.expectSelection([id]);
    await openTextOptions(page);
    await page.getByTestId('run-baseline').selectOption('normal');
    await page.getByTestId('run-caps').selectOption('none');
    await closeTextOptions(page);
    const after = await inlines(page);
    expect((after[0].marks ?? []).map((m: any) => m.type)).toEqual(
      expect.arrayContaining(['charSpacing', 'lang'])
    );
    expect((after[0].marks ?? []).map((m: any) => m.type)).not.toContain('baseline');
  });

  test('SL-TXT-11 a slide-number field shows the number of its slide; a field can be inserted while editing', async ({
    page,
  }) => {
    // The stored text is ‹#›; the rendered value is the slide's position in the deck.
    const number = await page.evaluate(() => {
      const deck = (window as any).__spindle.deck;
      return deck.getSlides().findIndex((sl: any) => sl.id === deck.getActiveSlideId()) + 1;
    });
    await expect(body(page).locator('[data-field="slidenum"]')).toHaveText(String(number));

    // Editing: place the caret at the end and insert a date field from the run options.
    await page.keyboard.press('Enter');
    await expect(page.locator('.ProseMirror[contenteditable="true"]')).toBeFocused();
    await page.keyboard.press('End');
    await openTextOptions(page);
    await page.getByTestId('run-insert-field').selectOption('date');
    await closeTextOptions(page);
    // Leave editing from the editor itself (closing the panel moved focus), which commits the text.
    await page.locator('.ProseMirror[contenteditable="true"]').press('Escape');
    await expect(page.locator('.ProseMirror[contenteditable="true"]')).toHaveCount(0);

    const runs = await inlines(page);
    const dates = runs.filter((r) => r.type === 'field' && r.attrs.kind === 'datetime');
    expect(dates).toHaveLength(1);
    expect(dates[0].attrs.format).toBe('date');
    await expect(body(page).locator('[data-field="datetime"]')).toContainText(
      String(new Date().getFullYear())
    );
    await expect(body(page).locator('[data-field="slidenum"]')).toHaveText(String(number));
  });
});
