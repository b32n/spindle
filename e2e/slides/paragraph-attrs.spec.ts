import { test, expect, type Page } from '@playwright/test';
import { SlidesPage } from '../support/slides';
import { closeToolbarMenu, toolbarButton } from '../support/toolbar';

/** Open the "Text box options" panel (it also holds the paragraph fields): an inline popover, or a modal from "More tools". */
async function openTextOptions(page: Page): Promise<void> {
  await toolbarButton(page, 'Text box options').click();
  await expect(page.getByTestId('para-bullet')).toBeVisible();
}

/** Close the panel (modal Close button, or the popover's backdrop) and any overflow menu left open. */
async function closeTextOptions(page: Page): Promise<void> {
  const dialog = page.getByRole('dialog', { name: 'Text box options' });
  if (await dialog.isVisible()) {
    await dialog.getByRole('button', { name: 'Close' }).click();
    await dialog.waitFor({ state: 'hidden' });
  } else {
    await page.mouse.click(2, 2);
    await expect(page.getByTestId('para-bullet')).toBeHidden();
  }
  await closeToolbarMenu(page);
}

test.describe('slides: paragraph attributes (bullets, spacing, margins)', () => {
  let s: SlidesPage;
  let id: string;
  test.beforeEach(async ({ page }) => {
    s = new SlidesPage(page);
    await s.goto();
    // A fresh text box with three paragraphs so the scenario does not depend on the demo deck's copy.
    id = await page.evaluate(() => {
      const deck = (window as any).__spindle.deck;
      const slideId = deck.getActiveSlideId();
      const para = (text: string) => ({ type: 'paragraph', content: [{ type: 'text', text }] });
      const el = deck.addElement(slideId, {
        type: 'text',
        x: 100,
        y: 500,
        w: 400,
        h: 200,
        richText: { type: 'doc', content: [para('First'), para('Second'), para('Third')] },
      });
      deck.setSelection({ slideId, elementIds: [el.id] });
      return el.id as string;
    });
    await s.expectSelection([id]);
  });

  const body = (page: Page) => s.stage.locator(`[data-element-id="${id}"] [data-text-body]`);
  const paragraphs = (page: Page) =>
    page.evaluate(
      (i) =>
        (window as any).__spindle.deck
          .getElement(i)
          .richText.content.map((p: any) => p.attrs ?? {}),
      id
    ) as Promise<Array<Record<string, any>>>;

  test('SL-TXT-09 custom bullet, numbering, line height, spacing and margins render and persist', async ({
    page,
  }) => {
    await openTextOptions(page);
    await page.getByTestId('para-bullet').selectOption('char');
    await page.getByTestId('para-bullet-char').fill('✓');
    await page.getByTestId('para-margin-left').fill('50');
    await page.getByTestId('para-first-indent').fill('-20');
    await page.getByTestId('para-space-after').fill('9');
    await page.getByTestId('para-lh-mode').selectOption('exact');
    await page.getByTestId('para-lh').fill('30');
    await closeTextOptions(page);

    const attrs = await paragraphs(page);
    expect(attrs).toHaveLength(3);
    for (const a of attrs) {
      expect(a.bullet).toMatchObject({ kind: 'char', char: '✓' });
      expect(a.marginLeft).toBe(50);
      expect(a.firstLineIndent).toBe(-20);
      expect(a.spaceAfter).toBe(9);
      expect(a.lineHeightPx).toBe(30);
    }
    const first = body(page).locator('div').filter({ hasText: 'First' }).last();
    await expect(first).toContainText('✓');
    await expect(first).toHaveCSS('padding-left', '50px');
    await expect(first).toHaveCSS('margin-bottom', '9px');
    await expect(first).toHaveCSS('line-height', '30px');

    // Switch to a numbered list in capital letters starting at C.
    await s.expectSelection([id]);
    await openTextOptions(page);
    await page.getByTestId('para-bullet').selectOption('auto');
    await page.getByTestId('para-bullet-scheme').selectOption('alphaUpper');
    await page.getByTestId('para-bullet-start').fill('3');
    await closeTextOptions(page);
    await expect(body(page)).toContainText('C.');
    await expect(body(page)).toContainText('D.');
    await expect(body(page)).toContainText('E.');
    expect((await paragraphs(page))[0].bullet).toMatchObject({
      kind: 'auto',
      scheme: 'alphaUpper',
      startAt: 3,
    });

    // No bullet removes every marker and keeps the text.
    await openTextOptions(page);
    await page.getByTestId('para-bullet').selectOption('none');
    await closeTextOptions(page);
    await expect(body(page)).not.toContainText('C.');
    await expect(body(page)).toContainText('Second');
  });
});
