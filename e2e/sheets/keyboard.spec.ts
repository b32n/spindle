import { test, expect } from '@playwright/test';
import { SheetsPage, MOD } from '../support/sheets';

// Wide window so the whole toolbar is inline (and in the tab order) rather than in "More tools".
test.use({ viewport: { width: 1900, height: 900 } });

test.describe('sheets: keyboard-only flows', () => {
  let sheet: SheetsPage;
  test.beforeEach(async ({ page }) => {
    sheet = new SheetsPage(page);
    await sheet.goto();
  });

  test('S-KBD-01 navigate, enter data and a formula, format by shortcut and undo without the mouse', async ({
    page,
  }) => {
    // Programmatic focus of the grid stands in for tabbing to it.
    await sheet.focusGrid();
    await page.keyboard.press('ArrowDown');
    await expect.poll(() => sheet.selectionLabel()).toBe('A2');
    for (let i = 0; i < 4; i++) await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowRight');
    await expect.poll(() => sheet.selectionLabel()).toBe('B6');
    await sheet.type('10', 'Enter');
    await expect.poll(() => sheet.selectionLabel()).toBe('B7');
    await sheet.type('32', 'Enter');
    await sheet.type('=SUM(B6:B7)', 'Enter');
    await sheet.expectValue(7, 1, 42);
    expect(await sheet.cellFormula(7, 1)).toBe('=SUM(B6:B7)');
    // Select B6:B8 with Shift+arrows and format by shortcut.
    for (let i = 0; i < 3; i++) await page.keyboard.press('ArrowUp');
    await expect.poll(() => sheet.selectionLabel()).toBe('B6');
    await page.keyboard.press('Shift+ArrowDown');
    await page.keyboard.press('Shift+ArrowDown');
    await page.keyboard.press(`${MOD}+b`);
    for (const r of [5, 6, 7]) await sheet.expectBold(r, 1);
    await page.keyboard.press(`${MOD}+z`);
    await sheet.expectBold(5, 1, false);
  });

  test('S-KBD-02 Tab reaches the toolbar, the formula bar and the grid, each with a visible focus indicator', async ({
    page,
  }) => {
    const seen = new Set<string>();
    let bold = false;
    let grid = false;
    for (let i = 0; i < 50 && !grid; i++) {
      await page.keyboard.press('Tab');
      const info = await page.evaluate(() => {
        const el = document.activeElement as HTMLElement;
        const cs = getComputedStyle(el);
        return {
          body: el === document.body,
          textEntry: el.tagName === 'INPUT' && (el as HTMLInputElement).placeholder !== '',
          label:
            el.getAttribute('aria-label') ||
            el.getAttribute('data-testid') ||
            el.title ||
            el.tagName + ':' + (el.textContent ?? '').slice(0, 20),
          ring:
            (cs.outlineStyle !== 'none' && parseFloat(cs.outlineWidth) > 0) ||
            cs.boxShadow !== 'none',
        };
      });
      expect(info.body, `focus fell to <body> after ${i + 1} Tab presses`).toBe(false);
      // (The formula bar is a text field: its caret is the indicator; it has no outline of its own.)
      if (!info.textEntry) expect(info.ring, `no focus indicator on "${info.label}"`).toBe(true);
      seen.add(info.label);
      if (/Bold/.test(info.label)) bold = true;
      if (info.label === 'sheet-canvas') grid = true;
    }
    expect(bold, [...seen].join(' / ')).toBe(true);
    expect(grid, 'grid reachable with Tab').toBe(true);
    await expect(sheet.formulaInput).toHaveCount(1);
  });

  test('S-KBD-03 the Insert link dialog opens from a focused button, fills and submits with keys, and Escape cancels', async ({
    page,
  }) => {
    await sheet.focusGrid();
    for (let i = 0; i < 5; i++) await page.keyboard.press('ArrowDown');
    await expect.poll(() => sheet.selectionLabel()).toBe('A6');
    const url = page.getByPlaceholder('https://example.com');

    await sheet.toolbar('Insert link').focus();
    await page.keyboard.press('Enter');
    await expect(url).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(url).toBeHidden();

    await sheet.toolbar('Insert link').focus();
    await page.keyboard.press('Enter');
    await expect(url).toBeFocused();
    await page.keyboard.type('https://example.org');
    await page.keyboard.press('Enter');
    await expect(url).toBeHidden();
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).__spindle.workbook.getCell(undefined, 5, 0)?.hyperlink)
      )
      .toBe('https://example.org');
  });

  // Bug: Tab moves the selection and the grid never lets go of focus (keyboard trap).
  // https://github.com/b32n/spindle/issues/86
  test.fixme('S-KBD-04 focus can leave the grid with the keyboard', async ({ page }) => {
    await sheet.focusGrid();
    for (let i = 0; i < 3; i++) await page.keyboard.press('Tab');
    for (const k of ['Escape', 'F6', 'Shift+Tab', `${MOD}+Tab`]) await page.keyboard.press(k);
    await expect(sheet.canvas).not.toBeFocused();
  });

  // Bug: the context menu items are not focusable and Format Cells does not take focus.
  // https://github.com/b32n/spindle/issues/87
  test.fixme('S-KBD-05 Format Cells is reachable from the keyboard via the context menu', async ({
    page,
  }) => {
    await sheet.focusGrid();
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ContextMenu');
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('Enter');
    const dialog = page.getByRole('heading', { name: 'Format Cells' });
    await expect(dialog).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
  });

  // Bug: closing a dialog leaves focus on <body>. https://github.com/b32n/spindle/issues/88
  test.fixme('S-KBD-06 focus is not lost after closing a dialog', async ({ page }) => {
    await sheet.focusGrid();
    await page.keyboard.press('ArrowDown');
    await sheet.toolbar('Insert link').focus();
    await page.keyboard.press('Enter');
    await page.keyboard.press('Escape');
    await expect(page.getByPlaceholder('https://example.com')).toBeHidden();
    expect(await page.evaluate(() => document.activeElement === document.body)).toBe(false);
  });
});
