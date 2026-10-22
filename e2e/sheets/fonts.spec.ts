import { test, expect } from '@playwright/test';
import { SheetsPage } from '../support/sheets';
import { trackExternalRequests, FONT_LINKS } from '../support/fonts';

async function openFontPicker(sheet: SheetsPage) {
  await sheet.select(0, 0);
  await sheet.page.getByRole('button', { name: /^Arial/ }).click();
  await expect(sheet.page.getByPlaceholder('Search fonts…')).toBeVisible();
}

test.describe('sheets: font source', () => {
  test('S-FNT-02 default: the font picker asks the Spindle mirror, never Google', async ({
    page,
  }) => {
    const external = await trackExternalRequests(page, { cssHosts: ['fonts.b32n.io'] });
    const sheet = new SheetsPage(page);
    await sheet.goto();
    await openFontPicker(sheet);
    await expect
      .poll(() => external.some((u) => u.startsWith('https://fonts.b32n.io/css2?family=')))
      .toBe(true);
    expect(external.filter((u) => !u.startsWith('https://fonts.b32n.io/'))).toEqual([]);
  });

  test("S-FNT-03 fontSource {kind:'local'}: the picker works and makes zero font requests", async ({
    page,
  }) => {
    const external = await trackExternalRequests(page);
    const sheet = new SheetsPage(page);
    await sheet.goto('?fontSource=local');
    await openFontPicker(sheet);
    await page.getByPlaceholder('Search fonts…').fill('Lato');
    await page.getByText('Lato', { exact: true }).first().click();
    await expect.poll(async () => (await sheet.cellStyle(0, 0)).fontFamily).toBe('Lato');
    await expect(page.locator(FONT_LINKS)).toHaveCount(0);
    expect(external).toEqual([]);
  });

  test("S-FNT-04 fontSource {kind:'css', url}: fonts load from the host's own endpoint", async ({
    page,
  }) => {
    const external = await trackExternalRequests(page, { cssHosts: ['fonts.example.test'] });
    const sheet = new SheetsPage(page);
    await sheet.goto(`?fontSource=${encodeURIComponent('css:https://fonts.example.test/css2')}`);
    await openFontPicker(sheet);
    await expect
      .poll(() => external.some((u) => u.startsWith('https://fonts.example.test/css2?family=')))
      .toBe(true);
    expect(external.filter((u) => !u.startsWith('https://fonts.example.test/'))).toEqual([]);
  });

  test('S-FNT-05 unreachable font host: the picker still works and failed stylesheets are dropped', async ({
    page,
  }) => {
    await trackExternalRequests(page, { fail: true });
    const sheet = new SheetsPage(page);
    await sheet.goto();
    await openFontPicker(sheet);
    await expect(page.locator(FONT_LINKS)).toHaveCount(0);
    await page.getByPlaceholder('Search fonts…').fill('Lato');
    await page.getByText('Lato', { exact: true }).first().click();
    await expect.poll(async () => (await sheet.cellStyle(0, 0)).fontFamily).toBe('Lato');
  });
});
