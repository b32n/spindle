import { test, expect } from '@playwright/test';
import { SlidesPage } from './support/slides';
test('probe present esc (temporary)', async ({ page, browserName }) => {
  test.skip(browserName !== 'webkit');
  const s = new SlidesPage(page);
  await s.goto();
  await page.evaluate(() => {
    (window as any).__k = [];
    window.addEventListener('keydown', (e) => (window as any).__k.push(e.key), true);
    document.addEventListener('fullscreenchange', () =>
      (window as any).__k.push('fsc:' + !!document.fullscreenElement)
    );
  });
  const out: any = { enabled: await page.evaluate(() => document.fullscreenEnabled) };
  await s.toolbar('Present').click();
  await page.waitForTimeout(800);
  out.fsBefore = await page.evaluate(() => !!document.fullscreenElement);
  await page.keyboard.press('Escape');
  await page.waitForTimeout(800);
  out.fsAfter = await page.evaluate(() => !!document.fullscreenElement);
  out.btn = await page.getByRole('button', { name: 'Exit (Esc)' }).count();
  out.events = await page.evaluate(() => (window as any).__k);
  expect('PROBE ' + JSON.stringify(out)).toBe('');
});
