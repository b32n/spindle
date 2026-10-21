import { test, expect, type Page } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { SheetsPage } from '../support/sheets';
import { DocsPage } from '../support/docs';
import { SlidesPage } from '../support/slides';

// Accessibility ratchet. Each app is scanned with axe in a few key states. Known violations live in
// baseline.json as `"<app>/<state>/<rule id>": <number of offending nodes>`. A run FAILS when a
// violation is new or has more nodes than the baseline, and also when a baseline entry no longer
// occurs (or occurs fewer times), so the baseline can only shrink. Regenerate it deliberately with
// `npm run a11y:baseline`.

const BASELINE = path.join(__dirname, 'baseline.json');
const UPDATE = !!process.env.A11Y_UPDATE;
type Baseline = Record<string, number>;
const readBaseline = (): Baseline =>
  existsSync(BASELINE) ? (JSON.parse(readFileSync(BASELINE, 'utf8')) as Baseline) : {};

test.describe('a11y (axe)', () => {
  test.skip(
    ({ browserName }) => browserName !== 'chromium',
    'axe results are browser independent; scan once'
  );
  test.use({ viewport: { width: 1900, height: 900 } });
  // The baseline is shared; regenerating it must not race.
  if (UPDATE) test.describe.configure({ mode: 'serial' });

  async function check(page: Page, key: string) {
    const results = await new AxeBuilder({ page })
      .withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'best-practice'])
      .analyze();
    const actual: Baseline = {};
    for (const v of results.violations) actual[`${key}/${v.id}`] = v.nodes.length;

    if (UPDATE) {
      const all = readBaseline();
      for (const k of Object.keys(all)) if (k.startsWith(`${key}/`)) delete all[k];
      Object.assign(all, actual);
      const sorted = Object.fromEntries(Object.entries(all).sort(([a], [b]) => a.localeCompare(b)));
      writeFileSync(BASELINE, JSON.stringify(sorted, null, 2) + '\n');
      return;
    }

    const base = Object.fromEntries(
      Object.entries(readBaseline()).filter(([k]) => k.startsWith(`${key}/`))
    );
    const detail = (id: string) => {
      const v = results.violations.find((x) => `${key}/${x.id}` === id);
      return v
        ? `${v.help} (${v.helpUrl})\n      ` +
            v.nodes
              .slice(0, 5)
              .map((n) => n.target.join(' '))
              .join('\n      ')
        : '';
    };
    const worse = Object.entries(actual)
      .filter(([k, n]) => n > (base[k] ?? 0))
      .map(
        ([k, n]) => `NEW/WORSE  ${k}: ${n} node(s), baseline ${base[k] ?? 0}\n      ${detail(k)}`
      );
    const better = Object.entries(base)
      .filter(([k, n]) => (actual[k] ?? 0) < n)
      .map(([k, n]) => `FIXED      ${k}: baseline ${n}, now ${actual[k] ?? 0}`);
    const problems = [...worse, ...better];
    expect(
      problems,
      problems.length
        ? 'a11y baseline mismatch. New violations must be fixed. If violations were fixed, ' +
            'shrink the baseline: run `npm run a11y:baseline` and commit e2e/a11y/baseline.json.'
        : ''
    ).toEqual([]);
  }

  test('A11Y-SHEETS-01 initial load', async ({ page }) => {
    await new SheetsPage(page).goto();
    await check(page, 'sheets/initial');
  });

  test('A11Y-SHEETS-02 Format Cells dialog', async ({ page }) => {
    const sheet = new SheetsPage(page);
    await sheet.goto();
    const p = await sheet.cellCenter(1, 1);
    await page.mouse.click(p.x, p.y, { button: 'right' });
    await page.getByText('Format Cells...').click();
    await expect(page.getByRole('heading', { name: 'Format Cells' })).toBeVisible();
    await check(page, 'sheets/format-cells-dialog');
  });

  test('A11Y-SHEETS-03 comments panel', async ({ page }) => {
    const sheet = new SheetsPage(page);
    await sheet.goto();
    await sheet.select(1, 1);
    await sheet.toolbar('Comments').click();
    await expect(page.getByPlaceholder(/Add a comment/)).toBeVisible();
    await check(page, 'sheets/comments-panel');
  });

  test('A11Y-DOCS-01 initial load', async ({ page }) => {
    await new DocsPage(page).goto();
    await check(page, 'docs/initial');
  });

  test('A11Y-DOCS-02 Insert link dialog', async ({ page }) => {
    const doc = new DocsPage(page);
    await doc.goto();
    await doc.toolbar('Insert link').click();
    await expect(page.getByPlaceholder('https://example.com')).toBeVisible();
    await check(page, 'docs/insert-link-dialog');
  });

  test('A11Y-DOCS-03 Insert image dialog', async ({ page }) => {
    const doc = new DocsPage(page);
    await doc.goto();
    await doc.toolbar('Insert image').click();
    await expect(page.getByRole('button', { name: 'By URL' })).toBeVisible();
    await check(page, 'docs/insert-image-dialog');
  });

  test('A11Y-SLIDES-01 initial load', async ({ page }) => {
    await new SlidesPage(page).goto();
    await check(page, 'slides/initial');
  });

  test('A11Y-SLIDES-02 comments panel', async ({ page }) => {
    const s = new SlidesPage(page);
    await s.goto();
    await s.toolbar('Comments').click();
    await expect(page.getByText('No comments yet.')).toBeVisible();
    await check(page, 'slides/comments-panel');
  });

  test('A11Y-SLIDES-03 present mode', async ({ page }) => {
    const s = new SlidesPage(page);
    await s.goto();
    await s.toolbar('Present').click();
    await expect(page.getByRole('button', { name: 'Exit (Esc)' })).toBeVisible();
    await check(page, 'slides/present-mode');
  });
});
