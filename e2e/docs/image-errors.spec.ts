import { test, expect, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { DocsPage, DOCS_URL } from '../support/docs';
import { captureDownload, collectAlerts, collectPageErrors, fixture } from '../support/file-io';

// Image URLs stay on the local demo origin (routed by the test), so no external network is touched.
const MISSING = `${DOCS_URL}/missing/not-here.png`;
const UNREACHABLE = `${DOCS_URL}/unreachable/never.png`;

test.describe('docs: image insertion failures', () => {
  let doc: DocsPage;
  let pageErrors: string[];
  let alerts: string[];
  test.beforeEach(async ({ page }) => {
    pageErrors = collectPageErrors(page);
    alerts = collectAlerts(page);
    await page.route('**/missing/**', (route) =>
      route.fulfill({ status: 404, contentType: 'text/plain', body: 'not found' })
    );
    await page.route('**/unreachable/**', (route) => route.abort('connectionrefused'));
    doc = new DocsPage(page);
    await doc.goto();
    await doc.clearAndFocus();
    await doc.typeText('Keep me');
  });
  test.afterEach(async () => {
    expect(pageErrors, 'uncaught page errors').toEqual([]);
    expect(alerts, 'unexpected alert dialogs').toEqual([]);
    await expect(doc.editor).toContainText('Keep me'); // existing content survives
  });

  async function openUrlTab(page: Page) {
    await doc.toolbar('Insert image').click();
    await page.getByRole('button', { name: 'By URL' }).click();
    return page.getByPlaceholder('https://example.com/image.png');
  }

  for (const [id, label, url] of [
    ['D-IMG-01', '404', MISSING],
    ['D-IMG-02', 'unreachable', UNREACHABLE],
  ] as const) {
    test(`${id} an image URL that is ${label} shows "Could not load image" in the dialog`, async ({
      page,
    }) => {
      const input = await openUrlTab(page);
      await input.fill(url);
      await input.blur();
      await expect(page.getByText('Could not load image. Please check the URL.')).toBeVisible();
      await page.keyboard.press('Escape');
      await expect(input).toBeHidden();
      await expect(doc.editor.locator('img')).toHaveCount(0);
    });
  }

  test('D-IMG-03 after a failed preview the Insert button is disabled, so no broken image enters the document', async ({
    page,
  }) => {
    const input = await openUrlTab(page);
    await input.fill(MISSING);
    await input.blur();
    await expect(page.getByText('Could not load image. Please check the URL.')).toBeVisible();
    await expect(page.getByRole('button', { name: 'Insert', exact: true })).toBeDisabled();
    await expect(doc.editor.locator('img')).toHaveCount(0);
  });

  test('D-IMG-04 uploading a non-image file in the dialog shows "Please select an image file"', async ({
    page,
  }) => {
    await doc.toolbar('Insert image').click();
    await page.locator('input[type="file"][accept="image/*"]').setInputFiles(fixture('notes.txt'));
    await expect(page.getByText('Please select an image file')).toBeVisible();
    await expect(doc.editor.locator('img')).toHaveCount(0);
  });

  test('D-IMG-05 uploading an image over 10 MB is refused ("too large")', async ({ page }) => {
    await doc.toolbar('Insert image').click();
    await page.locator('input[type="file"][accept="image/*"]').setInputFiles({
      name: 'huge.png',
      mimeType: 'image/png',
      buffer: Buffer.alloc(10 * 1024 * 1024 + 1),
    });
    await expect(page.getByText('Image file is too large (max 10MB)')).toBeVisible();
    await expect(doc.editor.locator('img')).toHaveCount(0);
  });

  test('D-IMG-06 uploading a valid PNG inserts it and the DOCX download embeds the same bytes', async ({
    page,
  }) => {
    await doc.toolbar('Insert image').click();
    await page.locator('input[type="file"][accept="image/*"]').setInputFiles(fixture('logo.png'));
    await page.getByRole('button', { name: 'Insert', exact: true }).click();
    await expect(doc.editor.locator('img')).toHaveCount(1);
    await expect(doc.editor.locator('img')).toHaveAttribute('src', /^data:image\/png;base64,/);
    const file = await captureDownload(page, () =>
      page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
    );
    const zip = await JSZip.loadAsync(file.bytes);
    const media = Object.keys(zip.files).filter((n) => /^word\/media\/.+/.test(n));
    expect(media).toHaveLength(1);
    const embedded = await zip.file(media[0])!.async('nodebuffer');
    expect(embedded.equals((await import('node:fs')).readFileSync(fixture('logo.png')))).toBe(true);
  });
});
