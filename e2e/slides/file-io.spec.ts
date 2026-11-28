import { test, expect, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { SlidesPage } from '../support/slides';
import {
  asUpload,
  captureDownload,
  collectAlerts,
  collectPageErrors,
  fixture,
  type Downloaded,
} from '../support/file-io';
import { closeToolbarMenu, toolbarButton } from '../support/toolbar';

// Open / Download (.pptx) in the slides toolbar (plan K15). The demo shows the controls by default, like the docs and sheets demos. Both controls may sit in the "More tools" overflow menu, which
// depends on the browser's fonts and widths, so they are clicked through `toolbarButton`.

const OPEN = 'Open file (.pptx)';
const DOWNLOAD = 'Download as Microsoft PowerPoint (.pptx)';
const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

type Opened = { fileName: string; loss: { items: Array<{ id: string; state?: string }> } };
const lastOpened = (page: Page) =>
  page.evaluate(() => (window as unknown as { __lastOpened?: Opened }).__lastOpened);

async function openFile(
  page: Page,
  file: string | { name: string; mimeType: string; buffer: Buffer }
): Promise<void> {
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    toolbarButton(page, OPEN).click(),
  ]);
  await chooser.setFiles(file);
}

const downloadPptx = (page: Page): Promise<Downloaded> =>
  captureDownload(page, async () => {
    await toolbarButton(page, DOWNLOAD).click();
  });

test.describe('slides: file open / download', () => {
  let slides: SlidesPage;
  let pageErrors: string[];
  let alerts: string[];
  const dialog = (page: Page) => page.getByRole('dialog', { name: /Import/ });
  test.beforeEach(async ({ page }) => {
    await page.route(/fonts\.b32n\.io/, (r) => r.abort());
    pageErrors = collectPageErrors(page);
    alerts = collectAlerts(page);
    slides = new SlidesPage(page);
  });
  test.afterEach(() => {
    expect(pageErrors, 'uncaught page errors').toEqual([]);
  });

  test('SL-FIO-01 @smoke opening a generated 6-slide deck shows 6 slides in the filmstrip and the first slide text', async ({
    page,
  }) => {
    await slides.goto();
    await openFile(page, fixture('six-slides.pptx'));
    await expect(slides.thumbs).toHaveCount(6);
    await expect(slides.slide.getByText('Quarterly review')).toBeVisible();
    await expect(slides.slide.getByText('Revenue grew in every region')).toBeVisible();
    await expect.poll(async () => (await lastOpened(page))?.fileName).toBe('six-slides.pptx');
    expect(alerts, 'unexpected alert dialogs').toEqual([]);
  });

  test('SL-FIO-02 @smoke a deck with transitions and charts asks first, listing them under "Kept, not editable"', async ({
    page,
  }) => {
    await slides.goto();
    const before = await slides.thumbs.count();
    await openFile(page, fixture('transitions-charts.pptx'));
    await expect(dialog(page)).toBeVisible();
    await expect(
      dialog(page).getByText('Kept, not editable (saved back unchanged):')
    ).toBeVisible();
    const kept = dialog(page).getByTestId('import-loss-kept-list').locator('li');
    await expect(kept.filter({ hasText: /transition/i })).not.toHaveCount(0);
    await expect(kept.filter({ hasText: /chart/i })).not.toHaveCount(0);
    await expect(slides.thumbs).toHaveCount(before); // not replaced yet
  });

  test('SL-FIO-03 Cancel keeps the open deck; Import anyway replaces it and reports what is kept', async ({
    page,
  }) => {
    await slides.goto();
    const before = await slides.thumbs.count();
    await openFile(page, fixture('transitions-charts.pptx'));
    await dialog(page).getByRole('button', { name: 'Cancel' }).click();
    await expect(dialog(page)).toBeHidden();
    await expect(slides.thumbs).toHaveCount(before);
    expect(await lastOpened(page)).toBeUndefined();

    await openFile(page, fixture('transitions-charts.pptx'));
    await dialog(page).getByRole('button', { name: 'Import anyway' }).click();
    await expect(slides.thumbs).toHaveCount(3);
    await expect
      .poll(async () => (await lastOpened(page))?.fileName)
      .toBe('transitions-charts.pptx');
    expect((await lastOpened(page))?.loss.items.some((i) => i.state === 'kept')).toBe(true);
    expect(alerts, 'unexpected alert dialogs').toEqual([]);
  });

  test('SL-FIO-04 @smoke Download then Open again keeps the slide count and the text', async ({
    page,
  }) => {
    await slides.goto('?importConfirmation=never');
    await openFile(page, fixture('six-slides.pptx'));
    await expect(slides.thumbs).toHaveCount(6);
    await closeToolbarMenu(page);

    const file = await downloadPptx(page);
    expect(file.fileName).toMatch(/\.pptx$/);
    const zip = await JSZip.loadAsync(file.bytes);
    expect(
      Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    ).toHaveLength(6);

    // a fresh page, so the six slides can only come from the downloaded file
    await slides.goto('?importConfirmation=never');
    await openFile(page, asUpload(file, PPTX_MIME));
    await expect(slides.thumbs).toHaveCount(6);
    await expect(slides.slide.getByText('Quarterly review')).toBeVisible();
    expect(alerts, 'unexpected alert dialogs').toEqual([]);
  });

  test('SL-FIO-05 a file that is not a .pptx is refused with a message and leaves the deck as it was', async ({
    page,
  }) => {
    await slides.goto();
    const before = await slides.thumbs.count();
    await openFile(page, fixture('notes.txt'));
    await expect.poll(() => alerts.length).toBe(1);
    expect(alerts[0]).toMatch(/Could not open the file: Unsupported file type: "\.txt"/);
    await expect(slides.thumbs).toHaveCount(before);
  });

  test('SL-FIO-06 the default toolbar has Open and Download, and the Open input is separate from the image input', async ({
    page,
  }) => {
    await slides.goto();
    await expect(page.getByRole('button', { name: OPEN, exact: true })).toHaveCount(1);
    await expect(page.getByRole('button', { name: DOWNLOAD, exact: true })).toHaveCount(1);
    await expect(page.locator('input[type="file"][accept="image/*"]')).toHaveCount(1);
    await expect(page.locator('input[type="file"]')).toHaveCount(2);
  });
});
