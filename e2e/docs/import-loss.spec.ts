import { test, expect, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { DocsPage } from '../support/docs';
import { openViaChooser, fixture } from '../support/file-io';

const OPEN = 'Open file (.docx)';

type Opened = { fileName: string; loss: { items: Array<{ id: string; count: number }> } };
const lastOpened = (page: Page) =>
  page.evaluate(() => (window as unknown as { __lastOpened?: Opened }).__lastOpened);

/** rich-document.docx plus a tracked insertion and a footnote. */
async function lossyDocx(): Promise<{ name: string; mimeType: string; buffer: Buffer }> {
  const zip = await JSZip.loadAsync(
    (await import('node:fs')).readFileSync(fixture('rich-document.docx'))
  );
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file(
    'word/document.xml',
    xml.replace(
      '</w:body>',
      '<w:p><w:ins w:id="99"><w:r><w:t>tracked</w:t></w:r></w:ins></w:p></w:body>'
    )
  );
  zip.file(
    'word/footnotes.xml',
    '<w:footnotes><w:footnote w:id="1"><w:p/></w:footnote></w:footnotes>'
  );
  return {
    name: 'tracked.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: Buffer.from(await zip.generateAsync({ type: 'uint8array' })),
  };
}

test.describe('docs: import loss report and confirmation', () => {
  let doc: DocsPage;
  const dialog = (page: Page) => page.getByRole('dialog', { name: /Import/ });
  test.beforeEach(async ({ page }) => {
    await page.route(/fonts\.b32n\.io/, (r) => r.abort());
    doc = new DocsPage(page);
  });

  test('D-FIO-07 a document with tracked changes and footnotes asks first, listing what would be lost', async ({
    page,
  }) => {
    await doc.goto();
    await openViaChooser(page, OPEN, await lossyDocx());
    await expect(dialog(page)).toBeVisible();
    const items = dialog(page).getByTestId('import-loss-list').locator('li');
    await expect(items).toHaveCount(2);
    await expect(items.filter({ hasText: 'tracked insertion' })).toHaveCount(1);
    await expect(items.filter({ hasText: 'footnote' })).toHaveCount(1);
    await expect(doc.editor.locator('h1')).not.toHaveText('Fixture Title'); // not replaced yet
  });

  test('D-FIO-08 Cancel keeps the open document; Import anyway replaces it and reports the loss', async ({
    page,
  }) => {
    await doc.goto();
    await openViaChooser(page, OPEN, await lossyDocx());
    await dialog(page).getByRole('button', { name: 'Cancel' }).click();
    await expect(dialog(page)).toBeHidden();
    await expect(doc.editor.locator('h1')).not.toHaveText('Fixture Title');
    expect(await lastOpened(page)).toBeUndefined();

    await openViaChooser(page, OPEN, await lossyDocx());
    await dialog(page).getByRole('button', { name: 'Import anyway' }).click();
    await expect(doc.editor.locator('h1')).toHaveText('Fixture Title');
    await expect
      .poll(async () => (await lastOpened(page))?.loss.items.map((i) => i.id).sort())
      .toEqual(['docx.footnotes', 'docx.trackedInsertions']);
  });

  test('D-FIO-09 a document that loses nothing opens without a dialog', async ({ page }) => {
    await doc.goto();
    await openViaChooser(page, OPEN, fixture('rich-document.docx'));
    await expect(doc.editor.locator('h1')).toHaveText('Fixture Title');
    await expect(dialog(page)).toHaveCount(0);
    expect((await lastOpened(page))!.loss.items).toEqual([]);
  });

  test('D-FIO-10 Escape cancels, and a host callback can decide instead of the dialog', async ({
    page,
  }) => {
    await doc.goto();
    await openViaChooser(page, OPEN, await lossyDocx());
    await expect(dialog(page)).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(dialog(page)).toBeHidden();
    await expect(doc.editor.locator('h1')).not.toHaveText('Fixture Title');

    await doc.goto('?importConfirmation=host-no');
    await openViaChooser(page, OPEN, await lossyDocx());
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as unknown as { __hostConfirmCalls?: number[] }).__hostConfirmCalls
        )
      )
      .toEqual([2]);
    await expect(doc.editor.locator('h1')).not.toHaveText('Fixture Title');

    await doc.goto('?importConfirmation=never');
    await openViaChooser(page, OPEN, await lossyDocx());
    await expect(doc.editor.locator('h1')).toHaveText('Fixture Title');
    await expect(dialog(page)).toHaveCount(0);
  });
});
