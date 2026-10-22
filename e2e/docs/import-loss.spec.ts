import { test, expect, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { DocsPage } from '../support/docs';
import { openViaChooser, fixture } from '../support/file-io';

const OPEN = 'Open file (.docx)';

async function lastOpened(page: Page) {
  return page.evaluate(
    () =>
      (
        window as unknown as {
          __lastOpened?: {
            fileName: string;
            loss: { items: Array<{ id: string; count: number }> };
          };
        }
      ).__lastOpened
  );
}

test.describe('docs: import loss report', () => {
  let doc: DocsPage;
  test.beforeEach(async ({ page }) => {
    await page.route(/fonts\.b32n\.io/, (r) => r.abort());
    doc = new DocsPage(page);
    await doc.goto();
  });

  test('D-FIO-07 opening a document with tracked changes and footnotes reports what was not imported', async ({
    page,
  }) => {
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
    const buffer = Buffer.from(await zip.generateAsync({ type: 'uint8array' }));
    await openViaChooser(page, OPEN, {
      name: 'tracked.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer,
    });
    await expect(doc.editor.locator('h1')).toHaveText('Fixture Title');
    await expect
      .poll(async () => (await lastOpened(page))?.loss.items.map((i) => i.id).sort())
      .toEqual(['docx.footnotes', 'docx.trackedInsertions']);
  });

  test('D-FIO-08 opening a plain exported-style document reports no loss', async ({ page }) => {
    await openViaChooser(page, OPEN, fixture('rich-document.docx'));
    await expect(doc.editor.locator('h1')).toHaveText('Fixture Title');
    await expect.poll(async () => (await lastOpened(page))?.fileName).toBe('rich-document.docx');
    expect((await lastOpened(page))!.loss.items).toEqual([]);
  });
});
