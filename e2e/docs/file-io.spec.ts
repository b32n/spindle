import { test, expect, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { DocsPage } from '../support/docs';
import {
  asUpload,
  captureDownload,
  collectAlerts,
  collectPageErrors,
  fixture,
  openViaChooser,
  type Downloaded,
} from '../support/file-io';

const OPEN = 'Open file (.docx)';
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const sha1 = (b: Buffer | Uint8Array) => createHash('sha1').update(b).digest('hex');

const downloadDocx = (page: Page): Promise<Downloaded> =>
  captureDownload(page, () =>
    page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
  );

async function unzip(
  file: Downloaded
): Promise<{ zip: JSZip; xml: (name: string) => Promise<string> }> {
  const zip = await JSZip.loadAsync(file.bytes);
  return { zip, xml: async (name) => (await zip.file(name)?.async('string')) ?? '' };
}

test.describe('docs: file open / download round-trips', () => {
  let doc: DocsPage;
  let pageErrors: string[];
  let alerts: string[];
  test.beforeEach(async ({ page }) => {
    pageErrors = collectPageErrors(page);
    alerts = collectAlerts(page);
    doc = new DocsPage(page);
    await doc.goto();
  });
  test.afterEach(() => {
    expect(pageErrors, 'uncaught page errors').toEqual([]);
    expect(alerts, 'unexpected alert dialogs').toEqual([]);
  });

  test('D-FIO-01 opening a LibreOffice .docx shows headings, marks, lists, table, link, image, page break, header and footer', async ({
    page,
  }) => {
    await openViaChooser(page, OPEN, fixture('rich-document.docx'));
    const ed = doc.editor;
    await expect(ed.locator('h1')).toHaveText('Fixture Title');
    await expect(ed.locator('h2')).toHaveText('Lists');
    await expect(ed.locator('h3')).toHaveText('Table, link, image');

    // inline marks
    await expect(ed.locator('strong', { hasText: /^bold$/ })).toHaveCount(1);
    await expect(ed.locator('em', { hasText: /^italic$/ })).toHaveCount(1);
    await expect(ed.locator('u', { hasText: /^underline$/ })).toHaveCount(1);
    await expect(ed.locator('strong em u', { hasText: 'all three' })).toHaveCount(1);

    // lists, including nesting level
    await expect(ed.locator('ul > li')).toHaveText(['Bullet one', 'Nested bullet', 'Bullet two']);
    await expect(ed.locator('ul > li').nth(1)).toHaveAttribute('data-level', '1');
    await expect(ed.locator('ol > li')).toHaveText(['Step one', 'Sub step', 'Step two']);
    await expect(ed.locator('ol > li').nth(1)).toHaveAttribute('data-level', '1');

    // table
    await expect(ed.locator('table tr')).toHaveCount(3);
    await expect(ed.locator('table tr').first().locator('td, th')).toHaveText(['Name', 'Value']);
    await expect(ed.locator('table tr').first().locator('strong')).toHaveCount(2);
    await expect(ed.locator('table tr').nth(2).locator('td')).toHaveText(['Beta', '2']);

    // link, image
    await expect(ed.locator('a')).toHaveAttribute('href', 'https://example.com/spindle');
    await expect(ed.locator('a')).toHaveText('the example site');
    const img = ed.locator('img');
    await expect(img).toHaveCount(1);
    await expect(img).toHaveAttribute('alt', 'Logo');
    await expect(img).toHaveAttribute('src', /^data:image\/png;base64,/);
    await expect(ed.getByText('Text after the page break.')).toBeVisible();

    // header / footer are imported into the model and painted on the page
    const hf = await page.evaluate(() => {
      const s = (window as any).__spindle.document.getSections()[0];
      return {
        header: JSON.stringify(s.header),
        footer: JSON.stringify(s.footer),
        title: (window as any).__spindle.document.getTitle(),
      };
    });
    expect(hf.header).toContain('Fixture header text');
    expect(hf.footer).toContain('Fixture footer');
    expect(hf.footer).toContain('pageNumber');
    expect(hf.title).toBe('Fixture Document');
    await expect(page.locator('.page').first()).toContainText('Fixture header text');
    await expect(page.locator('.page').first()).toContainText('Fixture footer');
  });

  test('D-FIO-02 DOCX -> DOCX: the downloaded file has the same structure, marks, list levels, table, link, image and header/footer @smoke', async ({
    page,
  }) => {
    await openViaChooser(page, OPEN, fixture('rich-document.docx'));
    await expect(doc.editor.locator('h1')).toHaveText('Fixture Title');
    const htmlBefore = await doc.editor.innerHTML();

    const file = await downloadDocx(page);
    expect(file.fileName).toMatch(/\.docx$/);
    const { zip, xml } = await unzip(file);
    const body = await xml('word/document.xml');

    expect(body).toMatch(/w:pStyle w:val="Heading1"[\s\S]*?Fixture Title/);
    expect(body).toMatch(/w:pStyle w:val="Heading2"[\s\S]*?Lists/);
    expect(body).toMatch(/w:pStyle w:val="Heading3"[\s\S]*?Table, link, image/);
    expect(body).toMatch(/<w:r>(?:(?!<\/w:r>)[\s\S])*<w:b\/>(?:(?!<\/w:r>)[\s\S])*bold/);
    expect(body).toMatch(/<w:i\/>(?:(?!<\/w:r>)[\s\S])*italic/);
    expect(body).toMatch(/<w:u [^>]*\/>(?:(?!<\/w:r>)[\s\S])*underline/);
    // lists: both levels present, each list item text exported
    for (const t of [
      'Bullet one',
      'Nested bullet',
      'Bullet two',
      'Step one',
      'Sub step',
      'Step two',
    ])
      expect(body).toContain(t);
    expect(body).toContain('<w:ilvl w:val="1"/>');
    expect(await xml('word/numbering.xml')).toContain('w:abstractNum');
    // table: 3 rows, 6 cells
    expect((body.match(/<w:tr[ >]/g) ?? []).length).toBe(3);
    expect((body.match(/<w:tc>/g) ?? []).length).toBe(6);
    // hyperlink resolves to the right target in the rels part
    const rId = body.match(/<w:hyperlink [^>]*r:id="([^"]+)"/)?.[1];
    expect(rId).toBeTruthy();
    expect(await xml('word/_rels/document.xml.rels')).toMatch(
      new RegExp(`Id="${rId}"[^>]*Target="https://example.com/spindle"`)
    );
    // image: embedded media part is byte-identical to the one in the source file
    expect(body).toContain('<w:drawing>');
    const media = Object.keys(zip.files).filter((n) => /^word\/media\/.+/.test(n));
    expect(media).toHaveLength(1);
    const source = await JSZip.loadAsync(readFileSync(fixture('rich-document.docx')));
    const sourceMedia = await source
      .file(Object.keys(source.files).find((n) => n.startsWith('word/media/'))!)!
      .async('uint8array');
    expect(sha1(await zip.file(media[0])!.async('uint8array'))).toBe(sha1(sourceMedia));
    // header and footer parts carry the text; footer keeps the PAGE field
    const parts = Object.keys(zip.files);
    const header = parts.filter((n) => /word\/header\d*\.xml$/.test(n));
    const footer = parts.filter((n) => /word\/footer\d*\.xml$/.test(n));
    expect(header.length).toBeGreaterThan(0);
    expect(await xml(header[0])).toContain('Fixture header text');
    expect(await xml(footer[0])).toContain('Fixture footer');
    expect(await xml(footer[0])).toMatch(/PAGE/);

    // Re-open the downloaded file through the UI: the editor content is identical.
    await openViaChooser(page, OPEN, asUpload(file, DOCX_MIME));
    await expect.poll(() => doc.editor.innerHTML()).toBe(htmlBefore);
  });

  test('D-FIO-03 text typed after opening is part of the downloaded .docx', async ({ page }) => {
    await openViaChooser(page, OPEN, fixture('rich-document.docx'));
    await expect(doc.editor.locator('h1')).toHaveText('Fixture Title');
    await doc.clickText('Text after the page break.');
    await doc.press('End');
    await doc.typeText(' EDITED-123');
    await expect(doc.editor.getByText('Text after the page break. EDITED-123')).toBeVisible();
    const { xml } = await unzip(await downloadDocx(page));
    expect(await xml('word/document.xml')).toContain('Text after the page break. EDITED-123');
  });

  test('D-FIO-04 a document built in the editor downloads and re-opens identically', async ({
    page,
  }) => {
    await doc.clearAndFocus();
    await doc.typeText('Hello from the editor');
    const file = await downloadDocx(page);
    expect(await (await unzip(file)).xml('word/document.xml')).toContain('Hello from the editor');
    await doc.clickText('Hello from the editor');
    await doc.press('End');
    await doc.typeText(' LOCAL-ONLY');
    await openViaChooser(page, OPEN, asUpload(file, DOCX_MIME));
    await expect(doc.editor).not.toContainText('LOCAL-ONLY');
    await expect(doc.editor).toHaveText('Hello from the editor');
  });

  // Bug (https://github.com/b32n/spindle/issues/76): LibreOffice writes a page break as a trailing <w:br w:type="page"/> run inside the preceding
  // paragraph (here the one holding the image). The importer only maps a break to a page-break block
  // when the paragraph is otherwise empty, so the break is silently dropped.
  test.fixme('D-FIO-05 a page break stored in a content paragraph (LibreOffice style) is imported', async ({
    page,
  }) => {
    await openViaChooser(page, OPEN, fixture('rich-document.docx'));
    await expect(doc.editor.locator('h1')).toHaveText('Fixture Title');
    await expect(doc.editor.locator('.page-break')).toHaveCount(1);
  });

  test('D-FIO-06 a page break inserted in the editor exports as w:br and re-opens as a page break', async ({
    page,
  }) => {
    await doc.clearAndFocus();
    await doc.typeText('Before');
    await doc.press('Enter');
    await doc.toolbar('Page break').click();
    await doc.typeText('After');
    await expect(doc.editor.locator('.page-break')).toHaveCount(1);
    const file = await downloadDocx(page);
    expect(await (await unzip(file)).xml('word/document.xml')).toMatch(/<w:br w:type="page"\/>/);
    await doc.clickText('After');
    await doc.press('End');
    await doc.typeText(' LOCAL-ONLY');
    await openViaChooser(page, OPEN, asUpload(file, DOCX_MIME));
    await expect(doc.editor).not.toContainText('LOCAL-ONLY');
    await expect(doc.editor.locator('.page-break')).toHaveCount(1);
    await expect(doc.editor).toContainText('Before');
    await expect(doc.editor).toContainText('After');
  });
});
