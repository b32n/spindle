import { test, expect, type Page } from '@playwright/test';
import { randomBytes } from 'node:crypto';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { DocsPage } from '../support/docs';
import {
  captureDownload,
  collectAlerts,
  collectPageErrors,
  fixture,
  openViaChooser,
} from '../support/file-io';
import { makePng } from '../support/png';

const OPEN = 'Open file (.docx)';
const junk = Buffer.from(Array.from({ length: 4096 }, (_, i) => (i * 7919 + 13) & 0xff));

const modelJson = (page: Page) =>
  page.evaluate(() => JSON.stringify((window as any).__spindle.document.getSections()));

test.describe('docs: file open error states', () => {
  let doc: DocsPage;
  let pageErrors: string[];
  let alerts: string[];
  let before: string;
  test.beforeEach(async ({ page }) => {
    pageErrors = collectPageErrors(page);
    alerts = collectAlerts(page);
    doc = new DocsPage(page);
    await doc.goto();
    await doc.clearAndFocus();
    await doc.typeText('Precious work');
    before = await modelJson(page);
  });
  test.afterEach(async ({ page }) => {
    expect(pageErrors, 'uncaught page errors').toEqual([]);
    // No data loss: the document is untouched and still editable.
    expect(await modelJson(page)).toBe(before);
    await expect(doc.editor).toHaveText('Precious work');
    await doc.clickText('Precious work');
    await doc.press('End');
    await doc.typeText('!');
    await expect(doc.editor).toHaveText('Precious work!');
  });

  async function expectRejected(
    page: Page,
    file: Parameters<typeof openViaChooser>[2],
    message: RegExp
  ) {
    await openViaChooser(page, OPEN, file);
    await expect.poll(() => alerts.length).toBe(1);
    expect(alerts[0]).toMatch(message);
  }

  test('D-ERR-01 a truncated .docx is rejected with an alert and the document is untouched', async ({
    page,
  }) => {
    const full = readFileSync(fixture('rich-document.docx'));
    await expectRejected(
      page,
      {
        name: 'truncated.docx',
        mimeType: 'application/octet-stream',
        buffer: full.subarray(0, Math.floor(full.length / 2)),
      },
      /^Could not open the file: .+/
    );
  });

  test('D-ERR-02 random bytes named .docx are rejected', async ({ page }) => {
    await expectRejected(
      page,
      { name: 'random.docx', mimeType: 'application/octet-stream', buffer: junk },
      /^Could not open the file: .+/
    );
  });

  test('D-ERR-03 a 0-byte .docx is rejected', async ({ page }) => {
    await expectRejected(
      page,
      { name: 'empty.docx', mimeType: 'application/octet-stream', buffer: Buffer.alloc(0) },
      /^Could not open the file: .+/
    );
  });

  test('D-ERR-04 a valid zip that is not a Word document is rejected', async ({ page }) => {
    const zip = new JSZip();
    zip.file('hello.txt', 'not a docx');
    await expectRejected(
      page,
      {
        name: 'fake.docx',
        mimeType: DOCX_MIME,
        buffer: await zip.generateAsync({ type: 'nodebuffer' }),
      },
      /^Could not open the file: .+/
    );
  });

  test('D-ERR-05 a .txt file is rejected as an unsupported type', async ({ page }) => {
    await expectRejected(
      page,
      fixture('notes.txt'),
      /Unsupported file type: "\.txt"\. Open a \.docx file\./
    );
  });

  test('D-ERR-06 a .png file is rejected as an unsupported type', async ({ page }) => {
    await expectRejected(page, fixture('logo.png'), /Unsupported file type: "\.png"/);
  });

  test('D-ERR-07 an .xlsx handed to the docs Open is rejected', async ({ page }) => {
    await expectRejected(
      page,
      fixture('styled-multisheet.xlsx'),
      /Unsupported file type: "\.xlsx"/
    );
  });

  test('D-ERR-08 a PNG renamed to .docx is rejected', async ({ page }) => {
    await expectRejected(
      page,
      { name: 'photo.docx', mimeType: 'image/png', buffer: readFileSync(fixture('logo.png')) },
      /^Could not open the file: .+/
    );
  });
});

test.describe('docs: large file', () => {
  test('D-ERR-09 a multi-MB .docx (800 paragraphs + a 3 MB image) opens within budget and stays responsive', async ({
    page,
  }, testInfo) => {
    const pageErrors = collectPageErrors(page);
    const alerts = collectAlerts(page);
    const doc = new DocsPage(page);
    await doc.goto();

    // Built at test time from the fixture, so no multi-MB blob is committed.
    const zip = await JSZip.loadAsync(readFileSync(fixture('rich-document.docx')));
    const xml = await zip.file('word/document.xml')!.async('string');
    const at = xml.lastIndexOf('<w:sectPr'); // body-level section properties come last
    const filler = Array.from(
      { length: 800 },
      (_, i) =>
        `<w:p><w:r><w:t>Filler paragraph ${i + 1} lorem ipsum dolor sit amet</w:t></w:r></w:p>`
    ).join('');
    zip.file('word/document.xml', xml.slice(0, at) + filler + xml.slice(at));
    const random = randomBytes(3_000_000); // incompressible, so the PNG (and the zip) really are ~3 MB
    const noise = makePng(1000, 1000, (x, y) => {
      const i = (y * 1000 + x) * 3;
      return [random[i], random[i + 1], random[i + 2]];
    });
    zip.file(
      Object.keys(zip.files).find((n) => /^word\/media\/.+\.png$/.test(n))!,
      noise
    );
    const big = await zip.generateAsync({ type: 'nodebuffer' });
    expect(big.length).toBeGreaterThan(2_500_000);

    const started = Date.now();
    await openViaChooser(page, OPEN, { name: 'big.docx', mimeType: DOCX_MIME, buffer: big });
    await expect(doc.editor.locator('h1')).toHaveText('Fixture Title', { timeout: 15_000 });
    await expect(
      doc.editor.getByText('Filler paragraph 800 lorem ipsum dolor sit amet')
    ).toHaveCount(1);
    const loadMs = Date.now() - started;
    testInfo.annotations.push({ type: 'load-ms', description: String(loadMs) });
    expect(loadMs, 'open of a multi-MB docx').toBeLessThan(15_000);

    // Responsive: edit the heading and see it in the model.
    await doc.clickText('Fixture Title');
    await doc.press('End');
    await doc.typeText(' ok');
    await expect(doc.editor.locator('h1')).toHaveText('Fixture Title ok');
    expect(alerts).toEqual([]);
    expect(pageErrors).toEqual([]);
  });
});

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
