import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { SheetsPage } from '../support/sheets';
import { collectAlerts, collectPageErrors, fixture, openViaChooser } from '../support/file-io';
import { dumpWorkbook } from '../support/sheets-dump';

const OPEN = 'Open file (.xlsx, .csv)';

// Deterministic "random" junk so a failure reproduces.
const junk = Buffer.from(Array.from({ length: 4096 }, (_, i) => (i * 7919 + 13) & 0xff));

test.describe('sheets: file open error states', () => {
  let sheet: SheetsPage;
  let pageErrors: string[];
  let alerts: string[];
  let before: string;
  test.beforeEach(async ({ page }) => {
    pageErrors = collectPageErrors(page);
    alerts = collectAlerts(page);
    sheet = new SheetsPage(page);
    await sheet.goto();
    await sheet.expectValue(0, 0, 'Product');
    before = JSON.stringify(await dumpWorkbook(page));
  });
  test.afterEach(async ({ page }) => {
    expect(pageErrors, 'uncaught page errors').toEqual([]);
    // No data loss: the existing workbook is untouched and still editable.
    expect(JSON.stringify(await dumpWorkbook(page))).toBe(before);
    await sheet.enter(8, 0, 'still works');
    await sheet.expectValue(8, 0, 'still works');
  });

  async function expectRejected(
    page: import('@playwright/test').Page,
    file: Parameters<typeof openViaChooser>[2],
    message: RegExp
  ) {
    await openViaChooser(page, OPEN, file);
    await expect.poll(() => alerts.length).toBe(1);
    expect(alerts[0]).toMatch(message);
  }

  test('S-ERR-01 a truncated .xlsx is rejected with an alert and the workbook is untouched', async ({
    page,
  }) => {
    const full = readFileSync(fixture('styled-multisheet.xlsx'));
    await expectRejected(
      page,
      {
        name: 'truncated.xlsx',
        mimeType: 'application/octet-stream',
        buffer: full.subarray(0, Math.floor(full.length / 2)),
      },
      /^Could not open the file: .+/
    );
  });

  test('S-ERR-02 random bytes named .xlsx are rejected', async ({ page }) => {
    await expectRejected(
      page,
      { name: 'random.xlsx', mimeType: 'application/octet-stream', buffer: junk },
      /^Could not open the file: .+/
    );
  });

  test('S-ERR-03 a 0-byte .xlsx is rejected', async ({ page }) => {
    await expectRejected(
      page,
      { name: 'empty.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.alloc(0) },
      /^Could not open the file: .+/
    );
  });

  test('S-ERR-04 a .txt file is rejected as an unsupported type', async ({ page }) => {
    await expectRejected(
      page,
      fixture('notes.txt'),
      /Unsupported file type: "\.txt"\. Open an \.xlsx or \.csv file\./
    );
  });

  test('S-ERR-05 a .png file is rejected as an unsupported type', async ({ page }) => {
    await expectRejected(page, fixture('logo.png'), /Unsupported file type: "\.png"/);
  });

  test('S-ERR-06 a .docx handed to the sheets Open is rejected', async ({ page }) => {
    await expectRejected(page, fixture('rich-document.docx'), /Unsupported file type: "\.docx"/);
  });

  test('S-ERR-07 a file without an extension is rejected', async ({ page }) => {
    await expectRejected(
      page,
      { name: 'README', mimeType: 'text/plain', buffer: Buffer.from('hello') },
      /Unsupported file type: "\.\?"/
    );
  });

  test('S-ERR-08 a PNG renamed to .xlsx is rejected', async ({ page }) => {
    await expectRejected(
      page,
      { name: 'photo.xlsx', mimeType: 'image/png', buffer: readFileSync(fixture('logo.png')) },
      /^Could not open the file: .+/
    );
  });

  test('S-ERR-10 a valid zip that is not a workbook is rejected and the workbook is untouched', async ({
    page,
  }) => {
    const zip = new JSZip();
    zip.file('hello.txt', 'not an xlsx');
    await expectRejected(
      page,
      {
        name: 'fake.xlsx',
        mimeType: 'application/octet-stream',
        buffer: await zip.generateAsync({ type: 'nodebuffer' }),
      },
      /^Could not open the file: .+/
    );
  });

  test('S-ERR-09 a 0-byte .csv is a valid empty file: it opens as an empty sheet without an alert', async ({
    page,
  }) => {
    // Current behaviour: "Open" replaces the workbook, so the demo data is gone (no unsaved-changes prompt).
    await openViaChooser(page, OPEN, {
      name: 'empty.csv',
      mimeType: 'text/csv',
      buffer: Buffer.alloc(0),
    });
    await expect.poll(async () => (await dumpWorkbook(page))[0].cells.A1).toBeUndefined();
    expect(alerts).toEqual([]);
    before = JSON.stringify(await dumpWorkbook(page)); // the empty workbook is the new baseline for the checks below
  });
});
