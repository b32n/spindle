import { test, expect, type Page } from '@playwright/test';
import JSZip from 'jszip';
import path from 'node:path';
import { DocsPage } from '../support/docs';
import {
  captureDownload,
  collectPageErrors,
  openDocxConfirmingLoss,
  openViaChooser,
} from '../support/file-io';

const OPEN = 'Open file (.docx)';
const NOTES_DOCX = path.resolve(__dirname, '..', '..', 'fixtures', 'docx', 'notes-toc-fields.docx');

async function downloadXml(page: Page, part: string): Promise<string> {
  const file = await captureDownload(page, () =>
    page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
  );
  const zip = await JSZip.loadAsync(file.bytes);
  return (await zip.file(part)?.async('string')) ?? '';
}

test.describe('docs: inserting and editing footnotes and endnotes', () => {
  let doc: DocsPage;
  let pageErrors: string[];
  test.beforeEach(async ({ page }) => {
    await page.route(/fonts\.b32n\.io/, (r) => r.abort());
    pageErrors = collectPageErrors(page);
    doc = new DocsPage(page);
    await doc.goto();
  });
  test.afterEach(() => {
    expect(pageErrors, 'uncaught page errors').toEqual([]);
  });

  /** Insert a note through the Footnote menu and type its text in the editor that opens over it. */
  async function insertNote(page: Page, kind: 'footnote' | 'endnote', noteText: string) {
    await doc.toolbar('Footnote').click();
    await page.getByText(kind === 'footnote' ? 'Insert footnote' : 'Insert endnote').click();
    const editor = page.getByTestId('note-editor').locator('.ProseMirror');
    await expect(editor).toBeFocused();
    await page.keyboard.type(noteText);
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('note-editor')).toHaveCount(0);
    await expect(doc.editor).toBeFocused();
  }

  test('D-NOT-04 Insert -> Footnote puts a numbered reference at the caret and edits the note in the notes area @smoke', async ({
    page,
  }) => {
    await doc.clearAndFocus();
    await doc.typeText('A claim');
    await insertNote(page, 'footnote', 'The proof is in the ledger.');

    // The reference is a superscript 1 after the text, and the note is listed with its number.
    const reference = doc.pages.locator('[data-note-ref="footnote:1"]:not([data-note-target] *)');
    await expect(reference).toHaveText(/^1/);
    const note = doc.pages.locator('[data-note-target="footnote:1"]');
    await expect(note).toContainText('1 The proof is in the ledger.');
    await expect(doc.pages.first()).toContainText('A claim');
    // It is at the bottom of the page the reference is on, under a separator.
    await expect(page.locator('.footnote-separator')).toHaveCount(1);
    const pageBox = await doc.pages.first().boundingBox();
    const noteBox = await note.boundingBox();
    expect(noteBox!.y).toBeGreaterThan(pageBox!.y + pageBox!.height - 200);
  });

  test('D-NOT-05 a second footnote placed earlier renumbers the first, and deleting a reference removes its note and renumbers', async ({
    page,
  }) => {
    await doc.clearAndFocus();
    await doc.typeText('First sentence.');
    await insertNote(page, 'footnote', 'Later note');
    // Now an earlier one, at the start of the text.
    await doc.clickText('First sentence.');
    await doc.press('Home');
    await doc.press('ArrowRight');
    await insertNote(page, 'footnote', 'Earlier note');

    await expect(doc.pages.locator('[data-note-target="footnote:1"]')).toContainText(
      '1 Earlier note'
    );
    await expect(doc.pages.locator('[data-note-target="footnote:2"]')).toContainText(
      '2 Later note'
    );

    // Deleting the first reference (the caret is right after it) removes its note and the other becomes 1.
    await page.keyboard.press('Backspace');
    await expect(doc.pages.locator('[data-note-target="footnote:2"]')).toHaveCount(0);
    await expect(doc.pages.locator('[data-note-target="footnote:1"]')).toContainText(
      '1 Later note'
    );
  });

  test('D-NOT-06 an endnote is numbered in roman numerals and listed under Endnotes; editing a note keeps its formatting', async ({
    page,
  }) => {
    await doc.clearAndFocus();
    await doc.typeText('Text');
    await insertNote(page, 'endnote', 'End');
    await expect(doc.pages.locator('[data-note-target="endnote:1"]')).toContainText('i End');
    await expect(doc.pages.getByText('Endnotes', { exact: true })).toBeVisible();

    // Edit a note that has formatting (the fixture's third footnote has italic, bold and a link) and add to the end.
    await openDocxConfirmingLoss(page, NOTES_DOCX);
    await doc.pages.locator('[data-note-ref="footnote:3"]:not([data-note-target] *)').click();
    const entry = doc.pages.locator('[data-note-target="footnote:3"]');
    await expect(entry).toContainText('A longer note.');
    await entry.getByText('A longer note.').click();
    const editor = page.getByTestId('note-editor').locator('.ProseMirror');
    await expect(editor).toBeFocused();
    await expect(editor).toContainText('A longer note.');
    await page.keyboard.press('Control+End');
    await page.keyboard.type(' Appended.');
    await page.keyboard.press('Escape');

    await expect(entry).toContainText('Appended.');
    // Bold and italic text and the link are still there in the file.
    const xml = await downloadXml(page, 'word/footnotes.xml');
    expect(xml).toContain('Appended.');
    expect(xml).toContain('<w:b/>');
    expect(xml).toContain('<w:i/>');
    expect(xml).toContain('<w:hyperlink');
  });

  test('D-NOT-08 a footnote near the end of a page makes the text above it make room, and later text moves to the next page', async ({
    page,
  }) => {
    await doc.clearAndFocus();
    for (let i = 1; i <= 45; i++) {
      await page.keyboard.type(`Line ${i}`);
      await page.keyboard.press('Enter');
    }
    await doc.settle();
    const pages = page.locator('.page');
    await expect(pages).toHaveCount(2);
    const lastOnFirstPage = async () => {
      const lines = pages.first().locator('.fragment:not(:has([data-note-target])) .line');
      return (await lines.last().innerText()).trim();
    };
    const before = await lastOnFirstPage();

    // A three-line note on a line a few lines above the end of the first page.
    const target = `Line ${Number(before.replace('Line ', '')) - 3}`;
    await doc.clickText(target);
    await doc.press('End');
    await insertNote(page, 'footnote', 'one\ntwo\nthree');
    await expect(pages.first().locator('[data-note-target="footnote:1"]')).toBeVisible();

    // Nothing of the text overlaps the notes' area, and the page now ends earlier.
    const separator = await pages.first().locator('.footnote-separator').boundingBox();
    const lines = pages.first().locator('.fragment:not(:has([data-note-target])) .line');
    for (let i = 0; i < (await lines.count()); i++) {
      const box = await lines.nth(i).boundingBox();
      expect(box!.y + box!.height).toBeLessThanOrEqual(separator!.y + 1);
    }
    expect(await lastOnFirstPage()).not.toBe(before);
    // The last line that was on the first page moved to the second.
    await expect(pages.nth(1)).toContainText(before);
  });

  test('D-NOT-07 notes inserted in the editor survive the download and the file opens with the same notes', async ({
    page,
  }) => {
    await doc.clearAndFocus();
    await doc.typeText('Body text');
    await insertNote(page, 'footnote', 'Saved footnote');
    const file = await captureDownload(page, () =>
      page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
    );
    const zip = await JSZip.loadAsync(file.bytes);
    expect(await zip.file('word/footnotes.xml')!.async('string')).toContain('Saved footnote');
    expect(await zip.file('word/document.xml')!.async('string')).toContain(
      '<w:footnoteReference w:id="1"/>'
    );

    // Reopen the downloaded file through the UI.
    await openViaChooser(page, OPEN, {
      name: file.fileName,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: file.bytes,
    });
    await expect(doc.pages.locator('[data-note-target="footnote:1"]')).toContainText(
      '1 Saved footnote'
    );
    await expect(doc.pages.first()).toContainText('Body text');
  });
});
