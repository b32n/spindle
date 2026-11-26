import { test, expect, type Page } from '@playwright/test';
import path from 'node:path';
import { DocsPage } from '../support/docs';
import { collectPageErrors, openViaChooser } from '../support/file-io';

const OPEN = 'Open file (.docx)';
/** The DOCX fidelity corpus: a document with footnotes and endnotes, one with a link and one in a table. */
const NOTES_DOCX = path.resolve(__dirname, '..', '..', 'fixtures', 'docx', 'notes-toc-fields.docx');

test.describe('docs: footnotes and endnotes from a .docx', () => {
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

  async function openNotes(page: Page) {
    await openViaChooser(page, OPEN, NOTES_DOCX);
    // The file has fields and more that cannot be imported, so the import is confirmed first.
    await page
      .getByRole('dialog', { name: /Import/ })
      .getByRole('button', { name: 'Import anyway' })
      .click();
  }

  test('D-NOT-01 a .docx with footnotes and endnotes shows numbered references and collects the notes at the end @smoke', async ({
    page,
  }) => {
    await openNotes(page);

    // The first footnote reference is a superscript 1 in the text; the endnote reference is a roman numeral.
    const footnote = doc.pages.locator('[data-note-ref="footnote:1"]:not([data-note-target] *)');
    await expect(footnote).toHaveText(/^1/);
    await expect(footnote).toHaveCSS('vertical-align', 'super');
    await expect(
      doc.pages.locator('[data-note-ref="endnote:1"]:not([data-note-target] *)')
    ).toHaveText(/^i/);

    // The notes are collected in an area at the end of the document, which is further down: clicking the reference goes there.
    const note = doc.pages.locator('[data-note-target="footnote:1"]');
    await footnote.click();
    await expect(note).toContainText("Source: the harbour master's annual ledger, table 4.");
    await expect(doc.pages.getByText('Footnotes', { exact: true })).toBeVisible();
    await expect(note).toBeInViewport();
  });

  test('D-NOT-02 clicking a note in the notes area goes back to its reference, and clicking the reference goes to the note again', async ({
    page,
  }) => {
    await openNotes(page);
    const reference = doc.pages.locator('[data-note-ref="footnote:2"]:not([data-note-target] *)');
    await reference.click();
    const note = doc.pages.locator('[data-note-target="footnote:2"]');
    await expect(note).toBeInViewport();
    await expect(reference).not.toBeInViewport();

    // Back to the reference.
    await note.click();
    await expect(reference).toBeInViewport();
    await expect(note).not.toBeInViewport();

    // And forth again.
    await reference.click();
    await expect(note).toBeInViewport();
  });

  test('D-NOT-03 the notes survive the download: the .docx has the notes in its notes parts', async ({
    page,
  }) => {
    await openNotes(page);
    const [download] = await Promise.all([
      page.waitForEvent('download'),
      page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click(),
    ]);
    const { readFile } = await import('node:fs/promises');
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(await readFile((await download.path())!));
    const footnotes = (await zip.file('word/footnotes.xml')?.async('string')) ?? '';
    const endnotes = (await zip.file('word/endnotes.xml')?.async('string')) ?? '';
    expect(footnotes).toContain('Source: the harbour master&apos;s annual ledger, table 4.');
    expect(footnotes).toContain('<w:footnoteRef/>');
    expect(endnotes).toContain('Disputed by the grain exchange');
    const body = (await zip.file('word/document.xml')?.async('string')) ?? '';
    expect(body).toContain('<w:footnoteReference w:id="1"/>');
    expect(body).toContain('<w:endnoteReference w:id="1"/>');
  });
});
