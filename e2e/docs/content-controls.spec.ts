import { test, expect, type Page } from '@playwright/test';
import JSZip from 'jszip';
import path from 'node:path';
import { DocsPage } from '../support/docs';
import { captureDownload, collectPageErrors, openViaChooser } from '../support/file-io';

const OPEN = 'Open file (.docx)';
/** The DOCX fidelity corpus: block and inline content controls of every type, nested, and a few that cannot be kept. */
const CONTROLS_DOCX = path.resolve(
  __dirname,
  '..',
  '..',
  'fixtures',
  'docx',
  'content-controls-forms.docx'
);

test.describe('docs: content controls from a .docx', () => {
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

  async function openControls(page: Page) {
    await openViaChooser(page, OPEN, CONTROLS_DOCX);
    // The file has repeating sections, a data-bound control and legacy form fields, so the import is confirmed first.
    await page
      .getByRole('dialog', { name: /Import/ })
      .getByRole('button', { name: 'Import anyway' })
      .click();
  }

  test('D-CCT-01 a .docx with content controls shows their content, with a title chip for each block control and a tooltip for each inline one @smoke', async ({
    page,
  }) => {
    await openControls(page);

    // The text of controls is on the page, not dropped.
    await expect(doc.pages.getByText('Ada Marlowe', { exact: true })).toBeVisible();
    await expect(doc.pages.getByText('Rich text control with')).toBeVisible();
    await expect(doc.pages.getByText('Innermost paragraph, 5 controls deep.')).toBeVisible();

    // A block control has a chip in the margin with its title; nested ones stack (outermost first).
    await expect(page.locator('.control-chip', { hasText: /^Summary$/ })).toHaveCount(1);
    await expect(page.locator('.control-chip', { hasText: /^Notes$/ })).toHaveCount(1);
    await expect(page.locator('.control-chip', { hasText: /^Level 1$/ })).toHaveCount(1);
    await expect(page.locator('.control-chip', { hasText: /^Level 5$/ })).toHaveCount(1);

    // An inline control is tinted and carries its title as a tooltip.
    const name = doc.pages.locator('[data-content-control="Name"]', { hasText: 'Ada Marlowe' });
    await expect(name).toBeVisible();
    await expect(name).toHaveAttribute('title', 'Name');

    // The summary control still shows its placeholder text, in grey.
    await expect(doc.pages.getByText('Click here to enter a summary.')).toBeVisible();
  });

  test('D-CCT-02 clicking a check box control toggles it, and the downloaded .docx carries the new state', async ({
    page,
  }) => {
    await openControls(page);

    const box = doc.pages.locator('[data-control-checkbox]', { hasText: /[☐☒]/ });
    // The file has it checked.
    await expect(box).toHaveText('☒');
    await box.click();
    await expect(box).toHaveText('☐');

    // The download writes the unchecked state (parsed here only to see that the editor's change reached the file).
    const file = await captureDownload(page, () =>
      page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
    );
    const zip = await JSZip.loadAsync(file.bytes);
    const xml = (await zip.file('word/document.xml')?.async('string')) ?? '';
    expect(xml).toMatch(/<w14:checkbox><w14:checked w14:val="0"\/>/);
    expect(xml).not.toMatch(/<w14:checked w14:val="1"\/>/);

    // And back again.
    await box.click();
    await expect(box).toHaveText('☒');
  });

  test('D-CCT-03 typing inside an inline control keeps the text in the control, and the controls survive the download', async ({
    page,
  }) => {
    await openControls(page);
    // Click just past the end of the control's text, which puts the caret at the end of the text inside the control.
    const name = doc.pages.locator('[data-content-control="Name"]', { hasText: 'Ada Marlowe' });
    const box = await name.boundingBox();
    if (!box) throw new Error('the control is not painted');
    await page.mouse.click(box.x + box.width - 1, box.y + box.height / 2);
    await expect(doc.editor).toBeFocused();
    await doc.waitForCaretIn('Ada Marlowe');
    await page.keyboard.type('X');
    await expect(
      doc.pages.locator('[data-content-control="Name"]', { hasText: /Ada MarloweX/ })
    ).toBeVisible();

    const file = await captureDownload(page, () =>
      page.getByRole('button', { name: 'Download as Microsoft Word (.docx)' }).click()
    );
    const zip = await JSZip.loadAsync(file.bytes);
    const xml = (await zip.file('word/document.xml')?.async('string')) ?? '';
    // The typed text sits in the control's content, and the other controls are still there.
    expect(xml).toMatch(
      /<w:tag w:val="name"\/>[\s\S]*?<w:sdtContent>[\s\S]*?Ada MarloweX[\s\S]*?<\/w:sdtContent>/
    );
    expect(xml).toContain('<w:tag w:val="berth"/>');
    expect(xml).toContain('<w:tag w:val="nest5"/>');
  });
});
