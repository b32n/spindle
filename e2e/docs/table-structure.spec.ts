import { test, expect } from '@playwright/test';
import { DocsPage } from '../support/docs';
import { collectPageErrors, openViaChooser } from '../support/file-io';
import { nestedTableDocx } from '../support/generated';

test.describe('docs: table structure from a .docx', () => {
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

  test('D-TBL-02 a cell with several paragraphs and a nested table is shown, and the nested table can be edited', async ({
    page,
  }) => {
    await openViaChooser(page, 'Open file (.docx)', {
      name: 'nested.docx',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: await nestedTableDocx(),
    });
    // Both paragraphs of the cell and the nested table are drawn on the page; the nested table is inside the first cell.
    const outerCell = doc.pages.locator('table.table-block:not(.nested-table) > tr > td').first();
    await expect(outerCell.locator('> p.cell-paragraph')).toHaveCount(3);
    await expect(outerCell.locator('> p.cell-paragraph').first()).toHaveText('Cell first');
    await expect(outerCell.locator('> p.cell-paragraph').nth(1)).toHaveText('Cell second');
    const nested = outerCell.locator('> table.nested-table');
    await expect(nested).toHaveCount(1);
    await expect(nested.locator('td')).toHaveText(['Inner one', 'Inner two']);
    // The content after the table is still below it.
    await expect(doc.pages.getByText('After the table')).toBeVisible();

    // Double-clicking a word in the nested table selects it there, and typing replaces it in that cell and nowhere else.
    const innerTwo = doc.pages.getByText('Inner two', { exact: true });
    const box = (await innerTwo.boundingBox())!;
    await innerTwo.dblclick({ position: { x: box.width - 6, y: box.height / 2 } });
    await doc.settle();
    await doc.typeText('2nd');
    await expect(nested.locator('td').nth(1)).toHaveText('Inner 2nd');
    await expect(nested.locator('td').first()).toHaveText('Inner one');
    await expect(outerCell.locator('> p.cell-paragraph').nth(1)).toHaveText('Cell second');
    await expect(doc.editor.locator('table table td').nth(1)).toContainText('Inner 2nd');
  });
});
