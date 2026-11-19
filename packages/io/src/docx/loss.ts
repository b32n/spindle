import JSZip from 'jszip';
import { scanDocx } from './loss-scan';
import { countMatches, createLossReport, type ImportLossItem, type ImportLossReport } from '@b32nio/spindle-shared';

export interface DetectDocxLossOptions {
  /**
   * `'summary'` (default) is what the import-confirmation dialog shows: content the importer drops outright
   * (tracked changes, notes, comments, floating images, fields…). `'full'` also reports formatting and
   * structure the importer flattens or ignores (named styles, table formatting, tabs and breaks, sections,
   * bookmarks, …) and breaks fields down by kind; the DOCX fidelity harness uses it.
   */
  detail?: 'summary' | 'full';
}

/**
 * Scan a .docx for content the importer does not bring across. Cheap (zip +
 * pattern counts); never throws on a malformed file — an unreadable zip yields
 * an empty report and the import itself reports the error.
 */
export async function detectDocxLoss(bytes: Uint8Array, options: DetectDocxLossOptions = {}): Promise<ImportLossReport> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    return createLossReport('docx', []);
  }
  const text = async (path: string) => (await zip.file(path)?.async('string')) ?? '';
  const names = Object.keys(zip.files);
  const doc = await text('word/document.xml');
  const items: ImportLossItem[] = [];
  const add = (id: string, label: string, count: number, effect: string) => {
    if (count > 0) items.push({ id, label, count, effect });
  };

  add('docx.trackedInsertions', 'tracked insertion(s)', countMatches(doc, /<w:ins[\s>/]/g), 'their text is not imported');
  add('docx.trackedDeletions', 'tracked deletion(s)', countMatches(doc, /<w:del[\s>/]/g), 'applied (the deleted text is gone)');
  add('docx.contentControls', 'content control(s)', countMatches(doc, /<w:sdt[\s>/]/g), 'their content is not imported');

  const footnotes = await text('word/footnotes.xml');
  const endnotes = await text('word/endnotes.xml');
  const notes = (xml: string) => countMatches(xml, /<w:(?:footnote|endnote)\b(?![^>]*w:type="(?:separator|continuationSeparator|continuationNotice)")[^>]*>/g);
  add('docx.footnotes', 'footnote(s)', notes(footnotes), 'not imported');
  add('docx.endnotes', 'endnote(s)', notes(endnotes), 'not imported');

  add('docx.comments', 'comment(s)', countMatches(await text('word/comments.xml'), /<w:comment[\s>/]/g), 'not imported');
  add('docx.textBoxes', 'text box(es)/shape(s)', countMatches(doc, /<w:txbxContent[\s>/]/g), 'their text is not imported');
  add('docx.equations', 'equation(s)', countMatches(doc, /<m:oMath[\s>/]/g), 'not imported');
  add('docx.embeddedObjects', 'embedded object(s)', countMatches(doc, /<w:object[\s>/]/g), 'not imported');

  const columns = [...doc.matchAll(/<w:cols\b[^>]*\bw:num="(\d+)"/g)].filter((m) => Number(m[1]) > 1).length;
  add('docx.columns', 'multi-column section(s)', columns, 'laid out as a single column');

  // Constructs that need element context (styles, tables, sections, fields…) are counted on the parsed XML.
  items.push(
    ...scanDocx(
      {
        document: doc,
        styles: await text('word/styles.xml'),
        numbering: await text('word/numbering.xml'),
        settings: await text('word/settings.xml'),
        headerFooters: await Promise.all(names.filter((n) => /^word\/(header|footer)\d*\.xml$/.test(n)).map(text)),
        fontFiles: names.filter((n) => /^word\/fonts\/[^/]+\.(odttf|ttf|otf|fntdata)$/i.test(n)).length,
      },
      options.detail === 'full'
    )
  );

  return createLossReport('docx', items);
}
