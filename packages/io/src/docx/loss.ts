import JSZip from 'jszip';
import { parseComments } from './comments';
import { parseNotes, referencedNoteIds } from './notes';
import { scanDocx } from './loss-scan';
import { countMatches, createLossReport, type ImportLossItem, type ImportLossReport } from '@b32nio/spindle-shared';

export interface DetectDocxLossOptions {
  /**
   * `'summary'` (default) is what the import-confirmation dialog shows: content the importer drops outright
   * (tracked changes, notes, comments, floating images, fields…). `'full'` also reports formatting and
   * structure the importer flattens or ignores (table structure, tabs and breaks, sections, …) and breaks fields down by kind; the DOCX fidelity harness uses it.
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

  // A note the body refers to is imported (with its text); one nothing in the body refers to, and any in a damaged part, is not.
  const noteLoss = async (kind: 'footnote' | 'endnote') => {
    const xml = await text(`word/${kind}s.xml`);
    const all = countMatches(xml, new RegExp(`<w:${kind}\\b(?![^>]*w:type="(?:separator|continuationSeparator|continuationNotice)")[^>]*>`, 'g'));
    // A reference inside a text box is not imported with the text box.
    const referenced = referencedNoteIds(doc.replace(/<w:txbxContent[\s\S]*?<\/w:txbxContent>/g, ''), kind);
    const imported = [...parseNotes(xml, kind).keys()].filter((id) => referenced.has(id)).length;
    return Math.max(0, all - imported);
  };
  add('docx.footnotes', 'footnote(s)', await noteLoss('footnote'), 'not imported (no reference to it in the body, or the part is damaged)');
  add('docx.endnotes', 'endnote(s)', await noteLoss('endnote'), 'not imported (no reference to it in the body, or the part is damaged)');

  // Comments are imported with their replies and anchors; only those in a damaged part are lost.
  const commentsXml = await text('word/comments.xml');
  const importedComments = parseComments({ comments: commentsXml }, []).threadOf.size;
  add('docx.comments', 'comment(s)', Math.max(0, countMatches(commentsXml, /<w:comment[\s>/]/g) - importedComments), 'not imported (the comments part is damaged)');
  add('docx.textBoxes', 'text box(es)/shape(s)', countMatches(doc, /<w:txbxContent[\s>/]/g), 'their text is not imported');
  add('docx.equations', 'equation(s)', countMatches(doc, /<m:oMath[\s>/]/g), 'not imported');
  add('docx.embeddedObjects', 'embedded object(s)', countMatches(doc, /<w:object[\s>/]/g), 'not imported');

  const columns = [...doc.matchAll(/<w:cols\b[^>]*\bw:num="(\d+)"/g)].filter((m) => Number(m[1]) > 1).length;
  add('docx.columns', 'multi-column section(s)', columns, 'kept for export; laid out as a single column');

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
