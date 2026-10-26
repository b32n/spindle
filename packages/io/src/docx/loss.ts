import JSZip from 'jszip';
import { countMatches, createLossReport, type ImportLossItem, type ImportLossReport } from '@b32nio/spindle-shared';

/**
 * Scan a .docx for content the importer does not bring across. Cheap (zip +
 * pattern counts); never throws on a malformed file — an unreadable zip yields
 * an empty report and the import itself reports the error.
 */
export async function detectDocxLoss(bytes: Uint8Array): Promise<ImportLossReport> {
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
  add('docx.floatingImages', 'floating image(s)', countMatches(doc, /<wp:anchor[\s>/]/g), 'not imported (only inline images are)');
  add('docx.textBoxes', 'text box(es)/shape(s)', countMatches(doc, /<w:txbxContent[\s>/]/g), 'their text is not imported');
  add('docx.equations', 'equation(s)', countMatches(doc, /<m:oMath[\s>/]/g), 'not imported');
  add('docx.embeddedObjects', 'embedded object(s)', countMatches(doc, /<w:object[\s>/]/g), 'not imported');
  add('docx.charts', 'chart(s)/diagram(s)', names.filter((n) => /^word\/(charts|diagrams)\//.test(n) && /\.xml$/.test(n) && !/rels|style|colors/i.test(n)).length, 'not imported');

  const columns = [...doc.matchAll(/<w:cols\b[^>]*\bw:num="(\d+)"/g)].filter((m) => Number(m[1]) > 1).length;
  add('docx.columns', 'multi-column section(s)', columns, 'laid out as a single column');

  // Fields other than PAGE / NUMPAGES lose their cached result (TOC, REF, DATE, MERGEFIELD…).
  const instr = [...doc.matchAll(/<w:instrText[^>]*>([^<]*)<\/w:instrText>/g)].map((m) => m[1].trim().toUpperCase());
  const simple = [...doc.matchAll(/<w:fldSimple\b[^>]*\bw:instr="([^"]*)"/g)].map((m) => m[1].trim().toUpperCase());
  const unsupported = [...instr, ...simple].filter((f) => f && !(f.startsWith('PAGE') && !f.startsWith('PAGEREF')) && !f.startsWith('NUMPAGES'));
  add('docx.fields', 'field(s) (table of contents, cross-references, dates…)', unsupported.length, 'their displayed value is dropped');

  return createLossReport('docx', items);
}
