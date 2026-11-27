import { isWellFormed, relationshipTargets, type OoxmlParts } from '../fidelity-core/ooxml-package';
import type { Construct } from '../fidelity-core/types';

const doc = (p: OoxmlParts) => p.xml['word/document.xml'] ?? '';
const all = (p: OoxmlParts, re: RegExp, only?: RegExp) =>
  Object.entries(p.xml)
    .filter(([n]) => (only ? only.test(n) : /^word\/[^/]+\.xml$/.test(n)))
    .reduce((sum, [, x]) => sum + (x.match(re)?.length ?? 0), 0);
const n = (s: string, re: RegExp) => s.match(re)?.length ?? 0;
const part = (p: OoxmlParts, name: string) => p.xml[name] ?? '';
const headersFooters = /^word\/(header|footer)\d*\.xml$/;
const authors = (p: OoxmlParts) =>
  new Set([...doc(p).matchAll(/<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange)\b[^>]*\bw:author="([^"]*)"/g)].map((m) => m[1])).size;

const lists = (p: OoxmlParts) => part(p, 'word/numbering.xml');
/** Paragraphs that use a list instance (w:num) carrying a startOverride, i.e. a restarted or re-numbered list. */
function restartedListParagraphs(p: OoxmlParts): number {
  const restarted = new Set(
    [...lists(p).matchAll(/<w:num\b[^>]*\bw:numId="(\d+)"[^>]*>((?:(?!<\/w:num>)[\s\S])*)<\/w:num>/g)].filter((m) => /<w:startOverride\b/.test(m[2])).map((m) => m[1])
  );
  return [...doc(p).matchAll(/<w:numId w:val="(\d+)"/g)].filter((m) => restarted.has(m[1])).length;
}
const instr = (p: OoxmlParts, re: RegExp) =>
  n(
    [...doc(p).matchAll(/<w:instrText[^>]*>([^<]*)<\/w:instrText>/g)].map((m) => m[1]).join('\n') +
      [...doc(p).matchAll(/<w:fldSimple\b[^>]*\bw:instr="([^"]*)"/g)].map((m) => m[1]).join('\n'),
    re
  );

export const CONSTRUCTS: Array<Construct<OoxmlParts>> = [
  { id: 'revisions.insertions', label: 'Tracked insertions', count: (p) => n(doc(p), /<w:ins\b/g) },
  { id: 'revisions.deletions', label: 'Tracked deletions', count: (p) => n(doc(p), /<w:del\b/g) },
  { id: 'revisions.moves', label: 'Tracked moves', count: (p) => Math.min(n(doc(p), /<w:moveFrom\b/g), n(doc(p), /<w:moveTo\b/g)) },
  { id: 'revisions.formatting', label: 'Tracked formatting changes', count: (p) => n(doc(p), /<w:(?:rPrChange|pPrChange)\b/g) },
  { id: 'revisions.multiAuthor', label: 'Revisions by 3+ authors', count: (p) => (authors(p) >= 3 ? authors(p) : 0) },
  { id: 'comments', label: 'Comments', count: (p) => n(part(p, 'word/comments.xml'), /<w:comment\b/g) },
  { id: 'comments.threaded', label: 'Threaded comment replies', count: (p) => n(part(p, 'word/commentsExtended.xml'), /paraIdParent=/g) },
  { id: 'notes.footnotes', label: 'Footnotes', count: (p) => n(part(p, 'word/footnotes.xml'), /<w:footnote\b(?![^>]*w:type=)/g) },
  { id: 'notes.endnotes', label: 'Endnotes', count: (p) => n(part(p, 'word/endnotes.xml'), /<w:endnote\b(?![^>]*w:type=)/g) },
  { id: 'fields.toc', label: 'Table of contents field', count: (p) => instr(p, /\bTOC\b/g) },
  { id: 'fields.page', label: 'PAGE field', count: (p) => all(p, /<w:instrText[^>]*>\s*PAGE\b|w:instr="\s*PAGE\b/g) },
  { id: 'fields.pageref', label: 'PAGEREF field', count: (p) => instr(p, /\bPAGEREF\b/g) },
  { id: 'fields.ref', label: 'REF field', count: (p) => instr(p, /(?<![A-Z])REF\s/g) },
  { id: 'fields.seq', label: 'SEQ field', count: (p) => instr(p, /\bSEQ\b/g) },
  { id: 'fields.date', label: 'DATE field', count: (p) => instr(p, /\b(?:DATE|CREATEDATE|SAVEDATE)\b/g) },
  { id: 'fields.other', label: 'Other fields (NUMPAGES, AUTHOR, STYLEREF, ...)', count: (p) => instr(p, /\b(?:NUMPAGES|AUTHOR|TITLE|STYLEREF|MERGEFIELD|IF|SECTIONPAGES|NUMWORDS)\b/g) },
  { id: 'fields.bookmarks', label: 'Bookmarks', count: (p) => n(doc(p), /<w:bookmarkStart\b/g) },
  { id: 'hyperlinks.external', label: 'External hyperlinks', count: (p) => n(part(p, 'word/_rels/document.xml.rels'), /Type="[^"]*\/hyperlink"[^>]*TargetMode="External"|TargetMode="External"[^>]*Type="[^"]*\/hyperlink"/g) },
  { id: 'hyperlinks.internal', label: 'Internal hyperlinks (anchors)', count: (p) => n(doc(p), /<w:hyperlink\b[^>]*\bw:anchor=/g) },
  { id: 'headings', label: 'Headings', count: (p) => n(doc(p), /<w:pStyle w:val="(?:Heading|heading)\s?[1-9]"/g) },
  { id: 'lists.numbered', label: 'Numbered / bulleted lists', count: (p) => n(doc(p), /<w:numPr>/g) },
  { id: 'lists.multilevel', label: 'Multilevel lists', count: (p) => n(doc(p), /<w:ilvl w:val="[1-8]"/g) },
  { id: 'lists.restart', label: 'List restarts / start overrides (used by a paragraph)', count: (p) => restartedListParagraphs(p) },
  { id: 'styles.inheritance', label: 'Style inheritance (basedOn)', count: (p) => n(part(p, 'word/styles.xml'), /<w:basedOn\b/g) },
  { id: 'styles.linked', label: 'Linked paragraph/character styles', count: (p) => n(part(p, 'word/styles.xml'), /<w:link\b/g) },
  { id: 'tables.nested', label: 'Nested tables', count: (p) => n(doc(p), /<w:tc>(?:(?!<\/w:tc>)[\s\S])*?<w:tbl>/g) },
  { id: 'tables.mergeH', label: 'Horizontally merged cells', count: (p) => n(doc(p), /<w:gridSpan\b/g) },
  { id: 'tables.mergeV', label: 'Vertically merged cells', count: (p) => n(doc(p), /<w:vMerge\b/g) },
  { id: 'tables.borders', label: 'Table / cell borders', count: (p) => n(doc(p), /<w:(?:tblBorders|tcBorders)\b/g) },
  { id: 'tables.shading', label: 'Cell shading', count: (p) => n(doc(p), /<w:tcPr>(?:(?!<\/w:tcPr>)[\s\S])*<w:shd\b/g) },
  { id: 'tables.headerRow', label: 'Repeating header rows', count: (p) => n(doc(p), /<w:tblHeader\b/g) },
  { id: 'tables.autofit', label: 'Autofit tables', count: (p) => (doc(p).match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/g) ?? []).filter((t) => /<w:tblW w:w="\d+" w:type="auto"|<w:tblLayout w:type="autofit"/.test(t) && !/<w:tblLayout w:type="fixed"/.test(t)).length },
  { id: 'sections.landscape', label: 'Landscape sections', count: (p) => n(doc(p), /<w:pgSz\b[^>]*w:orient="landscape"/g) },
  { id: 'sections.portrait', label: 'Portrait sections', count: (p) => n(doc(p), /<w:pgSz\b(?![^>]*w:orient="landscape")/g) },
  { id: 'sections.columns', label: 'Newspaper columns', count: (p) => [...doc(p).matchAll(/<w:cols\b[^>]*\bw:num="(\d+)"/g)].filter((m) => Number(m[1]) > 1).length },
  { id: 'sections.firstPage', label: 'Different first-page header/footer', count: (p) => n(doc(p), /<w:titlePg\b/g) },
  { id: 'sections.oddEven', label: 'Odd/even headers and footers', count: (p) => (n(part(p, 'word/settings.xml'), /<w:evenAndOddHeaders\b/g) ? n(doc(p), /<w:(?:header|footer)Reference\b[^>]*w:type="even"/g) : 0) },
  { id: 'sections.pageNumberRestart', label: 'Page-numbering restart', count: (p) => n(doc(p), /<w:pgNumType\b[^>]*w:start=/g) },
  { id: 'images.inline', label: 'Inline images', count: (p) => n(doc(p), /<wp:inline\b(?:(?!<\/wp:inline>)[\s\S])*<pic:pic\b/g) },
  { id: 'images.floating', label: 'Floating images', count: (p) => n(doc(p), /<wp:anchor\b(?:(?!<\/wp:anchor>)[\s\S])*<pic:pic\b/g) },
  { id: 'images.wrapped', label: 'Wrapped (square/tight/top-bottom) images', count: (p) => n(doc(p), /<wp:wrap(?:Square|Tight|Through|TopAndBottom)\b/g) },
  { id: 'images.cropped', label: 'Cropped images', count: (p) => n(doc(p), /<a:srcRect\b[^>]*\b[ltrb]="[1-9]/g) },
  { id: 'shapes', label: 'Drawing shapes', count: (p) => n(doc(p), /<wps:wsp\b|<v:(?:rect|roundrect|oval|shape)\b/g) },
  { id: 'textboxes', label: 'Text boxes', count: (p) => n(doc(p), /<w:txbxContent\b/g) },
  { id: 'smartart', label: 'SmartArt (with fallback drawing)', count: (p) => Math.min(p.names.filter((x) => /^word\/diagrams\/data\d*\.xml$/.test(x)).length, p.names.filter((x) => /^word\/diagrams\/drawing\d*\.xml$/.test(x)).length) },
  { id: 'equations', label: 'OMML equations', count: (p) => n(doc(p), /<m:oMath\b/g) },
  { id: 'sdt', label: 'Content controls (SDT)', count: (p) => n(doc(p), /<w:sdt\b/g) },
  { id: 'forms.legacy', label: 'Legacy form fields', count: (p) => n(doc(p), /<w:ffData\b/g) },
  { id: 'fonts.embedded', label: 'Embedded fonts', count: (p) => p.names.filter((x) => /^word\/fonts\//.test(x)).length },
  { id: 'themes', label: 'Theme colours / fonts used in content', count: (p) => (p.names.some((x) => /^word\/theme\/theme\d+\.xml$/.test(x)) ? all(p, /w:themeColor=|w:themeFill=|w:asciiTheme=/g, /^word\/(?:document|header\d*|footer\d*)\.xml$/) : 0) },
  { id: 'i18n.rtl', label: 'Right-to-left / bidi text', count: (p) => n(doc(p), /<w:bidi\b|<w:rtl\b/g) },
  { id: 'i18n.cjk', label: 'CJK text', count: (p) => n(doc(p), /[぀-ヿ㐀-鿿가-힯]/g) },
  { id: 'text.symbols', label: 'Symbols', count: (p) => n(doc(p), /<w:sym\b|[←-⏿☀-➿]/g) },
  { id: 'text.emoji', label: 'Emoji', count: (p) => n(doc(p), /[\u{1F300}-\u{1FAFF}]/gu) },
  { id: 'typography.hyphenation', label: 'Hyphenation', count: (p) => n(part(p, 'word/settings.xml'), /<w:autoHyphenation\b/g) + n(doc(p), /<w:noHyphen\b|<w:suppressAutoHyphens\b/g) },
  { id: 'typography.tabLeaders', label: 'Tabs with leaders', count: (p) => all(p, /<w:(?:tab|ptab)\b[^>]*w:leader="(?:dot|hyphen|underscore|heavy|middleDot)"/g, /^word\/(?:document|header\d*|footer\d*)\.xml$/) },
  { id: 'typography.dropCap', label: 'Drop caps', count: (p) => n(doc(p), /<w:framePr\b[^>]*w:dropCap=/g) },
  { id: 'watermark', label: 'Watermarks', count: (p) => all(p, /PowerPlusWaterMarkObject/g, headersFooters) },
  { id: 'charts', label: 'Charts', count: (p) => p.names.filter((x) => /^word\/charts\/chart\d*\.xml$/.test(x)).length },
  { id: 'objects.embedded', label: 'Embedded objects', count: (p) => n(doc(p), /<w:object\b/g) },
  { id: 'text.hidden', label: 'Hidden (vanish) text', count: (p) => n(doc(p), /<w:vanish\b/g) },
  { id: 'customXml', label: 'Custom XML parts', count: (p) => p.names.filter((x) => /^customXml\/item\d+\.xml$/.test(x)).length },
  { id: 'doc.long', label: 'Long document (15,000+ words)', count: (p) => { const words = n(doc(p).replace(/<[^>]+>/g, ' '), /\S+/g); return words >= 15000 ? words : 0; } },
  { id: 'headers.content', label: 'Content in headers/footers', count: (p) => all(p, /<w:t[ >]/g, headersFooters) },
  {
    id: 'edge.brokenRels',
    label: 'Relationship targets that do not exist',
    count: (p) => {
      let missing = 0;
      for (const [name, x] of Object.entries(p.xml))
        if (name.endsWith('.rels'))
          missing += relationshipTargets(name, x).filter((t) => t.target !== null && !p.names.includes(t.target)).length;
      return missing;
    },
  },
  { id: 'edge.invalidXml', label: 'Part that is not well-formed XML', count: (p) => Object.entries(p.xml).filter(([, x]) => !isWellFormed(x)).length },
  { id: 'edge.missingOptionalParts', label: 'Missing optional parts (styles/settings/numbering)', count: (p) => ['word/styles.xml', 'word/settings.xml', 'word/numbering.xml'].filter((x) => !p.names.includes(x)).length },
  {
    id: 'edge.zipBomb',
    label: 'Part over the import compression-ratio limit (zip bomb)',
    // mirrors DEFAULT_IMPORT_LIMITS (packages/io/src/common/safe-zip.ts): maxRatio 200 once a part inflates past ratioFloorBytes 16 MiB.
    // Not imported: this analyser must stay independent of the code under test and of the built io package; the edge-zip-bomb golden fails if the two drift.
    count: (p) => Object.values(p.sizes).filter((s) => s.uncompressed > 16 * 1024 * 1024 && s.uncompressed > 200 * Math.max(s.compressed, 1)).length,
  },
  { id: 'edge.entityExpansion', label: 'XML entity declarations (DOCTYPE)', count: (p) => Object.values(p.xml).reduce((sum, x) => sum + n(x, /<!ENTITY\b/g), 0) },
];
