import JSZip from 'jszip';

/** The OOXML parts of a .docx: part names plus the text of every XML/rels part. */
export interface DocxParts {
  names: string[];
  xml: Record<string, string>;
}

export async function readParts(bytes: Uint8Array): Promise<DocxParts> {
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const xml: Record<string, string> = {};
  for (const n of names) if (/\.(xml|rels)$/i.test(n)) xml[n] = await zip.files[n].async('string');
  return { names, xml };
}

export interface Construct {
  id: string;
  label: string;
  /** How many occurrences of the construct the file contains (0 = absent). */
  count: (p: DocxParts) => number;
}

const doc = (p: DocxParts) => p.xml['word/document.xml'] ?? '';
const all = (p: DocxParts, re: RegExp, only?: RegExp) =>
  Object.entries(p.xml)
    .filter(([n]) => (only ? only.test(n) : /^word\/[^/]+\.xml$/.test(n)))
    .reduce((sum, [, x]) => sum + (x.match(re)?.length ?? 0), 0);
const n = (s: string, re: RegExp) => s.match(re)?.length ?? 0;
const part = (p: DocxParts, name: string) => p.xml[name] ?? '';
const headersFooters = /^word\/(header|footer)\d*\.xml$/;
const authors = (p: DocxParts) =>
  new Set([...doc(p).matchAll(/<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange)\b[^>]*\bw:author="([^"]*)"/g)].map((m) => m[1])).size;

/** Tag-balance check; true when every element opens and closes properly (comments/CDATA/PIs skipped). */
export function isWellFormed(xml: string): boolean {
  const stack: string[] = [];
  const stripped = xml.replace(/<!--[\s\S]*?-->/g, '').replace(/<!\[CDATA\[[\s\S]*?\]\]>/g, '').replace(/<\?[\s\S]*?\?>/g, '');
  const re = /<(\/?)([A-Za-z_][\w:.-]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;
  let last = 0;
  for (let m = re.exec(stripped); m; m = re.exec(stripped)) {
    if (stripped.slice(last, m.index).includes('<')) return false;
    last = m.index + m[0].length;
    if (m[4]) continue;
    if (m[1]) {
      if (stack.pop() !== m[2]) return false;
    } else stack.push(m[2]);
  }
  return stack.length === 0 && !stripped.slice(last).includes('<');
}

const lists = (p: DocxParts) => part(p, 'word/numbering.xml');
const instr = (p: DocxParts, re: RegExp) =>
  n(
    [...doc(p).matchAll(/<w:instrText[^>]*>([^<]*)<\/w:instrText>/g)].map((m) => m[1]).join('\n') +
      [...doc(p).matchAll(/<w:fldSimple\b[^>]*\bw:instr="([^"]*)"/g)].map((m) => m[1]).join('\n'),
    re
  );

export const CONSTRUCTS: Construct[] = [
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
  { id: 'lists.restart', label: 'List restarts / start overrides', count: (p) => n(lists(p), /<w:startOverride\b/g) },
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
  { id: 'themes', label: 'Theme colours / fonts', count: (p) => (p.names.some((x) => /^word\/theme\/theme\d+\.xml$/.test(x)) ? n(doc(p), /w:themeColor=|w:asciiTheme=/g) + n(part(p, 'word/styles.xml'), /w:asciiTheme=|w:themeColor=/g) : 0) },
  { id: 'i18n.rtl', label: 'Right-to-left / bidi text', count: (p) => n(doc(p), /<w:bidi\b|<w:rtl\b/g) },
  { id: 'i18n.cjk', label: 'CJK text', count: (p) => n(doc(p), /[぀-ヿ㐀-鿿가-힯]/g) },
  { id: 'text.symbols', label: 'Symbols', count: (p) => n(doc(p), /<w:sym\b|[←-⏿☀-➿]/g) },
  { id: 'text.emoji', label: 'Emoji', count: (p) => n(doc(p), /[\u{1F300}-\u{1FAFF}]/gu) },
  { id: 'typography.hyphenation', label: 'Hyphenation', count: (p) => n(part(p, 'word/settings.xml'), /<w:autoHyphenation\b/g) + n(doc(p), /<w:noHyphen\b|<w:suppressAutoHyphens\b/g) },
  { id: 'typography.tabLeaders', label: 'Tabs with leaders', count: (p) => n(Object.values(p.xml).join('\n'), /<w:tab\b[^>]*w:leader="(?:dot|hyphen|underscore|heavy|middleDot)"/g) },
  { id: 'typography.dropCap', label: 'Drop caps', count: (p) => n(doc(p), /<w:framePr\b[^>]*w:dropCap=/g) },
  { id: 'watermark', label: 'Watermarks', count: (p) => all(p, /PowerPlusWaterMarkObject/g, headersFooters) },
  { id: 'charts', label: 'Charts', count: (p) => p.names.filter((x) => /^word\/charts\/chart\d*\.xml$/.test(x)).length },
  { id: 'objects.embedded', label: 'Embedded objects', count: (p) => n(doc(p), /<w:object\b/g) },
  { id: 'text.hidden', label: 'Hidden (vanish) text', count: (p) => n(doc(p), /<w:vanish\b/g) },
  { id: 'customXml', label: 'Custom XML parts', count: (p) => p.names.filter((x) => /^customXml\/item\d+\.xml$/.test(x)).length },
  { id: 'doc.long', label: 'Long document (15,000+ words)', count: (p) => { const words = n(doc(p).replace(/<[^>]+>/g, ' '), /\S+/g); return words >= 15000 ? words : 0; } },
  { id: 'content.headersOnly', label: 'Content in headers/footers', count: (p) => all(p, /<w:t[ >]/g, headersFooters) },
  {
    id: 'edge.brokenRels',
    label: 'Relationship targets that do not exist',
    count: (p) => {
      let missing = 0;
      for (const [name, x] of Object.entries(p.xml)) {
        if (!name.endsWith('.rels')) continue;
        const base = name.replace(/_rels\/[^/]*\.rels$/, '');
        for (const m of x.matchAll(/<Relationship\b[^>]*>/g)) {
          const target = /Target="([^"]*)"/.exec(m[0])?.[1];
          if (!target || /TargetMode="External"/.test(m[0])) continue;
          const resolved = target.startsWith('/') ? target.slice(1) : new URL(target, `http://x/${base}`).pathname.slice(1);
          if (!p.names.includes(decodeURIComponent(resolved))) missing++;
        }
      }
      return missing;
    },
  },
  { id: 'edge.invalidXml', label: 'Part that is not well-formed XML', count: (p) => Object.entries(p.xml).filter(([, x]) => !isWellFormed(x)).length },
  { id: 'edge.missingOptionalParts', label: 'Missing optional parts (styles/settings/numbering)', count: (p) => ['word/styles.xml', 'word/settings.xml', 'word/numbering.xml'].filter((x) => !p.names.includes(x)).length },
];

export function findConstruct(id: string): Construct | undefined {
  return CONSTRUCTS.find((c) => c.id === id);
}

/** Construct ids a fixture claims but whose probe finds nothing in the file. */
export function unverifiedClaims(claimed: string[], parts: DocxParts): string[] {
  return claimed.filter((id) => (findConstruct(id)?.count(parts) ?? 0) === 0);
}

/** Constructs actually found in the file (for reporting; claims are what the manifest lists). */
export function detectConstructs(parts: DocxParts): string[] {
  return CONSTRUCTS.filter((c) => c.count(parts) > 0).map((c) => c.id);
}
