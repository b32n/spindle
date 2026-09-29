import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import type {
  DocBlock,
  DocHeaderFooter,
  DocHeaderFooterParagraph,
  DocInline,
  DocParagraphStyle,
  DocRunBase,
  DocSection,
  DocTableRow,
  DocumentImportResult,
} from './types';

// ---------------------------------------------------------------------------
// XML tree utilities. Parsed with preserveOrder: true — required because a
// document body interleaves <w:p> and <w:tbl> children in document order;
// fast-xml-parser's default mode groups same-tag siblings into separate
// arrays and loses that interleaving entirely. The tradeoff is a more
// verbose node shape: each element is `{ tagName: [children], ':@': attrs }`.
// ---------------------------------------------------------------------------

type XmlNode = Record<string, unknown> & { ':@'?: Record<string, string> };

const xmlParser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
});

function parseXml(xml: string): XmlNode[] {
  return xmlParser.parse(xml) as XmlNode[];
}

/** The single non-attribute key of a preserveOrder node, e.g. "w:p". */
function tagOf(node: XmlNode): string | undefined {
  return Object.keys(node).find((k) => k !== ':@');
}

function childrenOf(node: XmlNode): XmlNode[] {
  const tag = tagOf(node);
  if (!tag) return [];
  const value = node[tag];
  return Array.isArray(value) ? (value as XmlNode[]) : [];
}

function attrsOf(node: XmlNode): Record<string, string> {
  return node[':@'] ?? {};
}

function findAll(nodes: XmlNode[], tag: string): XmlNode[] {
  return nodes.filter((n) => tagOf(n) === tag);
}

function find(nodes: XmlNode[], tag: string): XmlNode | undefined {
  return nodes.find((n) => tagOf(n) === tag);
}

/** Depth-first search for the first descendant with the given tag, regardless of exact nesting path (drawing markup nests several levels deep and the path is more than we want to hard-code). */
function findDeep(node: XmlNode, tag: string): XmlNode | undefined {
  for (const child of childrenOf(node)) {
    if (tagOf(child) === tag) return child;
    const found = findDeep(child, tag);
    if (found) return found;
  }
  return undefined;
}

/** Depth-first concatenation of every #text leaf under a node (a run's <w:t>, a field's cached-result text, ...). */
function textOf(node: XmlNode): string {
  let out = '';
  for (const child of childrenOf(node)) {
    const tag = tagOf(child);
    if (tag === '#text') out += String((child as Record<string, unknown>)['#text'] ?? '');
    else out += textOf(child);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Relationships (word/_rels/document.xml.rels and friends): rId -> target.
// A hyperlink's URL, and a header/footer's part filename, are never inline —
// always one level of indirection through here.
// ---------------------------------------------------------------------------

function parseRels(xml: string | undefined): Map<string, string> {
  const rels = new Map<string, string>();
  if (!xml) return rels;
  const root = parseXml(xml);
  const relationships = find(root, 'Relationships');
  if (!relationships) return rels;
  for (const rel of findAll(childrenOf(relationships), 'Relationship')) {
    const a = attrsOf(rel);
    if (a['@_Id'] && a['@_Target']) rels.set(a['@_Id'], a['@_Target']);
  }
  return rels;
}

/** Everything paragraph/table/header/footer parsing needs to resolve a part's cross-references: hyperlink and header/footer targets, plus embedded images pre-loaded as data: URIs (keyed by rId, since loading them is async and the parse tree walk below it isn't). */
interface PartContext {
  rels: Map<string, string>;
  media: Map<string, string>;
}

const MEDIA_EXT_TO_MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };

async function buildMediaMap(zip: JSZip, rels: Map<string, string>): Promise<Map<string, string>> {
  const media = new Map<string, string>();
  for (const [rId, target] of rels) {
    if (!/(^|\/)media\//i.test(target)) continue;
    const extMatch = target.match(/\.([a-zA-Z]+)$/);
    const mime = extMatch ? MEDIA_EXT_TO_MIME[extMatch[1].toLowerCase()] : undefined;
    if (!mime) continue;
    const file = zip.file(`word/${target}`);
    if (!file) continue;
    const base64 = await file.async('base64');
    media.set(rId, `data:${mime};base64,${base64}`);
  }
  return media;
}

// ---------------------------------------------------------------------------
// Styles (word/styles.xml): styleId -> style name. Real files (Word, Google
// Docs, LibreOffice) don't reliably use "Heading1" as the styleId itself —
// the human-readable name ("Heading 1") is the only convention that's
// actually universal, so heading detection goes through this map rather
// than assuming a styleId spelling.
// ---------------------------------------------------------------------------

function parseStyleNames(xml: string | undefined): Map<string, string> {
  const names = new Map<string, string>();
  if (!xml) return names;
  const root = parseXml(xml);
  const stylesEl = find(root, 'w:styles');
  if (!stylesEl) return names;
  for (const style of findAll(childrenOf(stylesEl), 'w:style')) {
    const id = attrsOf(style)['@_w:styleId'];
    const nameEl = find(childrenOf(style), 'w:name');
    const name = nameEl ? attrsOf(nameEl)['@_w:val'] : undefined;
    if (id && name) names.set(id, name);
  }
  return names;
}

const HEADING_STYLE_NAME = /^Heading\s*(\d)$/i;

// ---------------------------------------------------------------------------
// Numbering (word/numbering.xml): numId -> abstractNumId -> (per level) format.
// A list item only tells you its numId + ilvl; the bullet-vs-numbered
// distinction lives in this separate part.
// ---------------------------------------------------------------------------

interface NumberingInfo {
  numIdToAbstract: Map<string, string>;
  abstractLevelFormat: Map<string, Map<number, string>>; // abstractNumId -> ilvl -> w:numFmt value
}

function parseNumbering(xml: string | undefined): NumberingInfo {
  const numIdToAbstract = new Map<string, string>();
  const abstractLevelFormat = new Map<string, Map<number, string>>();
  if (!xml) return { numIdToAbstract, abstractLevelFormat };
  const root = parseXml(xml);
  const numberingEl = find(root, 'w:numbering');
  if (!numberingEl) return { numIdToAbstract, abstractLevelFormat };

  for (const abstractNum of findAll(childrenOf(numberingEl), 'w:abstractNum')) {
    const abstractId = attrsOf(abstractNum)['@_w:abstractNumId'];
    if (!abstractId) continue;
    const levels = new Map<number, string>();
    for (const lvl of findAll(childrenOf(abstractNum), 'w:lvl')) {
      const ilvl = Number(attrsOf(lvl)['@_w:ilvl'] ?? '0');
      const numFmtEl = find(childrenOf(lvl), 'w:numFmt');
      const fmt = numFmtEl ? attrsOf(numFmtEl)['@_w:val'] : undefined;
      if (fmt) levels.set(ilvl, fmt);
    }
    abstractLevelFormat.set(abstractId, levels);
  }
  for (const num of findAll(childrenOf(numberingEl), 'w:num')) {
    const numId = attrsOf(num)['@_w:numId'];
    const abstractEl = find(childrenOf(num), 'w:abstractNumId');
    const abstractId = abstractEl ? attrsOf(abstractEl)['@_w:val'] : undefined;
    if (numId && abstractId) numIdToAbstract.set(numId, abstractId);
  }
  return { numIdToAbstract, abstractLevelFormat };
}

function listTypeFor(numbering: NumberingInfo, numId: string, ilvl: number): 'bullet' | 'numbered' {
  const abstractId = numbering.numIdToAbstract.get(numId);
  const fmt = abstractId ? numbering.abstractLevelFormat.get(abstractId)?.get(ilvl) : undefined;
  return fmt === 'bullet' ? 'bullet' : 'numbered';
}

// ---------------------------------------------------------------------------
// Runs and inline content
// ---------------------------------------------------------------------------

const HIGHLIGHT_TO_HEX: Record<string, string> = {
  black: '#000000', blue: '#0000ff', cyan: '#00ffff', darkBlue: '#00008b', darkCyan: '#008b8b',
  darkGray: '#a9a9a9', darkGreen: '#006400', darkMagenta: '#8b008b', darkRed: '#8b0000',
  darkYellow: '#808000', green: '#00ff00', lightGray: '#d3d3d3', magenta: '#ff00ff',
  red: '#ff0000', white: '#ffffff', yellow: '#ffff00',
};

function runProps(rPrNode: XmlNode | undefined): DocRunBase {
  if (!rPrNode) return {};
  const props = childrenOf(rPrNode);
  const bold = !!find(props, 'w:b') && attrsOf(find(props, 'w:b')!)['@_w:val'] !== 'false';
  const italic = !!find(props, 'w:i') && attrsOf(find(props, 'w:i')!)['@_w:val'] !== 'false';
  const strike = !!find(props, 'w:strike') && attrsOf(find(props, 'w:strike')!)['@_w:val'] !== 'false';
  const uEl = find(props, 'w:u');
  const underline = !!uEl && attrsOf(uEl)['@_w:val'] !== 'none';
  const smallCaps = !!find(props, 'w:smallCaps');
  const vertAlignEl = find(props, 'w:vertAlign');
  const vertAlign = vertAlignEl ? attrsOf(vertAlignEl)['@_w:val'] : undefined;
  const colorEl = find(props, 'w:color');
  const colorVal = colorEl ? attrsOf(colorEl)['@_w:val'] : undefined;
  const color = colorVal && colorVal !== 'auto' ? `#${colorVal.toLowerCase()}` : undefined;
  const szEl = find(props, 'w:sz');
  const fontSize = szEl ? Number(attrsOf(szEl)['@_w:val']) / 2 : undefined;
  const fontsEl = find(props, 'w:rFonts');
  const fontFamily = fontsEl ? attrsOf(fontsEl)['@_w:ascii'] : undefined;
  const shdEl = find(props, 'w:shd');
  const fill = shdEl ? attrsOf(shdEl)['@_w:fill'] : undefined;
  const highlightEl = find(props, 'w:highlight');
  const highlightVal = highlightEl ? attrsOf(highlightEl)['@_w:val'] : undefined;
  const backgroundColor =
    fill && fill !== 'auto' && fill !== 'none' ? `#${fill.toLowerCase()}` : highlightVal ? HIGHLIGHT_TO_HEX[highlightVal] : undefined;

  return {
    ...(bold && { bold: true }),
    ...(italic && { italic: true }),
    ...(strike && { strikethrough: true }),
    ...(underline && { underline: true }),
    ...(smallCaps && { smallCaps: true }),
    ...(vertAlign === 'superscript' && { superscript: true }),
    ...(vertAlign === 'subscript' && { subscript: true }),
    ...(color && { color }),
    ...(fontSize && { fontSize }),
    ...(fontFamily && { fontFamily }),
    ...(backgroundColor && { backgroundColor }),
  };
}

/**
 * Walk a paragraph's direct children, producing inline content. Handles
 * plain runs, hyperlinks (resolved through rels), and field-code sequences
 * for PAGE/NUMPAGES — every other field type's cached display text is
 * dropped with a warning rather than silently kept (its literal value
 * would be stale outside Word's own pagination).
 *
 * A field is a fldChar "begin", an instrText carrying the instruction, an
 * optional "separate" marker, optional cached display text, then a fldChar
 * "end" — spread across sibling <w:r> elements per the OOXML spec, but our
 * own writer (and possibly other producers) packs the whole sequence into
 * a single <w:r>'s children instead. Walking element-by-element within
 * each run, with field state carried across runs, handles both shapes.
 * <w:fldSimple> is the shorthand form some producers use in place of the
 * begin/end sequence. A run's <w:drawing> (inline image) is resolved
 * through ctx.media, pre-loaded as a data: URI since loading it is async
 * and this tree walk isn't.
 */
function paragraphInlineContent(pNode: XmlNode, ctx: PartContext, warnings: string[]): DocInline[] {
  const out: DocInline[] = [];
  let pendingField = ''; // accumulated instrText while inside a begin/.../end sequence
  let inField = false;

  const consumeRun = (runNode: XmlNode, target: string | undefined) => {
    const rPr = find(childrenOf(runNode), 'w:rPr');
    const props = runProps(rPr);
    let text = '';
    for (const piece of childrenOf(runNode)) {
      const pieceTag = tagOf(piece);
      if (pieceTag === 'w:fldChar') {
        const type = attrsOf(piece)['@_w:fldCharType'];
        if (type === 'begin') {
          inField = true;
          pendingField = '';
        } else if (type === 'end') {
          if (inField) out.push(...fieldRunFor(pendingField, warnings));
          inField = false;
          pendingField = '';
        }
        // 'separate' — no-op marker between the instruction and its cached result.
      } else if (pieceTag === 'w:instrText') {
        if (inField) pendingField += textOf(piece);
      } else if (pieceTag === 'w:t') {
        if (!inField) text += textOf(piece); // inField's w:t is cached display text — discard, we use the instruction
      } else if (pieceTag === 'w:drawing') {
        const image = parseDrawing(piece, ctx.media, warnings);
        if (image) out.push(image);
      }
    }
    if (text) {
      if (target) out.push({ type: 'link', text, href: target, ...props });
      else out.push({ type: 'text', text, ...props });
    }
  };

  const consumeRuns = (runs: XmlNode[], target: string | undefined) => {
    for (const child of runs) {
      const tag = tagOf(child);
      if (tag === 'w:fldSimple') out.push(...fieldRunFor(attrsOf(child)['@_w:instr'] ?? '', warnings));
      else if (tag === 'w:r') consumeRun(child, target);
    }
  };

  for (const child of childrenOf(pNode)) {
    const tag = tagOf(child);
    if (tag === 'w:hyperlink') {
      const target = resolveHyperlinkTarget(child, ctx.rels, warnings);
      consumeRuns(childrenOf(child), target);
    } else if (tag === 'w:r' || tag === 'w:fldSimple') {
      consumeRuns([child], undefined);
    }
    // Bookmarks, proofErr, and other structural markers carry no content — skipped.
  }

  return out;
}

const EMU_TO_PX = 1 / 9525; // 914400 EMU/inch, 96px/inch

/** An inline (non-floating) image drawing: <w:drawing><wp:inline>...<a:blip r:embed="..."/></wp:inline></w:drawing>. Floating (wp:anchor) images aren't positioned the same way our block model supports, so they're dropped with a warning. */
function parseDrawing(drawingNode: XmlNode, media: Map<string, string>, warnings: string[]): DocInline | null {
  const inline = findDeep(drawingNode, 'wp:inline');
  if (!inline) {
    warnings.push('A floating (anchored) image was skipped — only inline images are supported.');
    return null;
  }
  const extent = findDeep(inline, 'wp:extent');
  const extentAttrs = extent ? attrsOf(extent) : {};
  const width = extentAttrs['@_cx'] ? Math.round(Number(extentAttrs['@_cx']) * EMU_TO_PX) : 0;
  const height = extentAttrs['@_cy'] ? Math.round(Number(extentAttrs['@_cy']) * EMU_TO_PX) : 0;
  const docPr = findDeep(inline, 'wp:docPr');
  const docPrAttrs = docPr ? attrsOf(docPr) : {};
  const alt = docPrAttrs['@_descr'] || docPrAttrs['@_name'] || undefined;
  const blip = findDeep(inline, 'a:blip');
  const rId = blip ? attrsOf(blip)['@_r:embed'] : undefined;
  const src = rId ? media.get(rId) : undefined;
  if (!src) {
    warnings.push('An embedded image could not be resolved and was skipped.');
    return null;
  }
  return { type: 'image', src, width, height, ...(alt && { alt }) };
}

function fieldRunFor(instruction: string, warnings: string[]): DocInline[] {
  const norm = instruction.trim().toUpperCase();
  if (norm.startsWith('PAGE') && !norm.startsWith('PAGEREF')) return [{ type: 'pageNumber' }];
  if (norm.startsWith('NUMPAGES')) return [{ type: 'totalPages' }];
  if (!norm) return [];
  warnings.push(`Unsupported field code "${instruction.trim()}" — its cached value was dropped.`);
  return [];
}

function resolveHyperlinkTarget(hyperlinkNode: XmlNode, rels: Map<string, string>, warnings: string[]): string | undefined {
  const rId = attrsOf(hyperlinkNode)['@_r:id'];
  if (!rId) return undefined;
  const target = rels.get(rId);
  if (!target) warnings.push(`Hyperlink relationship "${rId}" could not be resolved.`);
  return target;
}

// ---------------------------------------------------------------------------
// Paragraphs (-> paragraph / heading / listItem blocks)
// ---------------------------------------------------------------------------

const ALIGN_MAP: Record<string, DocParagraphStyle['alignment']> = {
  left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify',
};

function paragraphStyle(pPr: XmlNode | undefined): DocParagraphStyle {
  if (!pPr) return {};
  const kids = childrenOf(pPr);
  const jcEl = find(kids, 'w:jc');
  const alignment = jcEl ? ALIGN_MAP[attrsOf(jcEl)['@_w:val']] : undefined;
  const spacingEl = find(kids, 'w:spacing');
  const spacingAttrs = spacingEl ? attrsOf(spacingEl) : {};
  const spaceBeforePt = spacingAttrs['@_w:before'] ? Number(spacingAttrs['@_w:before']) / 20 : undefined;
  const spaceAfterPt = spacingAttrs['@_w:after'] ? Number(spacingAttrs['@_w:after']) / 20 : undefined;
  const indentEl = find(kids, 'w:ind');
  const indentAttrs = indentEl ? attrsOf(indentEl) : {};
  const leftTwips = indentAttrs['@_w:left'] ?? indentAttrs['@_w:start'];
  const leftIndentPt = leftTwips ? Number(leftTwips) / 20 : undefined;
  const firstLineTwips = indentAttrs['@_w:firstLine'];
  const firstLineIndentPt = firstLineTwips ? Number(firstLineTwips) / 20 : undefined;

  return {
    ...(alignment && { alignment }),
    ...(spaceBeforePt != null && { spaceBeforePt }),
    ...(spaceAfterPt != null && { spaceAfterPt }),
    ...(leftIndentPt != null && { leftIndentPt }),
    ...(firstLineIndentPt != null && { firstLineIndentPt }),
  };
}

function hasPageBreakRun(pNode: XmlNode): boolean {
  return childrenOf(pNode).some((child) => {
    if (tagOf(child) !== 'w:r') return false;
    const br = find(childrenOf(child), 'w:br');
    return !!br && attrsOf(br)['@_w:type'] === 'page';
  });
}

function hasBottomBorder(pPrKids: XmlNode[]): boolean {
  const pBdr = find(pPrKids, 'w:pBdr');
  return !!pBdr && !!find(childrenOf(pBdr), 'w:bottom');
}

function parseParagraph(
  pNode: XmlNode,
  styleNames: Map<string, string>,
  numbering: NumberingInfo,
  ctx: PartContext,
  warnings: string[]
): DocBlock {
  const kids = childrenOf(pNode);
  const pPr = find(kids, 'w:pPr');
  const pPrKids = pPr ? childrenOf(pPr) : [];
  const content = paragraphInlineContent(pNode, ctx, warnings);
  const style = paragraphStyle(pPr);

  // An otherwise-empty paragraph carrying only a page-break run or a
  // bottom-border (our own writer's idiom for a horizontal rule) maps back
  // to the corresponding structural block rather than an empty paragraph.
  if (content.length === 0 && hasPageBreakRun(pNode)) {
    return { type: 'pageBreak' };
  }
  if (content.length === 0 && hasBottomBorder(pPrKids)) {
    return { type: 'horizontalRule' };
  }

  // A paragraph whose only content is a single image is how our own writer
  // (and most producers) represents a block-level image — promote it
  // rather than returning a "paragraph" block wrapping one inline image.
  if (content.length === 1 && content[0].type === 'image') {
    const image = content[0];
    return {
      type: 'image',
      src: image.src,
      width: image.width,
      height: image.height,
      ...(image.alt && { alt: image.alt }),
      ...(style.alignment && style.alignment !== 'justify' ? { alignment: style.alignment } : {}),
    };
  }

  const numPr = find(pPrKids, 'w:numPr');
  if (numPr) {
    const numPrKids = childrenOf(numPr);
    const ilvlEl = find(numPrKids, 'w:ilvl');
    const numIdEl = find(numPrKids, 'w:numId');
    const level = ilvlEl ? Number(attrsOf(ilvlEl)['@_w:val']) : 0;
    const numId = numIdEl ? attrsOf(numIdEl)['@_w:val'] : undefined;
    const listType = numId ? listTypeFor(numbering, numId, level) : 'bullet';
    return { type: 'listItem', listType, level, content, ...style };
  }

  const pStyleEl = find(pPrKids, 'w:pStyle');
  const styleId = pStyleEl ? attrsOf(pStyleEl)['@_w:val'] : undefined;
  const styleName = styleId ? styleNames.get(styleId) : undefined;
  const headingMatch = styleName?.match(HEADING_STYLE_NAME);
  if (headingMatch) {
    const level = Math.min(6, Math.max(1, Number(headingMatch[1]))) as 1 | 2 | 3 | 4 | 5 | 6;
    return { type: 'heading', level, content, ...style };
  }

  return { type: 'paragraph', content, ...style };
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

function parseTable(tblNode: XmlNode, ctx: PartContext, warnings: string[]): DocBlock {
  const rowsXml = findAll(childrenOf(tblNode), 'w:tr');
  const rows: DocTableRow[] = [];
  // Column index -> the DocTableCell object currently accumulating a
  // vertical merge, so a later row's continuation cell can bump its rowspan
  // instead of emitting a new cell.
  const openVMerge = new Map<number, { cell: DocTableRow['cells'][number] }>();

  for (const trNode of rowsXml) {
    const cellsXml = findAll(childrenOf(trNode), 'w:tc');
    const rowCells: DocTableRow['cells'] = [];
    let col = 0;
    for (const tcNode of cellsXml) {
      const tcKids = childrenOf(tcNode);
      const tcPr = find(tcKids, 'w:tcPr');
      const tcPrKids = tcPr ? childrenOf(tcPr) : [];
      const gridSpanEl = find(tcPrKids, 'w:gridSpan');
      const colspan = gridSpanEl ? Number(attrsOf(gridSpanEl)['@_w:val']) : undefined;
      const vMergeEl = find(tcPrKids, 'w:vMerge');
      const vMergeVal = vMergeEl ? (attrsOf(vMergeEl)['@_w:val'] ?? 'continue') : undefined;

      if (vMergeVal === 'continue') {
        const open = openVMerge.get(col);
        if (open) open.cell.rowspan = (open.cell.rowspan ?? 1) + 1;
        col += colspan ?? 1;
        continue; // this cell contributes no content of its own
      }

      const content = collectCellContent(tcNode, ctx, warnings);
      const cell: DocTableRow['cells'][number] = { content, ...(colspan && colspan > 1 ? { colspan } : {}) };
      rowCells.push(cell);
      if (vMergeVal === 'restart') openVMerge.set(col, { cell });
      else openVMerge.delete(col);
      col += colspan ?? 1;
    }
    rows.push({ cells: rowCells });
  }

  return { type: 'table', rows };
}

/** A table cell's content is one or more paragraphs; flatten to the first non-empty paragraph's inline content (our own writer only ever emits one). */
function collectCellContent(tcNode: XmlNode, ctx: PartContext, warnings: string[]): DocInline[] {
  for (const p of findAll(childrenOf(tcNode), 'w:p')) {
    const content = paragraphInlineContent(p, ctx, warnings);
    if (content.length > 0) return content;
  }
  return [];
}

// ---------------------------------------------------------------------------
// Section properties (page size/margins/orientation, header/footer refs)
// ---------------------------------------------------------------------------

interface ParsedSectPr {
  pageWidthPx: number;
  pageHeightPx: number;
  marginTopPx: number;
  marginRightPx: number;
  marginBottomPx: number;
  marginLeftPx: number;
  orientation: 'portrait' | 'landscape';
  headerRId?: string;
  footerRId?: string;
}

const TWIPS_TO_PX = 96 / 1440;

function parseSectPr(sectPrNode: XmlNode): ParsedSectPr {
  const kids = childrenOf(sectPrNode);
  const pgSz = find(kids, 'w:pgSz');
  const pgSzAttrs = pgSz ? attrsOf(pgSz) : {};
  const orientation: 'portrait' | 'landscape' = pgSzAttrs['@_w:orient'] === 'landscape' ? 'landscape' : 'portrait';
  const widthTwips = Number(pgSzAttrs['@_w:w'] ?? 12240);
  const heightTwips = Number(pgSzAttrs['@_w:h'] ?? 15840);

  const pgMar = find(kids, 'w:pgMar');
  const marAttrs = pgMar ? attrsOf(pgMar) : {};
  const top = Number(marAttrs['@_w:top'] ?? 1440);
  const right = Number(marAttrs['@_w:right'] ?? 1440);
  const bottom = Number(marAttrs['@_w:bottom'] ?? 1440);
  const left = Number(marAttrs['@_w:left'] ?? 1440);

  const headerRef = find(kids, 'w:headerReference');
  const footerRef = find(kids, 'w:footerReference');

  return {
    pageWidthPx: Math.round(widthTwips * TWIPS_TO_PX),
    pageHeightPx: Math.round(heightTwips * TWIPS_TO_PX),
    marginTopPx: Math.round(top * TWIPS_TO_PX),
    marginRightPx: Math.round(right * TWIPS_TO_PX),
    marginBottomPx: Math.round(bottom * TWIPS_TO_PX),
    marginLeftPx: Math.round(left * TWIPS_TO_PX),
    orientation,
    headerRId: headerRef ? attrsOf(headerRef)['@_r:id'] : undefined,
    footerRId: footerRef ? attrsOf(footerRef)['@_r:id'] : undefined,
  };
}

// ---------------------------------------------------------------------------
// Headers / footers
// ---------------------------------------------------------------------------

function parseHeaderFooter(xml: string, ctx: PartContext, warnings: string[]): DocHeaderFooter {
  const root = parseXml(xml);
  const rootEl = find(root, 'w:hdr') ?? find(root, 'w:ftr');
  const paragraphs: DocHeaderFooterParagraph[] = [];
  if (rootEl) {
    for (const p of findAll(childrenOf(rootEl), 'w:p')) {
      const pPr = find(childrenOf(p), 'w:pPr');
      const jcEl = pPr ? find(childrenOf(pPr), 'w:jc') : undefined;
      const alignment = jcEl ? ALIGN_MAP[attrsOf(jcEl)['@_w:val']] : undefined;
      const content = paragraphInlineContent(p, ctx, warnings);
      paragraphs.push({ content, ...(alignment && alignment !== 'justify' ? { alignment } : {}) });
    }
  }
  return { paragraphs };
}

// ---------------------------------------------------------------------------
// Top-level orchestration
// ---------------------------------------------------------------------------

async function readZipText(zip: JSZip, path: string): Promise<string | undefined> {
  const file = zip.file(path);
  return file ? file.async('string') : undefined;
}

export async function importDocx(bytes: Uint8Array): Promise<DocumentImportResult> {
  const warnings: string[] = [];
  const zip = await JSZip.loadAsync(bytes);

  const [documentXml, docRelsXml, stylesXml, numberingXml] = await Promise.all([
    readZipText(zip, 'word/document.xml'),
    readZipText(zip, 'word/_rels/document.xml.rels'),
    readZipText(zip, 'word/styles.xml'),
    readZipText(zip, 'word/numbering.xml'),
  ]);

  if (!documentXml) {
    return { title: '', sections: [], warnings: ['word/document.xml is missing — this file may not be a valid .docx.'] };
  }

  const rels = parseRels(docRelsXml);
  const media = await buildMediaMap(zip, rels);
  const ctx: PartContext = { rels, media };
  const styleNames = parseStyleNames(stylesXml);
  const numbering = parseNumbering(numberingXml);

  const root = parseXml(documentXml);
  const documentEl = find(root, 'w:document');
  const bodyEl = documentEl ? find(childrenOf(documentEl), 'w:body') : undefined;
  const bodyChildren = bodyEl ? childrenOf(bodyEl) : [];

  const sections: DocSection[] = [];
  let currentBlocks: DocBlock[] = [];

  const flushSection = async (sectPrNode: XmlNode) => {
    const parsed = parseSectPr(sectPrNode);
    const header = parsed.headerRId ? await loadHeaderFooter(zip, rels, parsed.headerRId, warnings) : undefined;
    const footer = parsed.footerRId ? await loadHeaderFooter(zip, rels, parsed.footerRId, warnings) : undefined;
    sections.push({
      pageWidthPx: parsed.pageWidthPx,
      pageHeightPx: parsed.pageHeightPx,
      marginTopPx: parsed.marginTopPx,
      marginRightPx: parsed.marginRightPx,
      marginBottomPx: parsed.marginBottomPx,
      marginLeftPx: parsed.marginLeftPx,
      orientation: parsed.orientation,
      blocks: currentBlocks,
      ...(header && { header }),
      ...(footer && { footer }),
    });
    currentBlocks = [];
  };

  for (const node of bodyChildren) {
    const tag = tagOf(node);
    if (tag === 'w:p') {
      const pPr = find(childrenOf(node), 'w:pPr');
      const inlineSectPr = pPr ? find(childrenOf(pPr), 'w:sectPr') : undefined;
      if (inlineSectPr) {
        // A sectPr nested in a paragraph's pPr ends that section; the
        // paragraph itself carries no visible content in that case.
        await flushSection(inlineSectPr);
        continue;
      }
      currentBlocks.push(parseParagraph(node, styleNames, numbering, ctx, warnings));
    } else if (tag === 'w:tbl') {
      currentBlocks.push(parseTable(node, ctx, warnings));
    } else if (tag === 'w:sectPr') {
      // The trailing sectPr for the document's last section.
      await flushSection(node);
    }
    // Bookmarks and other body-level markers carry no block content — skipped.
  }

  // A document with content but no trailing sectPr (malformed, but don't
  // silently drop the content) still gets a section with sane defaults.
  if (currentBlocks.length > 0 || sections.length === 0) {
    sections.push({
      pageWidthPx: Math.round(12240 * TWIPS_TO_PX),
      pageHeightPx: Math.round(15840 * TWIPS_TO_PX),
      marginTopPx: 96,
      marginRightPx: 96,
      marginBottomPx: 96,
      marginLeftPx: 96,
      orientation: 'portrait',
      blocks: currentBlocks,
    });
  }

  const coreXml = await readZipText(zip, 'docProps/core.xml');
  const title = coreXml ? parseTitle(coreXml) : '';

  return { title, sections, warnings };
}

async function loadHeaderFooter(
  zip: JSZip,
  rels: Map<string, string>,
  rId: string,
  warnings: string[]
): Promise<DocHeaderFooter | undefined> {
  const target = rels.get(rId);
  if (!target) {
    warnings.push(`Header/footer relationship "${rId}" could not be resolved.`);
    return undefined;
  }
  const partXml = await readZipText(zip, `word/${target}`);
  if (!partXml) return undefined;
  const partRelsXml = await readZipText(zip, `word/_rels/${target}.rels`);
  const partRels = parseRels(partRelsXml);
  const partMedia = await buildMediaMap(zip, partRels);
  return parseHeaderFooter(partXml, { rels: partRels, media: partMedia }, warnings);
}

function parseTitle(coreXml: string): string {
  const root = parseXml(coreXml);
  const coreProps = find(root, 'cp:coreProperties');
  if (!coreProps) return '';
  const titleEl = find(childrenOf(coreProps), 'dc:title');
  return titleEl ? textOf(titleEl) : '';
}
