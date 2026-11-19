import JSZip from 'jszip';
import type {
  DocBlock,
  DocHeaderFooter,
  DocHeaderFooterParagraph,
  DocInline,
  DocParagraphStyle,
  DocSection,
  DocTableRow,
  DocumentImportResult,
} from '@b32nio/spindle-shared/export/docx';

import { createCommentTracker, parseComments, type CommentTracker } from './comments';
import { drawingKind, type DrawingKind } from './drawing-kind';
import { ListCounter, listTypeFor, parseNumbering, type NumberingInfo } from './numbering';
import { parsePPr, StyleSheet, type ParagraphPartial } from './styles';
import { parseTheme } from './theme';
import { attrsOf, childrenOf, find, findAll, findDeep, parseXml, tagOf, type XmlNode } from './xml-tree';

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
  styles: StyleSheet;
  /** Follows list numbering through the body (absent for headers/footers). */
  lists?: ListCounter;
  /** Comment ranges open at the current position of the body walk (absent for headers/footers). */
  comments?: CommentTracker;
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
// Runs and inline content
// ---------------------------------------------------------------------------

const NON_BREAKING_HYPHEN = '\u2011';
const SOFT_HYPHEN = '\u00ad';

/** Content controls and custom XML wrap block or inline content: replace each wrapper by what it contains (its properties are dropped). */
function unwrapContainer(node: XmlNode): XmlNode[] {
  const tag = tagOf(node);
  if (tag === 'w:sdt') {
    const content = find(childrenOf(node), 'w:sdtContent');
    return content ? childrenOf(content) : [];
  }
  return childrenOf(node).filter((k) => tagOf(k) !== 'w:customXmlPr');
}

/** Block-level children with w:sdt / w:customXml wrappers replaced by their content, at any depth. */
function flattenBlocks(nodes: XmlNode[]): XmlNode[] {
  return nodes.flatMap((n) => {
    const tag = tagOf(n);
    return tag === 'w:sdt' || tag === 'w:customXml' ? flattenBlocks(unwrapContainer(n)) : [n];
  });
}

/** Paragraph-level children with w:sdt, w:customXml, w:smartTag and w:ins (tracked insertion) wrappers replaced by their content, at any depth. */
function flattenInline(nodes: XmlNode[]): XmlNode[] {
  return nodes.flatMap((n) => {
    const tag = tagOf(n);
    if (tag === 'w:sdt' || tag === 'w:customXml') return flattenInline(unwrapContainer(n));
    if (tag === 'w:smartTag' || tag === 'w:ins') return flattenInline(childrenOf(n).filter((k) => tagOf(k) !== 'w:smartTagPr'));
    return [n];
  });
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
  const pStyleId = paragraphStyleId(pNode);
  ctx.comments?.paragraphStart();
  let pendingField = ''; // accumulated instrText while inside a begin/.../end sequence
  let inField = false;

  const consumeRun = (runNode: XmlNode, target: string | undefined) => {
    const props = ctx.styles.run(find(childrenOf(runNode), 'w:rPr'), pStyleId);
    let text = '';
    const flush = () => {
      if (!text) return;
      const commentThreadId = ctx.comments?.anchorText(text);
      const runProps = commentThreadId ? { ...props, commentThreadId } : props;
      if (target) out.push({ type: 'link', text, href: target, ...runProps });
      else out.push({ type: 'text', text, ...runProps });
      text = '';
    };
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
      } else if (pieceTag === 'w:tab' || pieceTag === 'w:ptab') {
        if (!inField) text += '\t';
      } else if (pieceTag === 'w:br') {
        // A page break is a structural block (see pageBreakPlacement); every other break is a line break.
        if (!inField && attrsOf(piece)['@_w:type'] !== 'page') text += '\n';
      } else if (pieceTag === 'w:cr') {
        if (!inField) text += '\n';
      } else if (pieceTag === 'w:noBreakHyphen') {
        if (!inField) text += NON_BREAKING_HYPHEN;
      } else if (pieceTag === 'w:softHyphen') {
        if (!inField) text += SOFT_HYPHEN;
      } else if (pieceTag === 'w:ruby') {
        // Keep the base text; the reading marks (w:rt) are not imported.
        flush();
        const base = find(childrenOf(piece), 'w:rubyBase');
        if (base) consumeRuns(childrenOf(base), target);
      } else if (pieceTag === 'w:drawing') {
        const image = parseDrawing(piece, ctx.media, warnings);
        if (image) out.push(image);
      } else if (pieceTag === 'mc:AlternateContent') {
        // LibreOffice (and Word for some shapes) wrap the drawing in
        // <mc:AlternateContent><mc:Choice><w:drawing>; the Fallback is legacy VML.
        const drawing = findDeep(piece, 'w:drawing');
        const image = drawing && parseDrawing(drawing, ctx.media, warnings);
        if (image) out.push(image);
      }
    }
    flush();
  };

  const consumeRuns = (runs: XmlNode[], target: string | undefined) => {
    for (const child of flattenInline(runs)) {
      const tag = tagOf(child);
      if (tag === 'w:fldSimple') out.push(...fieldRunFor(attrsOf(child)['@_w:instr'] ?? '', warnings));
      else if (tag === 'w:r') consumeRun(child, target);
      else ctx.comments?.mark(child);
    }
  };

  for (const child of flattenInline(childrenOf(pNode))) {
    const tag = tagOf(child);
    if (tag === 'w:hyperlink') {
      const target = resolveHyperlinkTarget(child, ctx.rels, warnings);
      consumeRuns(childrenOf(child), target);
    } else if (tag === 'w:r' || tag === 'w:fldSimple') {
      consumeRuns([child], undefined);
    } else {
      ctx.comments?.mark(child);
    }
    // Bookmarks, proofErr, and other structural markers carry no content — skipped.
  }

  return out;
}

const EMU_TO_PX = 1 / 9525; // 914400 EMU/inch, 96px/inch

const DRAWING_KIND_WARNINGS: Record<Exclude<DrawingKind, 'picture'>, string> = {
  chart: 'A chart was not imported — charts are not supported yet.',
  diagram: 'A SmartArt diagram was not imported — diagrams are not supported yet.',
  shape: 'A shape or text box was not imported — shapes and text boxes are not supported yet.',
  embeddedObject: 'An embedded object preview was not imported — embedded objects are not supported yet.',
  unknown: 'A drawing object of an unsupported kind was not imported.',
};

/**
 * A <w:drawing> run child. Only inline pictures (<wp:inline> + <a:blip r:embed>) become images. Anchored
 * pictures aren't positioned the way our block model supports, so they're dropped with a warning; every
 * other kind (chart, SmartArt, shape/text box, embedded-object preview) is reported as what it is rather
 * than as a missing image. Linked pictures (<a:blip r:link>) point outside the package and are never fetched.
 */
function parseDrawing(drawingNode: XmlNode, media: Map<string, string>, warnings: string[]): DocInline | null {
  const kind = drawingKind(drawingNode);
  if (kind !== 'picture') {
    warnings.push(DRAWING_KIND_WARNINGS[kind]);
    return null;
  }
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
  const blipAttrs = blip ? attrsOf(blip) : {};
  const embedId = blipAttrs['@_r:embed'];
  const src = embedId ? media.get(embedId) : undefined;
  if (!src) {
    if (!embedId && blipAttrs['@_r:link']) {
      warnings.push('A linked picture was skipped — images linked from outside the document are never fetched.');
    } else {
      warnings.push('An embedded image could not be resolved and was skipped.');
    }
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

/** The id of a paragraph's own style (`w:pStyle`), if it names one. */
function paragraphStyleId(pNode: XmlNode): string | undefined {
  const pPr = find(childrenOf(pNode), 'w:pPr');
  const el = pPr ? find(childrenOf(pPr), 'w:pStyle') : undefined;
  return el ? attrsOf(el)['@_w:val'] : undefined;
}

function hasPageBreakRun(pNode: XmlNode): boolean {
  return childrenOf(pNode).some((child) => {
    if (tagOf(child) !== 'w:r') return false;
    const br = find(childrenOf(child), 'w:br');
    return !!br && attrsOf(br)['@_w:type'] === 'page';
  });
}

/**
 * Where a page-break run sits relative to the paragraph's other content:
 * 'before' (nothing but empty runs precede it), 'after' otherwise, or null when
 * the paragraph has no page-break run. LibreOffice stores the break after the
 * content it follows (`<w:r><w:br w:type="page"/></w:r>` at the end of the
 * previous paragraph); Word's "page break before" style puts it first.
 */
function pageBreakPlacement(pNode: XmlNode): 'before' | 'after' | null {
  let sawContent = false;
  for (const child of childrenOf(pNode)) {
    const tag = tagOf(child);
    if (tag === 'w:r') {
      const runKids = childrenOf(child);
      const br = find(runKids, 'w:br');
      if (br && attrsOf(br)['@_w:type'] === 'page') return sawContent ? 'after' : 'before';
      if (runKids.some((k) => ['w:t', 'w:drawing', 'w:pict', 'w:tab'].includes(tagOf(k) ?? ''))) sawContent = true;
    } else if (tag === 'w:hyperlink' || tag === 'w:ins' || tag === 'w:sdt' || tag === 'w:smartTag') {
      sawContent = true;
    }
  }
  return null;
}

/**
 * Each numbered list item carries the number it shows (`start`), but the model keeps `start` only where a list
 * begins at something other than 1 or restarts: on the first item of a run of numbered items, and on a later
 * top-level item whose number is not simply the next one (a new list instance). Elsewhere the number follows from
 * the order.
 */
function settleListStarts(blocks: DocBlock[]): DocBlock[] {
  let inRun = false;
  let lastTopLevel = 0;
  for (const block of blocks) {
    if (block.type !== 'listItem' || block.listType !== 'numbered') {
      inRun = false;
      continue;
    }
    if (block.level === 0 && block.start !== undefined) {
      const n = block.start;
      const keep = inRun ? n !== lastTopLevel + 1 : n !== 1;
      lastTopLevel = n;
      if (!keep) delete block.start;
    } else {
      delete block.start;
    }
    inRun = true;
  }
  return blocks;
}

const LINE_KEYWORDS: Array<[number, 'single' | 'onePointFive' | 'double']> = [
  [1, 'single'],
  [1.5, 'onePointFive'],
  [2, 'double'],
];

/** The paragraph properties the interchange format carries, from a paragraph's effective formatting. */
function toDocParagraphStyle(effective: ParagraphPartial): DocParagraphStyle {
  const style: DocParagraphStyle = {
    ...(effective.alignment && { alignment: effective.alignment }),
    ...(effective.spaceBeforePt != null && { spaceBeforePt: effective.spaceBeforePt }),
    ...(effective.spaceAfterPt != null && { spaceAfterPt: effective.spaceAfterPt }),
    ...(effective.leftIndentPt != null && { leftIndentPt: effective.leftIndentPt }),
    ...(effective.rightIndentPt != null && { rightIndentPt: effective.rightIndentPt }),
    ...(effective.firstLineIndentPt != null && { firstLineIndentPt: effective.firstLineIndentPt }),
    ...(effective.keepNext && { keepWithNext: true }),
    ...(effective.keepLines && { keepTogether: true }),
    ...(effective.pageBreakBefore && { pageBreakBefore: true }),
  };
  const spacing = effective.lineSpacing;
  if (spacing?.rule === 'auto') {
    const keyword = LINE_KEYWORDS.find(([factor]) => Math.abs(factor - spacing.value) < 0.001);
    if (keyword) style.lineSpacing = keyword[1];
    else style.lineSpacingCustom = { rule: 'multiple', value: Math.round(spacing.value * 1000) / 1000 };
  } else if (spacing) {
    style.lineSpacingCustom = { rule: spacing.rule === 'exact' ? 'exactly' : 'atLeast', value: spacing.value };
  }
  return style;
}

/** parseParagraph plus the page break it may carry alongside its content. */
function parseParagraphBlocks(
  pNode: XmlNode,
  numbering: NumberingInfo,
  ctx: PartContext,
  warnings: string[]
): DocBlock[] {
  const block = parseParagraph(pNode, numbering, ctx, warnings);
  if (block.type === 'pageBreak') return [block];
  const placement = pageBreakPlacement(pNode);
  if (!placement) return [block];
  return placement === 'before' ? [{ type: 'pageBreak' }, block] : [block, { type: 'pageBreak' }];
}

function hasBottomBorder(pPrKids: XmlNode[]): boolean {
  const pBdr = find(pPrKids, 'w:pBdr');
  return !!pBdr && !!find(childrenOf(pBdr), 'w:bottom');
}

function parseParagraph(
  pNode: XmlNode,
  numbering: NumberingInfo,
  ctx: PartContext,
  warnings: string[]
): DocBlock {
  const kids = childrenOf(pNode);
  const pPr = find(kids, 'w:pPr');
  const pPrKids = pPr ? childrenOf(pPr) : [];
  const content = paragraphInlineContent(pNode, ctx, warnings);
  // Effective formatting: docDefaults, then the paragraph's style chain, then its own properties.
  const direct = parsePPr(pPr);
  const effective = ctx.styles.paragraph(paragraphStyleId(pNode), direct);
  const style = toDocParagraphStyle(effective);

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

  // A list item is a paragraph numbered directly, or through its style. A heading that is numbered through its
  // style (numbered headings) stays a heading.
  const level = effective.ilvl ?? 0;
  const listItem = (numId: string | undefined): DocBlock => {
    const listType = numId ? listTypeFor(numbering, numId, level) : 'bullet';
    // The number the item shows; importDocx keeps it as `start` only where a list begins or restarts.
    const shown = numId && listType === 'numbered' ? ctx.lists?.next(numId, level) : undefined;
    return { type: 'listItem', listType, level, content, ...(shown !== undefined && { start: shown }), ...style };
  };
  if (find(pPrKids, 'w:numPr') && direct.numId !== '0') return listItem(direct.numId);
  if (effective.headingLevel) return { type: 'heading', level: effective.headingLevel as 1 | 2 | 3 | 4 | 5 | 6, content, ...style };
  if (effective.numId && effective.numId !== '0') return listItem(effective.numId);

  return { type: 'paragraph', content, ...style };
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

function parseTable(tblNode: XmlNode, ctx: PartContext, warnings: string[]): DocBlock {
  const rows: DocTableRow[] = [];
  // Column index -> the DocTableCell object currently accumulating a
  // vertical merge, so a later row's continuation cell can bump its rowspan
  // instead of emitting a new cell.
  const openVMerge = new Map<number, { cell: DocTableRow['cells'][number] }>();

  for (const trChild of flattenBlocks(childrenOf(tblNode))) {
    if (tagOf(trChild) !== 'w:tr') {
      ctx.comments?.mark(trChild);
      continue;
    }
    const trNode = trChild;
    const rowCells: DocTableRow['cells'] = [];
    let col = 0;
    for (const tcNode of flattenBlocks(childrenOf(trNode))) {
      if (tagOf(tcNode) !== 'w:tc') {
        ctx.comments?.mark(tcNode);
        continue;
      }
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
        ctx.comments?.mark(tcNode);
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
  let result: DocInline[] | undefined;
  for (const child of flattenBlocks(childrenOf(tcNode))) {
    if (!result && tagOf(child) === 'w:p') {
      const content = paragraphInlineContent(child, ctx, warnings);
      if (content.length > 0) result = content;
    } else {
      // Content that is not imported can still open or close a comment range.
      ctx.comments?.mark(child);
    }
  }
  return result ?? [];
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
  const rawWidthTwips = Number(pgSzAttrs['@_w:w'] ?? 12240);
  const rawHeightTwips = Number(pgSzAttrs['@_w:h'] ?? 15840);
  // Some producers write landscape pages with w < h (the natural size) and
  // rely on w:orient; others omit w:orient and only swap w/h. Treat either
  // signal as landscape and report the on-page size (width >= height).
  const landscape = pgSzAttrs['@_w:orient'] === 'landscape' || rawWidthTwips > rawHeightTwips;
  const orientation: 'portrait' | 'landscape' = landscape ? 'landscape' : 'portrait';
  const widthTwips = landscape ? Math.max(rawWidthTwips, rawHeightTwips) : rawWidthTwips;
  const heightTwips = landscape ? Math.min(rawWidthTwips, rawHeightTwips) : rawHeightTwips;

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
    for (const p of findAll(flattenBlocks(childrenOf(rootEl)), 'w:p')) {
      const pPr = find(childrenOf(p), 'w:pPr');
      const alignment = ctx.styles.paragraph(paragraphStyleId(p), parsePPr(pPr)).alignment;
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

  const [documentXml, docRelsXml, stylesXml, numberingXml, commentsXml, commentsExtendedXml, peopleXml] = await Promise.all([
    readZipText(zip, 'word/document.xml'),
    readZipText(zip, 'word/_rels/document.xml.rels'),
    readZipText(zip, 'word/styles.xml'),
    readZipText(zip, 'word/numbering.xml'),
    readZipText(zip, 'word/comments.xml'),
    readZipText(zip, 'word/commentsExtended.xml'),
    readZipText(zip, 'word/people.xml'),
  ]);

  if (!documentXml) {
    return { title: '', sections: [], warnings: ['word/document.xml is missing — this file may not be a valid .docx.'] };
  }

  const rels = parseRels(docRelsXml);
  const media = await buildMediaMap(zip, rels);
  const parsedComments = parseComments({ comments: commentsXml, extended: commentsExtendedXml, people: peopleXml }, warnings);
  const comments = createCommentTracker(parsedComments, documentXml);
  const themeXml = await readZipText(zip, 'word/theme/theme1.xml');
  const styles = new StyleSheet(stylesXml, parseTheme(themeXml));
  const numbering = parseNumbering(numberingXml);
  const ctx: PartContext = { rels, media, styles, lists: new ListCounter(numbering), comments };

  const root = parseXml(documentXml);
  const documentEl = find(root, 'w:document');
  const bodyEl = documentEl ? find(childrenOf(documentEl), 'w:body') : undefined;
  const bodyChildren = bodyEl ? flattenBlocks(childrenOf(bodyEl)) : [];

  const sections: DocSection[] = [];
  let currentBlocks: DocBlock[] = [];

  const flushSection = async (sectPrNode: XmlNode) => {
    const parsed = parseSectPr(sectPrNode);
    const header = parsed.headerRId ? await loadHeaderFooter(zip, rels, parsed.headerRId, styles, warnings) : undefined;
    const footer = parsed.footerRId ? await loadHeaderFooter(zip, rels, parsed.footerRId, styles, warnings) : undefined;
    sections.push({
      pageWidthPx: parsed.pageWidthPx,
      pageHeightPx: parsed.pageHeightPx,
      marginTopPx: parsed.marginTopPx,
      marginRightPx: parsed.marginRightPx,
      marginBottomPx: parsed.marginBottomPx,
      marginLeftPx: parsed.marginLeftPx,
      orientation: parsed.orientation,
      blocks: settleListStarts(currentBlocks),
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
        // A sectPr nested in a paragraph's pPr ends that section, and per
        // OOXML the paragraph is the section's last and keeps its content
        // (Word and docx >= 9.12 emit it this way; older docx used an empty
        // carrier paragraph).
        currentBlocks.push(...parseParagraphBlocks(node, numbering, ctx, warnings));
        await flushSection(inlineSectPr);
        continue;
      }
      currentBlocks.push(...parseParagraphBlocks(node, numbering, ctx, warnings));
    } else if (tag === 'w:tbl') {
      currentBlocks.push(parseTable(node, ctx, warnings));
    } else if (tag === 'w:sectPr') {
      // The trailing sectPr for the document's last section.
      await flushSection(node);
    } else {
      ctx.comments?.mark(node);
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
      blocks: settleListStarts(currentBlocks),
    });
  }

  const coreXml = await readZipText(zip, 'docProps/core.xml');
  const title = coreXml ? parseTitle(coreXml) : '';

  const threads = comments.threads();
  return { title, sections, warnings, ...(threads.length > 0 && { comments: threads }) };
}

async function loadHeaderFooter(
  zip: JSZip,
  rels: Map<string, string>,
  rId: string,
  styles: StyleSheet,
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
  return parseHeaderFooter(partXml, { rels: partRels, media: partMedia, styles }, warnings);
}

function parseTitle(coreXml: string): string {
  const root = parseXml(coreXml);
  const coreProps = find(root, 'cp:coreProperties');
  if (!coreProps) return '';
  const titleEl = find(childrenOf(coreProps), 'dc:title');
  return titleEl ? textOf(titleEl) : '';
}
