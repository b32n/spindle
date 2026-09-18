import { exportDocx } from '@b32nio/spindle-shared/export/docx';
import type {
  DocBlock,
  DocHeaderFooter,
  DocHeaderFooterParagraph,
  DocInline,
  DocParagraphStyle,
  DocRunBase,
  DocSection,
  DocumentExportInput,
} from '@b32nio/spindle-shared/export/docx';
import type {
  Block,
  DocumentData,
  DynamicFieldRun,
  HeaderFooterContent,
  HeaderFooterInlineContent,
  HeaderFooterParagraph,
  InlineContent,
  LineSpacing,
  ParagraphStyle,
  SectionData,
  TextRun,
  TextStyle,
} from '../types';
import type { DocumentImpl } from '../document';

export async function exportToDocx(document: DocumentImpl): Promise<{ bytes: Uint8Array; warnings: string[] }> {
  const data = document.getData();
  return exportDocx(toDocumentExportInput(data));
}

function toDocumentExportInput(data: DocumentData): DocumentExportInput {
  const textStylePool = data.textStylePool;
  const paragraphStylePool = data.paragraphStylePool;
  return {
    title: data.title,
    sections: data.sections.map((s) => toSection(s, data.title, textStylePool, paragraphStylePool)),
  };
}

function toSection(
  section: SectionData,
  title: string,
  textStylePool: Record<string, TextStyle>,
  paragraphStylePool: Record<string, ParagraphStyle>
): DocSection {
  const { size, margins, orientation } = section.pageConfig;
  // PageConfig.size is always the natural (portrait) size; landscape swaps
  // width/height for rendering, same convention DocumentEditor uses.
  const pageWidthPx = orientation === 'landscape' ? size.h : size.w;
  const pageHeightPx = orientation === 'landscape' ? size.w : size.h;

  return {
    pageWidthPx,
    pageHeightPx,
    marginTopPx: margins.top,
    marginRightPx: margins.right,
    marginBottomPx: margins.bottom,
    marginLeftPx: margins.left,
    orientation,
    blocks: section.blocks.map((b) => toBlock(b, title, textStylePool, paragraphStylePool)),
    header: section.header ? toHeaderFooter(section.header, title) : undefined,
    footer: section.footer ? toHeaderFooter(section.footer, title) : undefined,
  };
}

function toBlock(
  block: Block,
  title: string,
  textStylePool: Record<string, TextStyle>,
  paragraphStylePool: Record<string, ParagraphStyle>
): DocBlock {
  switch (block.type) {
    case 'paragraph':
      return {
        type: 'paragraph',
        content: toInlineContent(block.content, title, textStylePool),
        ...resolveParagraphStyle(block, paragraphStylePool),
      };
    case 'heading':
      return {
        type: 'heading',
        level: block.level,
        content: toInlineContent(block.content, title, textStylePool),
        ...resolveParagraphStyle(block, paragraphStylePool),
      };
    case 'list-item':
      return {
        type: 'listItem',
        listType: block.listType,
        level: block.level,
        content: toInlineContent(block.content, title, textStylePool),
      };
    case 'table':
      return {
        type: 'table',
        rows: block.rows.map((row) => ({
          cells: row.cells.map((cell) => ({
            content: toInlineContent(cell.content, title, textStylePool),
            colspan: cell.colspan,
            rowspan: cell.rowspan,
          })),
        })),
        colWidthsPx: block.colWidths,
      };
    case 'image':
      return { type: 'image', src: block.src, width: block.width, height: block.height, alt: block.alt, alignment: block.alignment };
    case 'horizontal-rule':
      return { type: 'horizontalRule' };
    case 'page-break':
      return { type: 'pageBreak' };
  }
}

/** px (96dpi, this codebase's geometry convention throughout) -> points. */
function pxToPt(px: number): number {
  return px * 0.75;
}

const LINE_SPACING_MAP: Partial<Record<LineSpacing['type'], DocParagraphStyle['lineSpacing']>> = {
  single: 'single',
  onePointFive: 'onePointFive',
  double: 'double',
  // atLeast/exactly/multiple carry a numeric value our generic type doesn't
  // model — fall back to single rather than guess at a multiplier.
};

interface ParagraphLikeBlock {
  styleId?: string;
  alignment?: ParagraphStyle['alignment'];
  indent?: number;
  lineSpacing?: LineSpacing;
  spaceBefore?: number;
  spaceAfter?: number;
}

/** Pool style as a base, with the block's own inline properties (when present) overriding it. */
function resolveParagraphStyle(block: ParagraphLikeBlock, pool: Record<string, ParagraphStyle>): DocParagraphStyle {
  const base = block.styleId ? pool[block.styleId] : undefined;
  const alignment = block.alignment ?? base?.alignment;
  const lineSpacingType = (block.lineSpacing ?? base?.lineSpacing)?.type;
  const spaceBefore = block.spaceBefore ?? base?.spaceBefore;
  const spaceAfter = block.spaceAfter ?? base?.spaceAfter;
  const leftIndent = block.indent ?? base?.leftIndent;

  return {
    alignment,
    spaceBeforePt: spaceBefore != null ? pxToPt(spaceBefore) : undefined,
    spaceAfterPt: spaceAfter != null ? pxToPt(spaceAfter) : undefined,
    leftIndentPt: leftIndent != null ? pxToPt(leftIndent) : undefined,
    firstLineIndentPt: base?.firstLineIndent != null ? pxToPt(base.firstLineIndent) : undefined,
    lineSpacing: lineSpacingType ? (LINE_SPACING_MAP[lineSpacingType] ?? 'single') : undefined,
  };
}

function toInlineContent(content: InlineContent[], title: string, pool: Record<string, TextStyle>): DocInline[] {
  return content.map((item) => toInline(item, title, pool)).filter((i): i is DocInline => i != null);
}

function toInline(item: InlineContent, title: string, pool: Record<string, TextStyle>): DocInline | null {
  switch (item.type) {
    case 'text':
      return { type: 'text', text: item.text, ...resolveTextStyle(item, pool) };
    case 'link':
      return {
        type: 'link',
        text: item.text,
        href: item.href,
        ...(item.styleId && pool[item.styleId] ? toDocRunBase(pool[item.styleId]) : {}),
      };
    case 'image':
      return { type: 'image', src: item.src, width: item.width, height: item.height, alt: item.alt };
    case 'dynamicField':
      return toDynamicFieldRun(item, title);
  }
}

/** Either/or, matching how the ProseMirror sync layer resolves the same field (sync.ts's inlineContentToPm): styleId is authoritative when present, inline properties only apply otherwise. */
function resolveTextStyle(item: TextRun, pool: Record<string, TextStyle>): DocRunBase {
  if (item.styleId && pool[item.styleId]) {
    return toDocRunBase(pool[item.styleId]);
  }
  return {
    bold: item.bold,
    italic: item.italic,
    underline: item.underline,
    strikethrough: item.strikethrough,
    superscript: item.superscript,
    subscript: item.subscript,
    smallCaps: item.smallCaps,
    fontFamily: item.fontFamily,
    fontSize: item.fontSize,
    color: item.color,
    backgroundColor: item.backgroundColor,
  };
}

function toDocRunBase(style: TextStyle): DocRunBase {
  return {
    bold: style.bold,
    italic: style.italic,
    underline: style.underline,
    strikethrough: style.strikethrough,
    superscript: style.superscript,
    subscript: style.subscript,
    smallCaps: style.smallCaps,
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    color: style.color,
    backgroundColor: style.backgroundColor,
  };
}

/**
 * pageNumber/totalPages become live docx fields. date/time/title don't have
 * a docx-native dynamic equivalent tied to our data model, so they're
 * substituted with a literal snapshot at export time rather than dropped.
 */
function toDynamicFieldRun(item: DynamicFieldRun, title: string): DocInline | null {
  switch (item.fieldType) {
    case 'pageNumber':
      return { type: 'pageNumber' };
    case 'totalPages':
      return { type: 'totalPages' };
    case 'title':
      return title ? { type: 'text', text: title } : null;
    case 'date':
      return { type: 'text', text: new Date().toLocaleDateString() };
    case 'time':
      return { type: 'text', text: new Date().toLocaleTimeString() };
    default:
      return null;
  }
}

function toHeaderFooter(hf: HeaderFooterContent, title: string): DocHeaderFooter {
  return {
    paragraphs: hf.blocks.map((p) => toHeaderFooterParagraph(p, title)),
    differentFirstPage: hf.differentFirstPage,
    firstPageParagraphs: hf.firstPageBlocks?.map((p) => toHeaderFooterParagraph(p, title)),
  };
}

function toHeaderFooterParagraph(p: HeaderFooterParagraph, title: string): DocHeaderFooterParagraph {
  return {
    alignment: p.alignment,
    content: p.content.map((item) => toHeaderFooterInline(item, title)).filter((i): i is DocInline => i != null),
  };
}

function toHeaderFooterInline(item: HeaderFooterInlineContent, title: string): DocInline | null {
  switch (item.type) {
    case 'text':
      return {
        type: 'text',
        text: item.text,
        bold: item.bold,
        italic: item.italic,
        underline: item.underline,
        fontSize: item.fontSize,
        fontFamily: item.fontFamily,
        color: item.color,
      };
    case 'image':
      return { type: 'image', src: item.src, width: item.width ?? 100, height: item.height ?? 100, alt: item.alt };
    case 'dynamicField':
      return toDynamicFieldRun(item, title);
  }
}
