import type { ImportLossReport } from '@b32nio/spindle-shared/export/docx';

// The DOCX reader/writer lives in the separate @b32nio/spindle-io package; load it on first use so
// hosts that never open or save the format don't pay for its dependencies.
const loadIo = () => import('@b32nio/spindle-io/docx');
import type {
  DocBlock,
  DocColumns,
  DocCommentThread,
  DocHeaderFooter,
  DocHeaderFooterParagraph,
  DocImageRun,
  DocInline,
  DocNamedStyle,
  DocParagraphStyle,
  DocRunBase,
  DocSection,
  DocTableStyle,
  DocTableStylePart,
  DocTextRun,
  DocumentExportInput,
  DocumentImportResult,
} from '@b32nio/spindle-shared/export/docx';
import type {
  Block,
  DocumentData,
  DynamicFieldRun,
  HeaderFooterContent,
  HeaderFooterInlineContent,
  HeaderFooterParagraph,
  InlineContent,
  InlineImage,
  LineSpacing,
  NamedStyle,
  PageColumns,
  ParagraphFormat,
  ParagraphStyle,
  SectionData,
  TableStyleData,
  TableStylePart,
  TextRun,
  TextStyle,
} from '../types';
import type { DocsCommentThread } from '../comments';
import type { DocumentImpl } from '../document';
import { createSection } from '../document';
import { TextStylePoolImpl } from '../style-pool';
import { generateBlockId, generateCellId, generateRowId } from '../blocks/utils';

/** A comment with no recorded date still needs a timestamp (the model requires one); the epoch marks it unknown. */
const UNKNOWN_COMMENT_DATE = new Date(0).toISOString();

export async function exportToDocx(document: DocumentImpl): Promise<{ bytes: Uint8Array; warnings: string[] }> {
  const data = document.getData();
  const { exportDocx } = await loadIo();
  return exportDocx(toDocumentExportInput(data));
}

function toDocumentExportInput(data: DocumentData): DocumentExportInput {
  const textStylePool = data.textStylePool;
  const paragraphStylePool = data.paragraphStylePool;
  return {
    title: data.title,
    sections: data.sections.map((s) => toSection(s, data.title, textStylePool, paragraphStylePool)),
    ...(data.threads && data.threads.length > 0 && { comments: data.threads.map(toThreadForDocx) }),
    ...(data.namedStyles && Object.keys(data.namedStyles).length > 0 && { namedStyles: toDocNamedStyles(data.namedStyles) }),
  };
}

function toDocNamedStyles(styles: Record<string, NamedStyle>): Record<string, DocNamedStyle> {
  const out: Record<string, DocNamedStyle> = {};
  for (const [id, style] of Object.entries(styles)) {
    const text = style.text ? defined(toDocRunBase(style.text)) : undefined;
    const paragraph = style.paragraph ? defined(resolveParagraphStyle(style.paragraph, {})) : undefined;
    out[id] = {
      name: style.name,
      kind: style.kind,
      ...(style.basedOn && { basedOn: style.basedOn }),
      ...(style.next && { next: style.next }),
      ...(style.link && { link: style.link }),
      ...(style.isDefault && { isDefault: true }),
      ...(style.outlineLevel !== undefined && { outlineLevel: style.outlineLevel }),
      ...(text && { text }),
      ...(paragraph && { paragraph }),
      ...(style.table && { table: toDocTableStyle(style.table) }),
    };
  }
  return out;
}

function toDocTableStyle(table: TableStyleData): DocTableStyle {
  const part = (p: TableStylePart): DocTableStylePart => {
    const text = p.text ? defined(toDocRunBase(p.text)) : undefined;
    return { ...(text && Object.keys(text).length > 0 && { text }), ...(p.backgroundColor && { backgroundColor: p.backgroundColor }), ...(p.borders && { borders: p.borders }), ...(p.verticalAlign && { verticalAlign: p.verticalAlign }) };
  };
  return {
    ...(table.borders && { borders: table.borders }),
    ...(table.whole && { whole: part(table.whole) }),
    ...(table.conditions && { conditions: Object.fromEntries(Object.entries(table.conditions).map(([k, v]) => [k, part(v)])) }),
    ...(table.rowBandSize && { rowBandSize: table.rowBandSize }),
    ...(table.colBandSize && { colBandSize: table.colBandSize }),
  };
}

function toTableStyleData(table: DocTableStyle): TableStyleData {
  const part = (p: DocTableStylePart): TableStylePart => ({
    ...(p.text && { text: defined(toTextStyle(p.text)) }),
    ...(p.backgroundColor && { backgroundColor: p.backgroundColor }),
    ...(p.borders && { borders: p.borders }),
    ...(p.verticalAlign && { verticalAlign: p.verticalAlign }),
  });
  return {
    ...(table.borders && { borders: table.borders }),
    ...(table.whole && { whole: part(table.whole) }),
    ...(table.conditions && { conditions: Object.fromEntries(Object.entries(table.conditions).map(([k, v]) => [k, part(v)])) }),
    ...(table.rowBandSize && { rowBandSize: table.rowBandSize }),
    ...(table.colBandSize && { colBandSize: table.colBandSize }),
  };
}

/** Drops `undefined` fields so the interchange object only carries what is set. */
function defined<T extends object>(value: T): T {
  return Object.fromEntries(Object.entries(value).filter(([, v]) => v !== undefined)) as T;
}

function toThreadForDocx(thread: DocsCommentThread): DocCommentThread {
  return {
    id: thread.id,
    quote: thread.anchor.quote,
    ...(thread.status === 'resolved' && { resolved: true }),
    comments: thread.comments.map((c) => ({
      id: c.id,
      author: c.authorName,
      authorId: c.authorId,
      ...(c.authorInitials && { initials: c.authorInitials }),
      ...(c.createdAt !== UNKNOWN_COMMENT_DATE && { date: c.createdAt }),
      text: c.body,
    })),
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
    ...(section.pageConfig.columns && { columns: toDocColumns(section.pageConfig.columns) }),
    ...(section.pageNumbering && { pageNumbering: section.pageNumbering }),
    ...(section.verticalAlign && { verticalAlign: section.verticalAlign }),
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
        ...(block.start !== undefined && { start: block.start }),
        content: toInlineContent(block.content, title, textStylePool),
        ...resolveParagraphStyle(block, paragraphStylePool),
      };
    case 'table':
      return {
        type: 'table',
        rows: block.rows.map((row) => ({
          ...(row.repeatHeader && { repeatHeader: true }),
          ...(row.cantSplit && { cantSplit: true }),
          cells: row.cells.map((cell) => ({
            content: toInlineContent(cell.content, title, textStylePool),
            ...(cell.blocks && { blocks: cell.blocks.map((b) => toBlock(b, title, textStylePool, paragraphStylePool)) }),
            colspan: cell.colspan,
            rowspan: cell.rowspan,
            ...(cell.backgroundColor && { backgroundColor: cell.backgroundColor }),
            ...(cell.borders && { borders: cell.borders }),
            ...(cell.verticalAlign && { verticalAlign: cell.verticalAlign }),
            ...(cell.textDirection && { textDirection: cell.textDirection }),
            ...(cell.colwidth && cell.colwidth.length > 0 && { colwidthPx: cell.colwidth }),
          })),
        })),
        colWidthsPx: block.colWidths,
        ...(block.borders && { borders: block.borders }),
        ...(block.look && { look: block.look }),
        ...(block.namedStyleId && { namedStyleId: block.namedStyleId }),
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

const LINE_SPACING_KEYWORDS: Partial<Record<LineSpacing['type'], DocParagraphStyle['lineSpacing']>> = {
  single: 'single',
  onePointFive: 'onePointFive',
  double: 'double',
};

/** Model line spacing -> the interchange keyword, or a custom rule (atLeast/exactly heights go px -> pt). */
function toDocLineSpacing(spacing: LineSpacing | undefined): Pick<DocParagraphStyle, 'lineSpacing' | 'lineSpacingCustom'> {
  if (!spacing) return {};
  const keyword = LINE_SPACING_KEYWORDS[spacing.type];
  if (keyword) return { lineSpacing: keyword };
  if (spacing.value == null) return { lineSpacing: 'single' };
  if (spacing.type === 'multiple') return { lineSpacingCustom: { rule: 'multiple', value: spacing.value } };
  return { lineSpacingCustom: { rule: spacing.type === 'exactly' ? 'exactly' : 'atLeast', value: pxToPt(spacing.value) } };
}

function fromDocLineSpacing(style: DocParagraphStyle): LineSpacing | undefined {
  const custom = style.lineSpacingCustom;
  if (custom) {
    if (custom.rule === 'multiple') return { type: 'multiple', value: custom.value };
    return { type: custom.rule, value: ptToPx(custom.value) };
  }
  return style.lineSpacing ? ({ type: style.lineSpacing } as LineSpacing) : undefined;
}

/** Pixels per indent level of the indent/outdent commands. */
const INDENT_LEVEL_PX = 24;

interface ParagraphLikeBlock extends ParagraphFormat {
  styleId?: string;
  /** Indent level (24 px each), on top of `leftIndent`. */
  indent?: number;
}

/** Pool style as a base, with the block's own inline properties (when present) overriding it. */
function resolveParagraphStyle(block: ParagraphLikeBlock, pool: Record<string, ParagraphStyle>): DocParagraphStyle {
  const base = block.styleId ? pool[block.styleId] : undefined;
  const alignment = block.alignment ?? base?.alignment;
  const spaceBefore = block.spaceBefore ?? base?.spaceBefore;
  const spaceAfter = block.spaceAfter ?? base?.spaceAfter;
  const leftIndent = block.leftIndent ?? base?.leftIndent;
  const rightIndent = block.rightIndent ?? base?.rightIndent;
  const firstLineIndent = block.firstLineIndent ?? base?.firstLineIndent;
  const indentPx = block.indent ? block.indent * INDENT_LEVEL_PX : 0;
  const totalLeft = leftIndent != null || indentPx ? (leftIndent ?? 0) + indentPx : undefined;
  const keepWithNext = block.keepWithNext ?? base?.keepWithNext;
  const keepTogether = block.keepTogether ?? base?.keepTogether;
  const pageBreakBefore = block.pageBreakBefore ?? base?.pageBreakBefore;

  return {
    alignment,
    spaceBeforePt: spaceBefore != null ? pxToPt(spaceBefore) : undefined,
    spaceAfterPt: spaceAfter != null ? pxToPt(spaceAfter) : undefined,
    leftIndentPt: totalLeft != null ? pxToPt(totalLeft) : undefined,
    rightIndentPt: rightIndent != null ? pxToPt(rightIndent) : undefined,
    firstLineIndentPt: firstLineIndent != null ? pxToPt(firstLineIndent) : undefined,
    ...toDocLineSpacing(block.lineSpacing ?? base?.lineSpacing),
    ...(keepWithNext && { keepWithNext: true }),
    ...(keepTogether && { keepTogether: true }),
    ...(pageBreakBefore && { pageBreakBefore: true }),
    ...(block.namedStyleId && { namedStyleId: block.namedStyleId }),
  };
}

function toInlineContent(content: InlineContent[], title: string, pool: Record<string, TextStyle>): DocInline[] {
  return content.map((item) => toInline(item, title, pool)).filter((i): i is DocInline => i != null);
}

function toInline(item: InlineContent, title: string, pool: Record<string, TextStyle>): DocInline | null {
  switch (item.type) {
    case 'text':
      return {
        type: 'text',
        text: item.text,
        ...resolveTextStyle(item, pool),
        ...(item.namedStyleId && { namedStyleId: item.namedStyleId }),
        ...(item.commentThreadId && { commentThreadId: item.commentThreadId }),
      };
    case 'link':
      return {
        type: 'link',
        text: item.text,
        href: item.href,
        ...(item.styleId && pool[item.styleId] ? toDocRunBase(pool[item.styleId]) : {}),
        ...(item.namedStyleId && { namedStyleId: item.namedStyleId }),
        ...(item.commentThreadId && { commentThreadId: item.commentThreadId }),
      };
    case 'image':
      return { type: 'image', src: item.src, width: item.width, height: item.height, alt: item.alt };
    case 'dynamicField':
      return toDynamicFieldRun(item, title);
    case 'field':
      return {
        type: 'field',
        instruction: item.instruction,
        kind: item.kind,
        result:
          typeof item.result === 'string'
            ? item.result
            : item.result.map((r): DocTextRun | DocImageRun =>
                r.type === 'text'
                  ? { type: 'text', text: r.text, ...toDocRunBase(r), ...(r.namedStyleId && { namedStyleId: r.namedStyleId }) }
                  : { type: 'image', src: r.src, width: r.width, height: r.height, alt: r.alt }
              ),
        ...(item.dirty && { dirty: true }),
      };
    case 'bookmark':
      return { type: 'bookmark', name: item.name };
  }
}

/** Either/or, matching how the ProseMirror sync layer resolves the same field (sync.ts's inlineContentToPm): styleId is authoritative when present, inline properties only apply otherwise. */
function resolveTextStyle(item: TextRun, pool: Record<string, TextStyle>): DocRunBase {
  if (item.styleId && pool[item.styleId]) {
    return toDocRunBase(pool[item.styleId]);
  }
  return toDocRunBase(item);
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
    allCaps: style.allCaps,
    hidden: style.hidden,
    fontFamily: style.fontFamily,
    fontFamilyEastAsia: style.fontFamilyEastAsia,
    fontFamilyComplex: style.fontFamilyComplex,
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

function toDocColumns(columns: PageColumns): DocColumns {
  return {
    count: columns.count,
    ...(columns.spacing !== undefined && { spacingPx: columns.spacing }),
    ...(columns.separator && { separator: true }),
    ...(columns.widths && { widthsPx: columns.widths }),
  };
}

function fromDocColumns(columns: DocColumns): PageColumns {
  return {
    count: columns.count,
    ...(columns.spacingPx !== undefined && { spacing: columns.spacingPx }),
    ...(columns.separator && { separator: true }),
    ...(columns.widthsPx && { widths: columns.widthsPx }),
  };
}

function toHeaderFooter(hf: HeaderFooterContent, title: string): DocHeaderFooter {
  return {
    paragraphs: hf.blocks.map((p) => toHeaderFooterParagraph(p, title)),
    differentFirstPage: hf.differentFirstPage,
    firstPageParagraphs: hf.firstPageBlocks?.map((p) => toHeaderFooterParagraph(p, title)),
    ...(hf.differentOddEven && { differentOddEven: true }),
    ...(hf.evenPageBlocks && { evenPageParagraphs: hf.evenPageBlocks.map((p) => toHeaderFooterParagraph(p, title)) }),
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

// ---------------------------------------------------------------------------
// Import
// ---------------------------------------------------------------------------

/**
 * Replace `document`'s entire contents with what's in the docx file at
 * `bytes` — every existing section is discarded, matching what "import a
 * file" means elsewhere (see sheets-core's importFromXlsx). The document's
 * own id/createdAt survive; everything else (title, sections, style pools)
 * is rebuilt from the parsed file.
 */
export async function importFromDocx(
  bytes: Uint8Array,
  document: DocumentImpl
): Promise<{ warnings: string[]; loss: ImportLossReport }> {
  const { detectDocxLoss, importDocx } = await loadIo();
  const loss = await detectDocxLoss(bytes);
  const result = await importDocx(bytes);
  // A zip that isn't a .docx imports as "no sections" plus a warning; applying that would wipe
  // the open document, so refuse it before touching the document.
  if (result.sections.length === 0) {
    throw new Error(result.warnings[0] ?? 'The file contains no document content.');
  }
  const data = toDocumentData(result, document);
  document.setData(data);
  return { warnings: result.warnings, loss };
}

/**
 * What importing this .docx would not bring across (tracked changes, footnotes, content controls,
 * floating images…). Read-only: nothing is imported or modified. `detail: 'full'` also lists formatting and
 * structure that is flattened or ignored (table formatting, tabs and breaks, sections…).
 */
export function inspectDocx(bytes: Uint8Array, options?: { detail?: 'summary' | 'full' }): Promise<ImportLossReport> {
  return loadIo().then((io) => io.detectDocxLoss(bytes, options));
}

function toDocumentData(result: DocumentImportResult, document: DocumentImpl): DocumentData {
  const existing = document.getData();
  const textStylePool = new TextStylePoolImpl();
  return {
    id: existing.id,
    title: result.title || existing.title,
    sections: result.sections.map((s) => toSectionData(s, textStylePool)),
    defaultPageConfig: existing.defaultPageConfig,
    textStylePool: textStylePool.toData(),
    paragraphStylePool: {},
    createdAt: existing.createdAt,
    updatedAt: new Date().toISOString(),
    ...(result.comments && result.comments.length > 0 && { threads: result.comments.map(toThreadFromDocx) }),
    ...(result.namedStyles && Object.keys(result.namedStyles).length > 0 && { namedStyles: toNamedStyles(result.namedStyles) }),
  };
}

function toNamedStyles(styles: Record<string, DocNamedStyle>): Record<string, NamedStyle> {
  const out: Record<string, NamedStyle> = {};
  for (const [id, style] of Object.entries(styles)) {
    out[id] = {
      name: style.name,
      kind: style.kind,
      ...(style.basedOn && { basedOn: style.basedOn }),
      ...(style.next && { next: style.next }),
      ...(style.link && { link: style.link }),
      ...(style.isDefault && { isDefault: true }),
      ...(style.outlineLevel !== undefined && { outlineLevel: style.outlineLevel }),
      ...(style.text && { text: defined(toTextStyle(style.text)) }),
      ...(style.paragraph && { paragraph: toParagraphFormat(style.paragraph) }),
      ...(style.table && { table: toTableStyleData(style.table) }),
    };
  }
  return out;
}

function toThreadFromDocx(thread: DocCommentThread): DocsCommentThread {
  const createdAt = thread.comments[0]?.date ?? UNKNOWN_COMMENT_DATE;
  return {
    id: thread.id,
    anchor: { quote: thread.quote },
    status: thread.resolved ? 'resolved' : 'open',
    createdAt,
    comments: thread.comments.map((c) => ({
      id: `docx-comment-${c.id}`,
      authorId: c.authorId ?? c.author,
      authorName: c.author,
      ...(c.initials && { authorInitials: c.initials }),
      body: c.text,
      createdAt: c.date ?? UNKNOWN_COMMENT_DATE,
    })),
  };
}

/** points -> px (96dpi), the inverse of docx.ts's own pxToPt. */
function ptToPx(pt: number): number {
  return pt / 0.75;
}

function toSectionData(
  section: DocSection,
  textStylePool: TextStylePoolImpl
): SectionData {
  // Inverse of toSection's own convention: PageConfig.size is always the
  // natural (portrait) size, with landscape swapping width/height back.
  const w = section.orientation === 'landscape' ? section.pageHeightPx : section.pageWidthPx;
  const h = section.orientation === 'landscape' ? section.pageWidthPx : section.pageHeightPx;
  const built = createSection(
    {
      size: { w, h },
      margins: {
        top: section.marginTopPx,
        right: section.marginRightPx,
        bottom: section.marginBottomPx,
        left: section.marginLeftPx,
      },
      orientation: section.orientation,
      ...(section.columns && { columns: fromDocColumns(section.columns) }),
    },
    section.blocks.map((b) => toBlockFromDocx(b, textStylePool))
  );
  return {
    ...built,
    header: section.header ? toHeaderFooterContentFromDocx(section.header) : undefined,
    footer: section.footer ? toHeaderFooterContentFromDocx(section.footer) : undefined,
    ...(section.pageNumbering && { pageNumbering: section.pageNumbering }),
    ...(section.verticalAlign && { verticalAlign: section.verticalAlign }),
  };
}

/** The paragraph formatting of an imported block, as inline fields (pixels). */
function toParagraphFormat(style: DocParagraphStyle): ParagraphFormat {
  const lineSpacing = fromDocLineSpacing(style);
  return {
    ...(style.alignment && { alignment: style.alignment }),
    ...(lineSpacing && { lineSpacing }),
    ...(style.spaceBeforePt != null && { spaceBefore: ptToPx(style.spaceBeforePt) }),
    ...(style.spaceAfterPt != null && { spaceAfter: ptToPx(style.spaceAfterPt) }),
    ...(style.leftIndentPt != null && { leftIndent: ptToPx(style.leftIndentPt) }),
    ...(style.rightIndentPt != null && { rightIndent: ptToPx(style.rightIndentPt) }),
    ...(style.firstLineIndentPt != null && { firstLineIndent: ptToPx(style.firstLineIndentPt) }),
    ...(style.keepWithNext && { keepWithNext: true }),
    ...(style.keepTogether && { keepTogether: true }),
    ...(style.pageBreakBefore && { pageBreakBefore: true }),
    ...(style.namedStyleId && { namedStyleId: style.namedStyleId }),
  };
}

function toBlockFromDocx(block: DocBlock, textStylePool: TextStylePoolImpl): Block {
  switch (block.type) {
    case 'paragraph':
      return {
        id: generateBlockId(),
        type: 'paragraph',
        content: toInlineContentFromDocx(block.content, textStylePool),
        ...toParagraphFormat(block),
      };
    case 'heading':
      return {
        id: generateBlockId(),
        type: 'heading',
        level: block.level,
        content: toInlineContentFromDocx(block.content, textStylePool),
        ...toParagraphFormat(block),
      };
    case 'listItem':
      return {
        id: generateBlockId(),
        type: 'list-item',
        listType: block.listType,
        level: block.level,
        ...(block.start !== undefined && { start: block.start }),
        content: toInlineContentFromDocx(block.content, textStylePool),
        ...toParagraphFormat(block),
      };
    case 'table':
      return {
        id: generateBlockId(),
        type: 'table',
        rows: block.rows.map((row) => ({
          id: generateRowId(),
          ...(row.repeatHeader && { repeatHeader: true }),
          ...(row.cantSplit && { cantSplit: true }),
          cells: row.cells.map((cell) => ({
            id: generateCellId(),
            content: toInlineContentFromDocx(cell.content, textStylePool),
            ...(cell.blocks && { blocks: cell.blocks.map((b) => toBlockFromDocx(b, textStylePool)) }),
            colspan: cell.colspan,
            rowspan: cell.rowspan,
            ...(cell.backgroundColor && { backgroundColor: cell.backgroundColor }),
            ...(cell.borders && { borders: cell.borders }),
            ...(cell.verticalAlign && { verticalAlign: cell.verticalAlign }),
            ...(cell.textDirection && { textDirection: cell.textDirection }),
            ...(cell.colwidthPx && { colwidth: cell.colwidthPx }),
          })),
        })),
        colWidths: block.colWidthsPx,
        ...(block.borders && { borders: block.borders }),
        ...(block.look && { look: block.look }),
        ...(block.namedStyleId && { namedStyleId: block.namedStyleId }),
      };
    case 'image':
      return {
        id: generateBlockId(),
        type: 'image',
        src: block.src,
        width: block.width,
        height: block.height,
        alt: block.alt,
        alignment: block.alignment,
      };
    case 'horizontalRule':
      return { id: generateBlockId(), type: 'horizontal-rule' };
    case 'pageBreak':
      return { id: generateBlockId(), type: 'page-break' };
  }
}

function toTextStyle(base: DocRunBase): TextStyle {
  return {
    bold: base.bold,
    italic: base.italic,
    underline: base.underline,
    strikethrough: base.strikethrough,
    superscript: base.superscript,
    subscript: base.subscript,
    smallCaps: base.smallCaps,
    allCaps: base.allCaps,
    hidden: base.hidden,
    fontFamily: base.fontFamily,
    fontFamilyEastAsia: base.fontFamilyEastAsia,
    fontFamilyComplex: base.fontFamilyComplex,
    fontSize: base.fontSize,
    color: base.color,
    backgroundColor: base.backgroundColor,
  };
}

function hasAnyRunStyle(base: DocRunBase): boolean {
  return (
    base.bold != null ||
    base.italic != null ||
    base.underline != null ||
    base.strikethrough != null ||
    base.superscript != null ||
    base.subscript != null ||
    base.smallCaps != null ||
    base.allCaps != null ||
    base.hidden != null ||
    base.fontFamily != null ||
    base.fontFamilyEastAsia != null ||
    base.fontFamilyComplex != null ||
    base.fontSize != null ||
    base.color != null ||
    base.backgroundColor != null
  );
}

function toInlineContentFromDocx(content: DocInline[], textStylePool: TextStylePoolImpl): InlineContent[] {
  return content.map((item) => toInlineFromDocx(item, textStylePool)).filter((i): i is InlineContent => i != null);
}

function toInlineFromDocx(item: DocInline, textStylePool: TextStylePoolImpl): InlineContent | null {
  switch (item.type) {
    case 'text':
      return {
        type: 'text',
        text: item.text,
        ...toTextStyle(item),
        ...(item.namedStyleId && { namedStyleId: item.namedStyleId }),
        ...(item.commentThreadId && { commentThreadId: item.commentThreadId }),
      };
    case 'link':
      // InlineLink carries no inline style fields of its own — only a
      // styleId reference — so any run formatting on the link has to go
      // through the text style pool.
      return {
        type: 'link',
        text: item.text,
        href: item.href,
        styleId: hasAnyRunStyle(item) ? textStylePool.getOrCreate(toTextStyle(item)) : undefined,
        ...(item.namedStyleId && { namedStyleId: item.namedStyleId }),
        ...(item.commentThreadId && { commentThreadId: item.commentThreadId }),
      };
    case 'image':
      return { type: 'image', src: item.src, width: item.width, height: item.height, alt: item.alt };
    case 'pageNumber':
      return { type: 'dynamicField', fieldType: 'pageNumber' };
    case 'totalPages':
      return { type: 'dynamicField', fieldType: 'totalPages' };
    case 'field':
      return {
        type: 'field',
        instruction: item.instruction,
        kind: item.kind,
        result:
          typeof item.result === 'string'
            ? item.result
            : item.result.map((r): TextRun | InlineImage =>
                r.type === 'text'
                  ? {
                      type: 'text',
                      text: r.text,
                      ...toTextStyle(r),
                      ...(r.namedStyleId && { namedStyleId: r.namedStyleId }),
                    }
                  : { type: 'image', src: r.src, width: r.width, height: r.height, alt: r.alt }
              ),
        ...(item.dirty && { dirty: true }),
      };
    case 'bookmark':
      return { type: 'bookmark', name: item.name };
  }
}

function toHeaderFooterContentFromDocx(hf: DocHeaderFooter): HeaderFooterContent {
  return {
    blocks: hf.paragraphs.map(toHeaderFooterParagraphFromDocx),
    differentFirstPage: hf.differentFirstPage,
    firstPageBlocks: hf.firstPageParagraphs?.map(toHeaderFooterParagraphFromDocx),
    ...(hf.differentOddEven && { differentOddEven: true }),
    ...(hf.evenPageParagraphs && { evenPageBlocks: hf.evenPageParagraphs.map(toHeaderFooterParagraphFromDocx) }),
  };
}

function toHeaderFooterParagraphFromDocx(p: DocHeaderFooterParagraph): HeaderFooterParagraph {
  return {
    type: 'paragraph',
    content: p.content.map(toHeaderFooterInlineFromDocx).filter((i): i is HeaderFooterInlineContent => i != null),
    alignment: p.alignment,
  };
}

function toHeaderFooterInlineFromDocx(item: DocInline): HeaderFooterInlineContent | null {
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
    // HeaderFooterInlineContent has no link variant — flatten to plain text
    // rather than drop the (usually short, e.g. a URL-as-label) content.
    case 'link':
      return { type: 'text', text: item.text };
    case 'image':
      return { type: 'image', src: item.src, width: item.width, height: item.height, alt: item.alt };
    case 'pageNumber':
      return { type: 'dynamicField', fieldType: 'pageNumber' };
    case 'totalPages':
      return { type: 'dynamicField', fieldType: 'totalPages' };
    // Headers and footers have no generic field or bookmark.
    case 'field':
    case 'bookmark':
      return null;
  }
}
