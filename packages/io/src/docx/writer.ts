import {
  AlignmentType,
  BuilderElement,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  Document,
  ExternalHyperlink,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
  LevelFormat,
  Packer,
  PageBreak,
  PageNumber,
  PageOrientation,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  Tab,
  TextRun,
  UnderlineType,
  WidthType,
  type IRunOptions,
  type ParagraphChild,
} from 'docx';
import type {
  DocBlock,
  DocHeaderFooter,
  DocHeaderFooterParagraph,
  DocInline,
  DocParagraphStyle,
  DocRunBase,
  DocSection,
  DocTableRow,
  DocumentExportInput,
  DocumentExportResult,
} from '@b32nio/spindle-shared/export/docx';

import { addCommentParts, buildCommentParts, planComments, type CommentPlan, type CommentPlanThread } from './comments-writer';

// docx's own unit conventions: twips (1/1440in) for most measurements, but
// half-points for font size — both derived from our px-at-96dpi convention.
const PX_TO_TWIPS = 15; // 1440 twips/in / 96px/in
const PT_TO_TWIPS = 20; // 1440 twips/in / 72pt/in
const PT_TO_HALF_POINTS = 2;

function pxToTwips(px: number): number {
  return Math.round(px * PX_TO_TWIPS);
}

function ptToTwips(pt: number): number {
  return Math.round(pt * PT_TO_TWIPS);
}

function toDocxColor(hex: string): string {
  return hex.replace('#', '').toUpperCase();
}

const BULLET_CHARS = ['●', '○', '■', '◆']; // filled/hollow circle, square, diamond — cycled by level
const NUMBERING_LEVELS = 9; // Word's own max nesting depth

const BULLET_REFERENCE = 'spindle-bullet-list';
const NUMBERED_REFERENCE = 'spindle-numbered-list';

function numberingConfig() {
  return {
    config: [
      {
        reference: BULLET_REFERENCE,
        levels: Array.from({ length: NUMBERING_LEVELS }, (_, level) => ({
          level,
          format: LevelFormat.BULLET,
          text: BULLET_CHARS[level % BULLET_CHARS.length],
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: ptToTwips(18 * (level + 1)), hanging: ptToTwips(18) } } },
        })),
      },
      {
        reference: NUMBERED_REFERENCE,
        levels: Array.from({ length: NUMBERING_LEVELS }, (_, level) => ({
          level,
          format: LevelFormat.DECIMAL,
          text: `%${level + 1}.`,
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: ptToTwips(18 * (level + 1)), hanging: ptToTwips(18) } } },
        })),
      },
    ],
  };
}

class DocxBuilder {
  readonly warnings: string[] = [];
  /** Orphan comment threads still waiting for the first body paragraph to be anchored to. */
  private pendingOrphans: CommentPlanThread[];

  constructor(private readonly comments: CommentPlan) {
    this.pendingOrphans = comments.orphans;
  }

  /** `anchorComments: false` for header/footer content, which has no comment anchors. */
  buildRuns(content: DocInline[], anchorComments = true): ParagraphChild[] {
    const out: ParagraphChild[] = [];
    for (const inline of content) {
      const planned =
        anchorComments && (inline.type === 'text' || inline.type === 'link') && inline.commentThreadId
          ? this.comments.threads.get(inline.commentThreadId)
          : undefined;
      // A comment's range runs from its first tagged run to its last; replies share the range of their root.
      if (planned && planned.seen === 0) out.push(...this.rangeStarts(planned));
      out.push(...this.buildRun(inline).filter((r): r is ParagraphChild => r != null));
      if (planned && ++planned.seen === planned.total) out.push(...this.rangeEnds(planned));
    }
    return out;
  }

  private rangeStarts(planned: CommentPlanThread): ParagraphChild[] {
    return planned.ids.map((id) => new CommentRangeStart(id));
  }

  private rangeEnds(planned: CommentPlanThread): ParagraphChild[] {
    return planned.ids.flatMap((id) => [new CommentRangeEnd(id), new TextRun({ children: [new CommentReference(id)] })]);
  }

  /** Comments without a surviving anchor are kept as collapsed ranges at the start of the first body paragraph. */
  private withOrphans(children: ParagraphChild[]): ParagraphChild[] {
    if (this.pendingOrphans.length === 0) return children;
    const markers = this.pendingOrphans.flatMap((planned) => [...this.rangeStarts(planned), ...this.rangeEnds(planned)]);
    this.pendingOrphans = [];
    return [...markers, ...children];
  }

  /** Orphan threads that never found a body paragraph (a document without one), so cannot be written. */
  unplaced(): CommentPlanThread[] {
    return this.pendingOrphans;
  }

  /**
   * A text run's text with '\t' written as w:tab and '\n' as w:br, all inside one run so the formatting
   * is shared. Text with neither stays a plain `text` run.
   */
  private textRuns(base: DocRunBase, text: string, extra: Partial<IRunOptions> = {}): TextRun[] {
    if (!/[\t\n]/.test(text)) return [new TextRun(this.runOptions(base, { ...extra, text }))];
    const children: (string | Tab | BuilderElement)[] = [];
    for (const part of text.split(/(\t|\n)/)) {
      if (part === '\t') children.push(new Tab());
      else if (part === '\n') children.push(new BuilderElement({ name: 'w:br' }));
      else if (part) children.push(part);
    }
    return [new TextRun(this.runOptions(base, { ...extra, children }))];
  }

  private buildRun(inline: DocInline): (ParagraphChild | null)[] {
    switch (inline.type) {
      case 'text':
        return this.textRuns(inline, inline.text);
      case 'link':
        return [
          new ExternalHyperlink({
            link: inline.href,
            children: this.textRuns(inline, inline.text, { style: 'Hyperlink' }),
          }),
        ];
      case 'image':
        return [this.buildImage(inline.src, inline.width, inline.height, inline.alt)];
      case 'pageNumber':
        return [new TextRun({ children: [PageNumber.CURRENT] })];
      case 'totalPages':
        return [new TextRun({ children: [PageNumber.TOTAL_PAGES] })];
      default:
        return [];
    }
  }

  private runOptions(base: DocRunBase, extra: Partial<IRunOptions>): IRunOptions {
    return {
      ...extra,
      ...(base.bold != null && { bold: base.bold }),
      ...(base.italic != null && { italics: base.italic }),
      ...(base.underline && { underline: { type: UnderlineType.SINGLE } }),
      ...(base.strikethrough != null && { strike: base.strikethrough }),
      ...(base.superscript && { superScript: true }),
      ...(base.subscript && { subScript: true }),
      ...(base.smallCaps != null && { smallCaps: base.smallCaps }),
      ...(base.fontFamily != null && { font: base.fontFamily }),
      ...(base.fontSize != null && { size: Math.round(base.fontSize * PT_TO_HALF_POINTS) }),
      ...(base.color != null && { color: toDocxColor(base.color) }),
      ...(base.backgroundColor != null && {
        shading: { type: ShadingType.CLEAR, fill: toDocxColor(base.backgroundColor) },
      }),
    } as IRunOptions;
  }

  buildImage(src: string, widthPx: number, heightPx: number, alt?: string): ParagraphChild | null {
    const match = src.match(/^data:image\/(png|jpe?g|gif|bmp);base64,(.+)$/i);
    if (!match) {
      this.warnings.push(`Image with src "${truncate(src)}" is not an embeddable data: URI — skipped.`);
      return new TextRun({ text: alt ? `[Image: ${alt}]` : '[Image]', italics: true });
    }
    const ext = match[1].toLowerCase();
    const type = ext === 'jpeg' ? 'jpg' : (ext as 'png' | 'jpg' | 'gif' | 'bmp');
    // atob + Uint8Array, not Buffer: Buffer does not exist in browsers.
    const data = Uint8Array.from(atob(match[2]), (c) => c.charCodeAt(0));
    return new ImageRun({
      type,
      data,
      transformation: { width: widthPx, height: heightPx },
      altText: alt ? { name: alt, description: alt, title: alt } : undefined,
    } as ConstructorParameters<typeof ImageRun>[0]);
  }

  buildParagraph(style: DocParagraphStyle, children: ParagraphChild[], extra: Record<string, unknown> = {}): Paragraph {
    return new Paragraph({
      children,
      alignment: style.alignment ? ALIGNMENT_MAP[style.alignment] : undefined,
      spacing: {
        before: style.spaceBeforePt != null ? ptToTwips(style.spaceBeforePt) : undefined,
        after: style.spaceAfterPt != null ? ptToTwips(style.spaceAfterPt) : undefined,
        line: style.lineSpacing ? LINE_SPACING_MAP[style.lineSpacing] : undefined,
      },
      indent: {
        left: style.leftIndentPt != null ? ptToTwips(style.leftIndentPt) : undefined,
        firstLine: style.firstLineIndentPt != null ? ptToTwips(style.firstLineIndentPt) : undefined,
      },
      ...extra,
    });
  }

  buildBlock(block: DocBlock): Paragraph | Table {
    switch (block.type) {
      case 'paragraph':
        return this.buildParagraph(block, this.withOrphans(this.buildRuns(block.content)));
      case 'heading':
        return this.buildParagraph(block, this.withOrphans(this.buildRuns(block.content)), { heading: HEADING_MAP[block.level] });
      case 'listItem':
        return this.buildParagraph(block, this.withOrphans(this.buildRuns(block.content)), {
          numbering: {
            reference: block.listType === 'bullet' ? BULLET_REFERENCE : NUMBERED_REFERENCE,
            level: Math.min(block.level, NUMBERING_LEVELS - 1),
          },
        });
      case 'table':
        return this.buildTable(block.rows, block.colWidthsPx);
      case 'image': {
        const run = this.buildImage(block.src, block.width, block.height, block.alt);
        return this.buildParagraph(
          { alignment: block.alignment },
          run ? [run] : []
        );
      }
      case 'horizontalRule':
        // A common docx idiom: an empty paragraph with only a bottom border.
        return new Paragraph({
          border: { bottom: { style: 'single', size: 6, color: '999999' } },
        });
      case 'pageBreak':
        return new Paragraph({ children: [new PageBreak()] });
    }
  }

  private buildTable(rows: DocTableRow[], colWidthsPx: number[] | undefined): Table {
    const tableRows = rows.map(
      (row) =>
        new TableRow({
          children: row.cells.map(
            (cell) =>
              new TableCell({
                children: [this.buildParagraph({}, this.buildRuns(cell.content))],
                columnSpan: cell.colspan && cell.colspan > 1 ? cell.colspan : undefined,
                rowSpan: cell.rowspan && cell.rowspan > 1 ? cell.rowspan : undefined,
              })
          ),
        })
    );
    return new Table({
      rows: tableRows,
      columnWidths: colWidthsPx?.map(pxToTwips),
      width: { size: 100, type: WidthType.PERCENTAGE },
    });
  }

  buildHeaderFooterParagraphs(paragraphs: DocHeaderFooterParagraph[]): Paragraph[] {
    return paragraphs.map((p) =>
      this.buildParagraph({ alignment: p.alignment }, this.buildRuns(p.content, false))
    );
  }

  buildHeader(hf: DocHeaderFooter | undefined): Header | undefined {
    if (!hf) return undefined;
    return new Header({ children: this.buildHeaderFooterParagraphs(hf.paragraphs) });
  }

  buildFooter(hf: DocHeaderFooter | undefined): Footer | undefined {
    if (!hf) return undefined;
    return new Footer({ children: this.buildHeaderFooterParagraphs(hf.paragraphs) });
  }

  buildSection(section: DocSection) {
    // `docx` swaps width/height itself when the orientation is landscape, so it
    // is handed the portrait pair; the file then carries the on-page dimensions.
    const landscape = section.orientation === 'landscape';
    const width = pxToTwips(landscape ? section.pageHeightPx : section.pageWidthPx);
    const height = pxToTwips(landscape ? section.pageWidthPx : section.pageHeightPx);
    const header = this.buildHeader(section.header);
    const footer = this.buildFooter(section.footer);
    return {
      properties: {
        page: {
          size: { width, height, orientation: landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
          margin: {
            top: pxToTwips(section.marginTopPx),
            right: pxToTwips(section.marginRightPx),
            bottom: pxToTwips(section.marginBottomPx),
            left: pxToTwips(section.marginLeftPx),
          },
        },
        titlePage: section.header?.differentFirstPage || section.footer?.differentFirstPage,
      },
      headers: header ? { default: header } : undefined,
      footers: footer ? { default: footer } : undefined,
      children: section.blocks.map((b) => this.buildBlock(b)),
    };
  }
}

const ALIGNMENT_MAP = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
} as const;

const LINE_SPACING_MAP = {
  single: 240,
  onePointFive: 360,
  double: 480,
} as const;

const HEADING_MAP = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
  5: HeadingLevel.HEADING_5,
  6: HeadingLevel.HEADING_6,
} as const;

function truncate(s: string, max = 60): string {
  return s.length > max ? `${s.slice(0, max)}…` : s;
}

export async function exportDocx(input: DocumentExportInput): Promise<DocumentExportResult> {
  const comments = planComments(input.comments, input.sections);
  const builder = new DocxBuilder(comments);
  const doc = new Document({
    title: input.title,
    numbering: numberingConfig(),
    sections: input.sections.map((s) => builder.buildSection(s)),
  });
  // toArrayBuffer, not toBuffer: JSZip's "nodebuffer" output does not exist in browsers,
  // so toBuffer made every in-browser DOCX download fail.
  const buffer = await Packer.toArrayBuffer(doc);
  let bytes = new Uint8Array(buffer);
  if (comments.threads.size > 0) {
    const unplaced = builder.unplaced();
    const lost = unplaced.reduce((n, t) => n + t.ids.length, 0);
    if (lost > 0) builder.warnings.push(`${lost} comment(s) could not be anchored (the document has no paragraph) and were not written.`);
    for (const t of unplaced) comments.threads.delete(t.thread.id);
    if (comments.threads.size > 0) bytes = await addCommentParts(bytes, buildCommentParts(comments));
  }
  return { bytes, warnings: builder.warnings };
}
