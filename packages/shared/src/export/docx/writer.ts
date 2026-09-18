import {
  AlignmentType,
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
} from './types';

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

  buildRuns(content: DocInline[]): ParagraphChild[] {
    return content.map((inline) => this.buildRun(inline)).filter((r): r is ParagraphChild => r != null);
  }

  private buildRun(inline: DocInline): ParagraphChild | null {
    switch (inline.type) {
      case 'text':
        return new TextRun(this.runOptions(inline, { text: inline.text }));
      case 'link':
        return new ExternalHyperlink({
          link: inline.href,
          children: [new TextRun(this.runOptions(inline, { text: inline.text, style: 'Hyperlink' }))],
        });
      case 'image':
        return this.buildImage(inline.src, inline.width, inline.height, inline.alt);
      case 'pageNumber':
        return new TextRun({ children: [PageNumber.CURRENT] });
      case 'totalPages':
        return new TextRun({ children: [PageNumber.TOTAL_PAGES] });
      default:
        return null;
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
    const data = Buffer.from(match[2], 'base64');
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
        return this.buildParagraph(block, this.buildRuns(block.content));
      case 'heading':
        return this.buildParagraph(block, this.buildRuns(block.content), { heading: HEADING_MAP[block.level] });
      case 'listItem':
        return this.buildParagraph(block, this.buildRuns(block.content), {
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
      this.buildParagraph({ alignment: p.alignment }, this.buildRuns(p.content))
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
    const width = pxToTwips(section.pageWidthPx);
    const height = pxToTwips(section.pageHeightPx);
    const header = this.buildHeader(section.header);
    const footer = this.buildFooter(section.footer);
    return {
      properties: {
        page: {
          size: { width, height, orientation: section.orientation === 'landscape' ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
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
  const builder = new DocxBuilder();
  const doc = new Document({
    title: input.title,
    numbering: numberingConfig(),
    sections: input.sections.map((s) => builder.buildSection(s)),
  });
  const buffer = await Packer.toBuffer(doc);
  return { bytes: new Uint8Array(buffer), warnings: builder.warnings };
}
