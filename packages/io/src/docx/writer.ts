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
  LineRuleType,
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
  DocFieldRun,
  DocInline,
  DocNamedStyle,
  DocParagraphStyle,
  DocRunBase,
  DocSection,
  DocTableRow,
  DocumentExportInput,
  DocumentExportResult,
} from '@b32nio/spindle-shared/export/docx';

import { addNamedStyles, usableNamedStyles } from './styles-writer';
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

/** The numbering reference for a numbered list whose first number is `start` (1 uses the plain reference). */
function numberedReference(start: number): string {
  return start === 1 ? NUMBERED_REFERENCE : `${NUMBERED_REFERENCE}-from-${start}`;
}

/** Every first number a numbered list in the document starts at, other than 1. */
function numberedStarts(sections: DocSection[]): number[] {
  const starts = new Set<number>();
  for (const section of sections) {
    for (const block of section.blocks) {
      if (block.type === 'listItem' && block.listType === 'numbered' && block.start !== undefined && block.start !== 1 && Number.isInteger(block.start) && block.start > 0) {
        starts.add(block.start);
      }
    }
  }
  return [...starts];
}

function numberingConfig(starts: number[] = []) {
  const numbered = (reference: string, start: number) => ({
    reference,
    levels: Array.from({ length: NUMBERING_LEVELS }, (_, level) => ({
      level,
      format: LevelFormat.DECIMAL,
      text: `%${level + 1}.`,
      alignment: AlignmentType.LEFT,
      // Only the top level of a list carries the document's first number; nested levels start at 1.
      start: level === 0 ? start : 1,
      style: { paragraph: { indent: { left: ptToTwips(18 * (level + 1)), hanging: ptToTwips(18) } } },
    })),
  });
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
      numbered(NUMBERED_REFERENCE, 1),
      ...starts.map((start) => numbered(numberedReference(start), start)),
    ],
  };
}

/** A plain font name sets the Latin, East Asian and complex-script fonts alike; separate ones are written per script. */
function fontOptions(base: DocRunBase): Partial<IRunOptions> {
  const { fontFamily, fontFamilyEastAsia, fontFamilyComplex } = base;
  if (fontFamilyEastAsia == null && fontFamilyComplex == null) return fontFamily != null ? { font: fontFamily } : {};
  return {
    font: {
      ...(fontFamily != null && { ascii: fontFamily, hAnsi: fontFamily }),
      ...((fontFamilyEastAsia ?? fontFamily) != null && { eastAsia: (fontFamilyEastAsia ?? fontFamily) as string }),
      ...((fontFamilyComplex ?? fontFamily) != null && { cs: (fontFamilyComplex ?? fontFamily) as string }),
    },
  } as Partial<IRunOptions>;
}

class DocxBuilder {
  readonly warnings: string[] = [];
  /** Orphan comment threads still waiting for the first body paragraph to be anchored to. */
  private pendingOrphans: CommentPlanThread[];
  private inNumberedRun = false;
  private numberedRun = { reference: NUMBERED_REFERENCE, instance: 0 };
  private numberedInstances = 0;

  constructor(
    private readonly comments: CommentPlan,
    private readonly namedStyles: Record<string, DocNamedStyle> = {}
  ) {
    this.pendingOrphans = comments.orphans;
  }

  /** A named style of the given kind that the document defines (a reference to a removed or foreign style is dropped). */
  private namedStyle(id: string | undefined, kind: DocNamedStyle['kind']): string | undefined {
    return id && Object.prototype.hasOwnProperty.call(this.namedStyles, id) && this.namedStyles[id].kind === kind ? id : undefined;
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
      case 'field':
        return this.buildField(inline);
      default:
        return [];
    }
  }

  /** A complex field: begin, instruction, separate, the cached result runs, end, so Word can refresh it. */
  private buildField(field: DocFieldRun): ParagraphChild[] {
    const fldChar = (type: 'begin' | 'separate' | 'end') =>
      new TextRun({
        children: [
          new BuilderElement({
            name: 'w:fldChar',
            attributes: {
              type: { key: 'w:fldCharType', value: type },
              ...(type === 'begin' && field.dirty && { dirty: { key: 'w:dirty', value: 'true' } }),
            },
          }),
        ],
      });
    const instruction = new TextRun({
      children: [
        new BuilderElement({
          name: 'w:instrText',
          attributes: { space: { key: 'xml:space', value: 'preserve' } },
          // The xml formatter writes string children as the element's text.
          children: [` ${field.instruction} `] as unknown as BuilderElement[],
        }),
      ],
    });
    const result: ParagraphChild[] =
      typeof field.result === 'string'
        ? field.result
          ? this.textRuns({}, field.result)
          : []
        : field.result.flatMap((r) =>
            r.type === 'text' ? this.textRuns(r, r.text) : [this.buildImage(r.src, r.width, r.height, r.alt)].filter((c): c is ParagraphChild => c != null)
          );
    return [fldChar('begin'), instruction, fldChar('separate'), ...result, fldChar('end')];
  }

  private runOptions(base: DocRunBase, extra: Partial<IRunOptions>): IRunOptions {
    const style = this.namedStyle(base.namedStyleId, 'character');
    return {
      ...extra,
      ...(style && { style }),
      ...(base.bold != null && { bold: base.bold }),
      ...(base.italic != null && { italics: base.italic }),
      ...(base.underline && { underline: { type: UnderlineType.SINGLE } }),
      ...(base.strikethrough != null && { strike: base.strikethrough }),
      ...(base.superscript && { superScript: true }),
      ...(base.subscript && { subScript: true }),
      ...(base.smallCaps != null && { smallCaps: base.smallCaps }),
      ...(base.allCaps != null && { allCaps: base.allCaps }),
      ...(base.hidden != null && { vanish: base.hidden }),
      ...fontOptions(base),
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
    const custom = style.lineSpacingCustom;
    const spacingLine =
      custom != null
        ? custom.rule === 'multiple'
          ? { line: Math.round(custom.value * 240), lineRule: LineRuleType.AUTO }
          : { line: ptToTwips(custom.value), lineRule: custom.rule === 'exactly' ? LineRuleType.EXACT : LineRuleType.AT_LEAST }
        : { line: style.lineSpacing ? LINE_SPACING_MAP[style.lineSpacing] : undefined };
    const firstLine = style.firstLineIndentPt;
    const namedStyle = this.namedStyle(style.namedStyleId, 'paragraph');
    return new Paragraph({
      children,
      ...(namedStyle && { style: namedStyle }),
      alignment: style.alignment ? ALIGNMENT_MAP[style.alignment] : undefined,
      spacing: {
        before: style.spaceBeforePt != null ? ptToTwips(style.spaceBeforePt) : undefined,
        after: style.spaceAfterPt != null ? ptToTwips(style.spaceAfterPt) : undefined,
        ...spacingLine,
      },
      indent: {
        left: style.leftIndentPt != null ? ptToTwips(style.leftIndentPt) : undefined,
        right: style.rightIndentPt != null ? ptToTwips(style.rightIndentPt) : undefined,
        // A negative first-line indent is a hanging indent.
        firstLine: firstLine != null && firstLine >= 0 ? ptToTwips(firstLine) : undefined,
        hanging: firstLine != null && firstLine < 0 ? ptToTwips(-firstLine) : undefined,
      },
      ...(style.keepWithNext && { keepNext: true }),
      ...(style.keepTogether && { keepLines: true }),
      ...(style.pageBreakBefore && { pageBreakBefore: true }),
      ...extra,
    });
  }

  /** The numbering of the current list run: a new run (instance) begins where the previous block is not a numbered item, or an item states its own first number. */
  private listNumbering(block: Extract<DocBlock, { type: 'listItem' }>) {
    const level = Math.min(block.level, NUMBERING_LEVELS - 1);
    if (block.listType === 'bullet') return { reference: BULLET_REFERENCE, level };
    if (!this.inNumberedRun || block.start !== undefined) {
      const start = block.start !== undefined && Number.isInteger(block.start) && block.start > 0 ? block.start : 1;
      this.numberedRun = { reference: numberedReference(start), instance: ++this.numberedInstances };
    }
    this.inNumberedRun = true;
    return { ...this.numberedRun, level };
  }

  buildBlock(block: DocBlock): Paragraph | Table {
    if (!(block.type === 'listItem' && block.listType === 'numbered')) this.inNumberedRun = false;
    switch (block.type) {
      case 'paragraph':
        return this.buildParagraph(block, this.withOrphans(this.buildRuns(block.content)));
      case 'heading':
        // A heading in a named style of the document is written in that style (which carries the outline level).
        return this.buildParagraph(
          block,
          this.withOrphans(this.buildRuns(block.content)),
          this.namedStyle(block.namedStyleId, 'paragraph') ? {} : { heading: HEADING_MAP[block.level] }
        );
      case 'listItem':
        return this.buildParagraph(block, this.withOrphans(this.buildRuns(block.content)), {
          numbering: this.listNumbering(block),
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
    this.inNumberedRun = false;
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
  const namedStyles = usableNamedStyles(input.namedStyles);
  const builder = new DocxBuilder(comments, namedStyles);
  const doc = new Document({
    title: input.title,
    numbering: numberingConfig(numberedStarts(input.sections)),
    sections: input.sections.map((s) => builder.buildSection(s)),
  });
  // toArrayBuffer, not toBuffer: JSZip's "nodebuffer" output does not exist in browsers,
  // so toBuffer made every in-browser DOCX download fail.
  const buffer = await Packer.toArrayBuffer(doc);
  let bytes: Uint8Array = new Uint8Array(buffer);
  bytes = await addNamedStyles(bytes, namedStyles);
  if (comments.threads.size > 0) {
    const unplaced = builder.unplaced();
    const lost = unplaced.reduce((n, t) => n + t.ids.length, 0);
    if (lost > 0) builder.warnings.push(`${lost} comment(s) could not be anchored (the document has no paragraph) and were not written.`);
    for (const t of unplaced) comments.threads.delete(t.thread.id);
    if (comments.threads.size > 0) bytes = await addCommentParts(bytes, buildCommentParts(comments));
  }
  return { bytes, warnings: builder.warnings };
}
