import {
  AlignmentType,
  Column,
  Bookmark,
  BorderStyle,
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
  InternalHyperlink,
  LevelFormat,
  LineRuleType,
  Packer,
  PageBreak,
  PageNumber,
  PageOrientation,
  Paragraph,
  ShadingType,
  TextDirection,
  Table,
  TableCell,
  TableRow,
  Tab,
  TextRun,
  UnderlineType,
  WidthType,
  NumberFormat,
  SectionType,
  VerticalAlignSection,
  type IBorderOptions,
  type IRunOptions,
  type ITableBordersOptions,
  type ITableCellBorders,
  type ParagraphChild,
} from 'docx';
import type {
  DocBlock,
  DocBorder,
  DocBorders,
  DocContentControl,
  DocHeaderFooter,
  DocHeaderFooterParagraph,
  DocFieldRun,
  DocInline,
  DocNamedStyle,
  DocParagraphStyle,
  DocRunBase,
  DocSection,
  DocumentExportInput,
  DocumentExportResult,
} from '@b32nio/spindle-shared/export/docx';

import { addNamedStyles, usableNamedStyles } from './styles-writer';
import { addNoteParts, type WrittenNotes } from './notes-writer';
import { addCommentParts, buildCommentParts, planComments, type CommentPlan, type CommentPlanThread } from './comments-writer';

// docx's own unit conventions: twips (1/1440in) for most measurements, but
// half-points for font size — both derived from our px-at-96dpi convention.
const PX_TO_TWIPS = 15; // 1440 twips/in / 96px/in
const PT_TO_TWIPS = 20; // 1440 twips/in / 72pt/in
const PT_TO_HALF_POINTS = 2;

function pxToTwips(px: number): number {
  return Math.round(px * PX_TO_TWIPS);
}

const BORDER_STYLES = { single: BorderStyle.SINGLE, double: BorderStyle.DOUBLE, dashed: BorderStyle.DASHED, dotted: BorderStyle.DOTTED, none: BorderStyle.NONE } as const;

function borderOptions(border: DocBorder | undefined): IBorderOptions {
  if (!border || border.style === 'none') return { style: BorderStyle.NONE, size: 0, color: 'auto' };
  // `w:sz` is in eighths of a point (a pixel is 6); Word's range is 2 to 96.
  const size = Math.min(96, Math.max(2, Math.round((border.widthPx ?? 1) * 6)));
  return { style: BORDER_STYLES[border.style], size, color: border.color ? border.color.replace('#', '').toUpperCase() : 'auto' };
}

/** A cell lists only the sides it sets (the rest come from the table). */
function cellBorders(b: DocBorders): ITableCellBorders {
  return {
    ...(b.top && { top: borderOptions(b.top) }),
    ...(b.left && { left: borderOptions(b.left) }),
    ...(b.bottom && { bottom: borderOptions(b.bottom) }),
    ...(b.right && { right: borderOptions(b.right) }),
  };
}

/** A table that has `borders` has exactly those lines: a side it does not list has none (the table's style borders are already in them). */
function tableBorders(b: DocBorders): ITableBordersOptions {
  return {
    top: borderOptions(b.top),
    left: borderOptions(b.left),
    bottom: borderOptions(b.bottom),
    right: borderOptions(b.right),
    insideHorizontal: borderOptions(b.insideH),
    insideVertical: borderOptions(b.insideV),
  };
}

const BREAK_TYPES = { continuous: SectionType.CONTINUOUS, evenPage: SectionType.EVEN_PAGE, oddPage: SectionType.ODD_PAGE } as const;

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
  const visit = (block: DocBlock) => {
    if (block.type === 'listItem' && block.listType === 'numbered' && block.start !== undefined && block.start !== 1 && Number.isInteger(block.start) && block.start > 0) {
      starts.add(block.start);
    } else if (block.type === 'table') {
      for (const row of block.rows) for (const cell of row.cells) cell.blocks?.forEach(visit);
    } else if (block.type === 'contentControl') {
      block.blocks.forEach(visit);
    }
  };
  for (const section of sections) section.blocks.forEach(visit);
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

/** Whether the last thing a list of blocks ends with is a table (a table cell must end with a paragraph). */
function endsWithTable(block: DocBlock | undefined): boolean {
  if (!block) return false;
  if (block.type === 'table') return true;
  return block.type === 'contentControl' && endsWithTable(block.blocks[block.blocks.length - 1]);
}

/** The glyphs Word draws in a check box content control. */
const CHECKED_GLYPH = '2612';
const UNCHECKED_GLYPH = '2610';

/** The inline content with note references taken out (a note holds no notes). */
function withoutNoteRefs(content: DocInline[]): DocInline[] {
  return content.flatMap((item): DocInline[] => {
    if (item.type === 'noteRef') return [];
    return item.type === 'contentControl' ? [{ ...item, content: withoutNoteRefs(item.content) }] : [item];
  });
}

class DocxBuilder {
  readonly warnings: string[] = [];
  /** The notes referenced so far, in document order (a note's id in the file is its position + 1). */
  readonly notes: WrittenNotes = { footnote: [], endnote: [] };
  /** `w:id` of the next content control. */
  private nextControlId = 1000001;
  /** Orphan comment threads still waiting for the first body paragraph to be anchored to. */
  private pendingOrphans: CommentPlanThread[];
  private inNumberedRun = false;
  private numberedRun = { reference: NUMBERED_REFERENCE, instance: 0 };
  private numberedInstances = 0;
  /** Some header or footer of the document shows other content on even pages, so the document switches even/odd headers on. */
  evenAndOdd = false;

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
      case 'bookmark':
        return inline.name ? [new Bookmark({ id: inline.name, children: [] })] : [];
      case 'link':
        // `#name` points at a bookmark in this document.
        if (inline.href.startsWith('#') && inline.href.length > 1) {
          return [
            new InternalHyperlink({
              anchor: inline.href.slice(1),
              // No Hyperlink character style: Word's own internal links (tables of contents) look like their text.
              children: this.textRuns(inline, inline.text),
            }),
          ];
        }
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
      case 'contentControl':
        return [this.buildControl(inline.control, this.buildRuns(inline.content))];
      case 'noteRef':
        return [this.buildNoteReference(inline)];
      default:
        return [];
    }
  }

  /** A superscript run holding the reference to a note; the note itself is written to the notes part afterwards. */
  private buildNoteReference(ref: Extract<DocInline, { type: 'noteRef' }>): ParagraphChild {
    const list = this.notes[ref.kind];
    list.push({ content: withoutNoteRefs(ref.content), ...(ref.customMark && { customMark: ref.customMark }) });
    const reference = new BuilderElement({
      name: ref.kind === 'footnote' ? 'w:footnoteReference' : 'w:endnoteReference',
      attributes: {
        id: { key: 'w:id', value: String(list.length) },
        ...(ref.customMark && { follows: { key: 'w:customMarkFollows', value: '1' } }),
      },
    });
    return new TextRun({ superScript: true, children: ref.customMark ? [reference, ref.customMark] : [reference] });
  }

  /** A `w:sdt` (run-level or block-level) with the properties of `control` around already built `content`. */
  private buildControl(control: DocContentControl, content: unknown[]): ParagraphChild {
    const el = (name: string, attrs: Record<string, string | undefined> = {}, children?: BuilderElement[]) =>
      new BuilderElement({
        name,
        attributes: Object.fromEntries(Object.entries(attrs).filter(([, v]) => v !== undefined).map(([k, v]) => [k, { key: k, value: v as string }])),
        ...(children && { children }),
      });
    const listItems = (tag: string) =>
      el(tag, {}, (control.items ?? []).map((item) => el('w:listItem', { 'w:displayText': item.text, 'w:value': item.value })));
    const properties: BuilderElement[] = [];
    if (control.alias) properties.push(el('w:alias', { 'w:val': control.alias }));
    if (control.tag) properties.push(el('w:tag', { 'w:val': control.tag }));
    properties.push(el('w:id', { 'w:val': String(this.nextControlId++) }));
    if (control.lock) properties.push(el('w:lock', { 'w:val': control.lock }));
    if (control.placeholder) properties.push(el('w:placeholder', {}, [el('w:docPart', { 'w:val': control.placeholder })]));
    if (control.showingPlaceholder) properties.push(el('w:showingPlcHdr'));
    switch (control.type) {
      case 'text':
        properties.push(el('w:text', control.multiline ? { 'w:multiLine': '1' } : {}));
        break;
      case 'dropDown':
        properties.push(listItems('w:dropDownList'));
        break;
      case 'comboBox':
        properties.push(listItems('w:comboBox'));
        break;
      case 'date':
        properties.push(
          el('w:date', control.dateValue ? { 'w:fullDate': control.dateValue } : {}, [
            ...(control.dateFormat ? [el('w:dateFormat', { 'w:val': control.dateFormat })] : []),
            el('w:lid', { 'w:val': 'en-US' }),
            el('w:storeMappedDataAs', { 'w:val': 'dateTime' }),
            el('w:calendar', { 'w:val': 'gregorian' }),
          ])
        );
        break;
      case 'checkbox':
        properties.push(
          el('w14:checkbox', {}, [
            el('w14:checked', { 'w14:val': control.checked ? '1' : '0' }),
            el('w14:checkedState', { 'w14:val': CHECKED_GLYPH, 'w14:font': 'MS Gothic' }),
            el('w14:uncheckedState', { 'w14:val': UNCHECKED_GLYPH, 'w14:font': 'MS Gothic' }),
          ])
        );
        break;
      case 'picture':
        properties.push(el('w:picture'));
        break;
      case 'docPart':
        properties.push(el('w:docPartObj', {}, control.gallery ? [el('w:docPartGallery', { 'w:val': control.gallery }), el('w:docPartUnique')] : [el('w:docPartUnique')]));
        break;
      case 'richText':
        break;
    }
    return el('w:sdt', {}, [el('w:sdtPr', {}, properties), el('w:sdtContent', {}, content as BuilderElement[])]) as unknown as ParagraphChild;
  }

  /** A block-level control wraps its blocks; it must hold at least one paragraph. */
  private buildControlBlock(block: Extract<DocBlock, { type: 'contentControl' }>): Paragraph {
    this.inNumberedRun = false;
    const children: unknown[] = block.blocks.map((b) => this.buildBlock(b));
    if (children.length === 0) children.push(this.buildParagraph({}, []));
    this.inNumberedRun = false;
    // The writer's block builders return docx paragraphs and tables, which the section and cell take as children; the wrapper is another XML element of the same kind.
    return this.buildControl(block.control, children) as unknown as Paragraph;
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
        return this.buildTable(block);
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
      case 'contentControl':
        return this.buildControlBlock(block);
    }
  }

  /** A cell's paragraphs and nested tables. A numbered list in a cell numbers on its own, and a cell must end with a paragraph. */
  private buildCellChildren(cell: Extract<DocBlock, { type: 'table' }>['rows'][number]['cells'][number]): (Paragraph | Table)[] {
    if (!cell.blocks || cell.blocks.length === 0) return [this.buildParagraph({}, this.buildRuns(cell.content))];
    this.inNumberedRun = false;
    const children = cell.blocks.map((b) => this.buildBlock(b));
    this.inNumberedRun = false;
    if (endsWithTable(cell.blocks[cell.blocks.length - 1])) children.push(new Paragraph({}));
    return children;
  }

  private buildTable(block: Extract<DocBlock, { type: 'table' }>): Table {
    const { rows, colWidthsPx } = block;
    const tableRows = rows.map(
      (row) =>
        new TableRow({
          ...(row.repeatHeader && { tableHeader: true }),
          ...(row.cantSplit && { cantSplit: true }),
          children: row.cells.map((cell) => {
            const widthPx = cell.colwidthPx?.reduce((sum, w) => sum + w, 0);
            return new TableCell({
              children: this.buildCellChildren(cell),
              columnSpan: cell.colspan && cell.colspan > 1 ? cell.colspan : undefined,
              rowSpan: cell.rowspan && cell.rowspan > 1 ? cell.rowspan : undefined,
              ...(cell.backgroundColor && { shading: { type: ShadingType.CLEAR, color: 'auto', fill: cell.backgroundColor.replace('#', '').toUpperCase() } }),
              ...(cell.borders && { borders: cellBorders(cell.borders) }),
              ...(cell.verticalAlign && { verticalAlign: cell.verticalAlign }),
              ...(cell.textDirection && { textDirection: cell.textDirection === 'btLr' ? TextDirection.BOTTOM_TO_TOP_LEFT_TO_RIGHT : TextDirection.TOP_TO_BOTTOM_RIGHT_TO_LEFT }),
              ...(widthPx && widthPx > 0 && { width: { size: pxToTwips(widthPx), type: WidthType.DXA } }),
            });
          }),
        })
    );
    const style = this.namedStyle(block.namedStyleId, 'table');
    const look = block.look;
    return new Table({
      rows: tableRows,
      columnWidths: colWidthsPx?.map(pxToTwips),
      width: colWidthsPx && colWidthsPx.length > 0 ? { size: pxToTwips(colWidthsPx.reduce((sum, w) => sum + w, 0)), type: WidthType.DXA } : { size: 100, type: WidthType.PERCENTAGE },
      ...(style && { style }),
      ...(block.borders && { borders: tableBorders(block.borders) }),
      ...(look && { tableLook: { firstRow: !!look.firstRow, lastRow: !!look.lastRow, firstColumn: !!look.firstColumn, lastColumn: !!look.lastColumn, noHBand: !!look.noHBand, noVBand: !!look.noVBand } }),
    });
  }

  buildHeaderFooterParagraphs(paragraphs: DocHeaderFooterParagraph[]): Paragraph[] {
    return paragraphs.map((p) =>
      this.buildParagraph({ alignment: p.alignment }, this.buildRuns(p.content, false))
    );
  }

  /**
   * The parts a header or footer is written as: the default one, a first-page one and an even-page one. A variant that
   * has content is written as it is. One that is switched on without content gets no part (the first page, or the even
   * pages, are blank, or take the previous section's, as in Word). One that is not switched on for this header or
   * footer but is for the section (`w:titlePg` because the header differs) or the document (even/odd headers) repeats
   * the default content, so the header or footer still shows there.
   */
  private buildVariants<T>(
    hf: DocHeaderFooter | undefined,
    titlePage: boolean,
    make: (paragraphs: DocHeaderFooterParagraph[]) => T
  ): { default: T; first?: T; even?: T } | undefined {
    if (!hf) return undefined;
    const firstParagraphs = hf.firstPageParagraphs ?? (titlePage && !hf.differentFirstPage ? hf.paragraphs : undefined);
    const evenParagraphs = hf.evenPageParagraphs ?? (this.evenAndOdd && !hf.differentOddEven ? hf.paragraphs : undefined);
    return {
      default: make(hf.paragraphs),
      ...(firstParagraphs && { first: make(firstParagraphs) }),
      ...(evenParagraphs && { even: make(evenParagraphs) }),
    };
  }

  buildHeader(hf: DocHeaderFooter | undefined, titlePage = false) {
    return this.buildVariants(hf, titlePage, (paragraphs) => new Header({ children: this.buildHeaderFooterParagraphs(paragraphs) }));
  }

  buildFooter(hf: DocHeaderFooter | undefined, titlePage = false) {
    return this.buildVariants(hf, titlePage, (paragraphs) => new Footer({ children: this.buildHeaderFooterParagraphs(paragraphs) }));
  }

  buildSection(section: DocSection) {
    // `docx` swaps width/height itself when the orientation is landscape, so it
    // is handed the portrait pair; the file then carries the on-page dimensions.
    const landscape = section.orientation === 'landscape';
    const width = pxToTwips(landscape ? section.pageHeightPx : section.pageWidthPx);
    const height = pxToTwips(landscape ? section.pageWidthPx : section.pageHeightPx);
    this.inNumberedRun = false;
    const titlePage = !!(section.header?.differentFirstPage || section.footer?.differentFirstPage);
    const header = this.buildHeader(section.header, titlePage);
    const footer = this.buildFooter(section.footer, titlePage);
    const numbering = section.pageNumbering;
    const numberFormat = numbering?.format && (Object.values(NumberFormat) as string[]).includes(numbering.format) ? numbering.format : undefined;
    const columns = section.columns;
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
          ...(numbering && {
            pageNumbers: {
              ...(numbering.start !== undefined && Number.isInteger(numbering.start) && { start: numbering.start }),
              ...(numberFormat && { formatType: numberFormat as (typeof NumberFormat)[keyof typeof NumberFormat] }),
            },
          }),
        },
        titlePage: titlePage || undefined,
        ...(section.breakType && { type: BREAK_TYPES[section.breakType] }),
        ...(section.verticalAlign && { verticalAlign: VerticalAlignSection[section.verticalAlign.toUpperCase() as 'CENTER' | 'BOTTOM' | 'BOTH'] }),
        ...(columns && columns.count > 1 && {
          column: {
            count: columns.count,
            ...(columns.spacingPx !== undefined && { space: pxToTwips(columns.spacingPx) }),
            ...(columns.separator && { separate: true }),
            ...(columns.widthsPx && columns.widthsPx.length === columns.count
              ? { equalWidth: false, children: columns.widthsPx.map((w) => new Column({ width: pxToTwips(w), space: columns.spacingPx !== undefined ? pxToTwips(columns.spacingPx) : undefined })) }
              : {}),
          },
        }),
      },
      headers: header,
      footers: footer,
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
  builder.evenAndOdd = input.sections.some((s) => s.header?.differentOddEven || s.footer?.differentOddEven);
  const doc = new Document({
    title: input.title,
    ...(builder.evenAndOdd && { evenAndOddHeaderAndFooters: true }),
    numbering: numberingConfig(numberedStarts(input.sections)),
    sections: input.sections.map((s) => builder.buildSection(s)),
  });
  // toArrayBuffer, not toBuffer: JSZip's "nodebuffer" output does not exist in browsers,
  // so toBuffer made every in-browser DOCX download fail.
  const buffer = await Packer.toArrayBuffer(doc);
  let bytes: Uint8Array = new Uint8Array(buffer);
  bytes = await addNamedStyles(bytes, namedStyles);
  if (builder.notes.footnote.length > 0 || builder.notes.endnote.length > 0) {
    bytes = await addNoteParts(bytes, builder.notes, async (bodies) => {
      const blocks: DocBlock[] = bodies.map((content) => ({ type: 'paragraph', content }));
      return (await exportDocx({ title: '', sections: [{ ...input.sections[0], blocks, header: undefined, footer: undefined, columns: undefined }], namedStyles: input.namedStyles })).bytes;
    });
  }
  if (comments.threads.size > 0) {
    const unplaced = builder.unplaced();
    const lost = unplaced.reduce((n, t) => n + t.ids.length, 0);
    if (lost > 0) builder.warnings.push(`${lost} comment(s) could not be anchored (the document has no paragraph) and were not written.`);
    for (const t of unplaced) comments.threads.delete(t.thread.id);
    if (comments.threads.size > 0) bytes = await addCommentParts(bytes, buildCommentParts(comments));
  }
  return { bytes, warnings: builder.warnings };
}
