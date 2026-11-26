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
  TabStopType,
  LeaderType,
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
  DocParagraphFormat,
  DocRevision,
  DocRevisionMeta,
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

/**
 * The entry of a table of contents made in the editor is a link to a heading, a tab and the page number as text. Word's
 * tables give the page number as a `PAGEREF` field of the heading's bookmark, which it can refresh: the number after the last tab
 * becomes one (an entry that already has fields, or no link to a bookmark, is left as it is).
 */
function withPageReferences(content: DocInline[]): DocInline[] {
  if (content.some((c) => c.type === 'field')) return content;
  const anchor = content.find((c): c is Extract<DocInline, { type: 'link' }> => c.type === 'link' && c.href.startsWith('#') && c.href.length > 1)?.href.slice(1);
  const last = content.length - 1;
  const tail = content[last];
  const tab = tail && tail.type === 'text' ? tail.text.lastIndexOf('\t') : -1;
  if (!anchor || !tail || tail.type !== 'text' || tab < 0 || tab === tail.text.length - 1) return content;
  const { text, ...run } = tail;
  const before = text.slice(0, tab + 1);
  return [...content.slice(0, last), { ...run, type: 'text', text: before }, { type: 'field', instruction: `PAGEREF ${anchor} \\h`, kind: 'PAGEREF', result: text.slice(tab + 1) }];
}

/** The inline content with note references taken out (a note holds no notes). */
function withoutNoteRefs(content: DocInline[]): DocInline[] {
  return content.flatMap((item): DocInline[] => {
    if (item.type === 'noteRef') return [];
    return item.type === 'contentControl' ? [{ ...item, content: withoutNoteRefs(item.content) }] : [item];
  });
}

/** The author and date of a change can come from a collaborator: keep characters XML cannot carry out of the file. */
// eslint-disable-next-line no-control-regex
const XML_INVALID_UNIT = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/;
const isHigh = (code: number) => code >= 0xd800 && code <= 0xdbff;
const isLow = (code: number) => code >= 0xdc00 && code <= 0xdfff;
/** Text without control characters and without unpaired surrogates (no look-behind: older Safari cannot parse one). */
function xmlText(text: string): string {
  let out = '';
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (isHigh(code) && i + 1 < text.length && isLow(text.charCodeAt(i + 1))) {
      out += text[i] + text[i + 1];
      i++;
    } else if (!isHigh(code) && !isLow(code) && !XML_INVALID_UNIT.test(text[i])) {
      out += text[i];
    }
  }
  return out.trim();
}
const ISO_DATE = /^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2})(:\d{2}(?:\.\d+)?)?(Z|[+-]\d{2}:\d{2})?$/;
/** `w:date` is an xsd:dateTime, which needs the seconds: a date without them gets `:00`; anything else is dropped. */
const wordDate = (date: string | undefined): string | undefined => {
  const match = date ? ISO_DATE.exec(date) : null;
  if (!match) return undefined;
  const [, day, hour, minute] = /^\d{4}-(\d{2}-\d{2})T(\d{2}):(\d{2})/.exec(match[1]) as RegExpExecArray;
  const [month, dayOfMonth] = day.split('-').map(Number);
  const seconds = match[2] ? Number(match[2].slice(1, 3)) : 0;
  if (month < 1 || month > 12 || dayOfMonth < 1 || dayOfMonth > 31 || Number(hour) > 23 || Number(minute) > 59 || seconds > 59) return undefined;
  return `${match[1]}${match[2] ?? ':00'}${match[3] ?? ''}`;
};

/** True when the run is part of a tracked insertion or deletion. */
function hasRevision(revision: DocRevision | undefined): revision is DocRevision {
  return !!revision && (!!revision.insertion || !!revision.deletion);
}

function sameMeta(a: DocRevisionMeta | undefined, b: DocRevisionMeta | undefined): boolean {
  return a === b || (!!a && !!b && a.id === b.id && a.author === b.author && a.date === b.date && a.moveId === b.moveId);
}

function sameRevision(a: DocRevision, b: DocRevision): boolean {
  return sameMeta(a.insertion, b.insertion) && sameMeta(a.deletion, b.deletion);
}

class DocxBuilder {
  readonly warnings: string[] = [];
  /** The notes referenced so far, in document order (a note's id in the file is its position + 1). */
  readonly notes: WrittenNotes = { footnote: [], endnote: [] };
  /** `w:id` of the next content control. */
  private nextControlId = 1000001;
  /** `w:id` written for each tracked change (by kind and id in the model), the ids taken, and the next free number. */
  private readonly revisionIds = new Map<string, number>();
  readonly usedRevisionIds = new Set<number>();
  private nextRevisionId = 1;
  /** Width of the content area of the section being built, in twips: where the right tab of a table of contents entry goes. */
  private contentWidthTwips = 9360;
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
    // Consecutive text runs of one tracked change are written inside one w:ins / w:del.
    let group: { revision: DocRevision; runs: ParagraphChild[] } | undefined;
    const flushGroup = () => {
      if (group) out.push(...this.wrapRevision(group.revision, group.runs));
      group = undefined;
    };
    for (const inline of content) {
      const planned =
        anchorComments && (inline.type === 'text' || inline.type === 'link') && inline.commentThreadId
          ? this.comments.threads.get(inline.commentThreadId)
          : undefined;
      // A comment's range runs from its first tagged run to its last; replies share the range of their root.
      if (planned && planned.seen === 0) {
        flushGroup();
        out.push(...this.rangeStarts(planned));
      }
      if (inline.type === 'text' && hasRevision(inline.revision)) {
        if (group && !sameRevision(group.revision, inline.revision)) flushGroup();
        group ??= { revision: inline.revision, runs: [] };
        group.runs.push(...this.textRuns(inline, inline.text, {}, !!inline.revision.deletion));
      } else {
        flushGroup();
        out.push(...this.buildRun(inline).filter((r): r is ParagraphChild => r != null));
      }
      if (planned && ++planned.seen === planned.total) {
        flushGroup();
        out.push(...this.rangeEnds(planned));
      }
    }
    flushGroup();
    return out;
  }

  /**
   * The `w:id` a tracked change is written with: its own id when that is a number not used yet by another change, else the next
   * free number. A change split over several elements (by a comment range, a non-text inline, or runs of different formatting
   * for a formatting change) is written with the same `w:id` on each, so it reads back as one change; Word accepts repeated ids.
   * A change is identified by kind, author and id, so two authors' changes with the same id stay two changes.
   */
  private revisionId(kind: string, meta: { id: string; author?: string }): number {
    const key = `${kind}:${meta.author ?? ''}:${meta.id}`;
    const known = this.revisionIds.get(key);
    if (known !== undefined) return known;
    const wanted = /^\d{1,9}$/.test(meta.id) ? Number(meta.id) : -1;
    const id = wanted >= 0 && !this.usedRevisionIds.has(wanted) ? wanted : this.freshRevisionId();
    this.usedRevisionIds.add(id);
    this.revisionIds.set(key, id);
    return id;
  }

  /** The next `w:id` nobody has used, for the markers of a move's range. */
  private freshRevisionId(): number {
    while (this.usedRevisionIds.has(this.nextRevisionId)) this.nextRevisionId++;
    this.usedRevisionIds.add(this.nextRevisionId);
    return this.nextRevisionId++;
  }

  /** A change in the form the `docx` library takes (`w:id`, `w:author`, `w:date`); a change without a date is written without one. */
  private change(kind: string, meta: DocRevisionMeta) {
    return { id: this.revisionId(kind, meta), author: xmlText(meta.author) || 'Unknown', date: wordDate(meta.date) as string };
  }

  /** An element `name` with the attributes of a tracked change (and any `extra`), around `children`. */
  private changeElement(name: string, kind: string, meta: DocRevisionMeta, children: unknown[], extra: Record<string, string> = {}): BuilderElement {
    const date = wordDate(meta.date);
    const attributes = {
      id: { key: 'w:id', value: String(this.revisionId(kind, meta)) },
      ...Object.fromEntries(Object.entries(extra).map(([k, v]) => [k, { key: k, value: v }])),
      author: { key: 'w:author', value: xmlText(meta.author) || 'Unknown' },
      ...(date !== undefined && { date: { key: 'w:date', value: date } }),
    };
    return new BuilderElement({ name, attributes, children: children as BuilderElement[] });
  }

  /**
   * `runs` inside the tracked changes of `revision`: a deletion directly around the runs, an insertion around that. One end
   * of a move is `w:moveFrom` (a deletion) or `w:moveTo` (an insertion) between the start and end of the move's range.
   */
  private wrapRevision(revision: DocRevision, runs: ParagraphChild[]): ParagraphChild[] {
    let wrapped: unknown[] = runs;
    const wrap = (meta: DocRevisionMeta, plain: 'w:ins' | 'w:del', moved: 'w:moveTo' | 'w:moveFrom') => {
      if (!meta.moveId) {
        wrapped = [this.changeElement(plain, plain === 'w:ins' ? 'ins' : 'del', meta, wrapped)];
        return;
      }
      const range = moved === 'w:moveTo' ? 'w:moveToRange' : 'w:moveFromRange';
      const rangeId = this.freshRevisionId();
      const date = wordDate(meta.date);
      const start = new BuilderElement({
        name: `${range}Start`,
        attributes: {
          id: { key: 'w:id', value: String(rangeId) },
          name: { key: 'w:name', value: xmlText(meta.moveId) },
          author: { key: 'w:author', value: xmlText(meta.author) || 'Unknown' },
          ...(date !== undefined && { date: { key: 'w:date', value: date } }),
        },
      });
      const end = new BuilderElement({ name: `${range}End`, attributes: { id: { key: 'w:id', value: String(rangeId) } } });
      wrapped = [start, this.changeElement(moved, moved === 'w:moveTo' ? 'moveTo' : 'moveFrom', meta, wrapped), end];
    };
    if (revision.deletion) wrap(revision.deletion, 'w:del', 'w:moveFrom');
    if (revision.insertion) wrap(revision.insertion, 'w:ins', 'w:moveTo');
    return wrapped as ParagraphChild[];
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
  private textRuns(base: DocRunBase, text: string, extra: Partial<IRunOptions> = {}, deleted = false): TextRun[] {
    if (!deleted && !/[\t\n]/.test(text)) return [new TextRun(this.runOptions(base, { ...extra, text }))];
    const children: (string | Tab | BuilderElement)[] = [];
    for (const part of text.split(/(\t|\n)/)) {
      if (part === '\t') children.push(new Tab());
      else if (part === '\n') children.push(new BuilderElement({ name: 'w:br' }));
      // Inside a tracked deletion the text of a run is w:delText.
      else if (part && deleted) children.push(new BuilderElement({ name: 'w:delText', attributes: { space: { key: 'xml:space', value: 'preserve' } }, children: [part] as unknown as BuilderElement[] }));
      else if (part) children.push(part);
    }
    return [new TextRun(this.runOptions(base, { ...extra, children }))];
  }

  /** The runs inside a hyperlink; its tracked changes go inside the link (a w:ins cannot hold a w:hyperlink). */
  private linkRuns(link: Extract<DocInline, { type: 'link' }>, extra: Partial<IRunOptions> = {}): ParagraphChild[] {
    const runs = this.textRuns(link, link.text, extra, !!link.revision?.deletion);
    return hasRevision(link.revision) ? this.wrapRevision(link.revision, runs) : runs;
  }

  private buildRun(inline: DocInline): (ParagraphChild | null)[] {
    switch (inline.type) {
      case 'text':
        return hasRevision(inline.revision) ? this.wrapRevision(inline.revision, this.textRuns(inline, inline.text, {}, !!inline.revision.deletion)) : this.textRuns(inline, inline.text);
      case 'bookmark':
        return inline.name ? [new Bookmark({ id: inline.name, children: [] })] : [];
      case 'link':
        // `#name` points at a bookmark in this document.
        if (inline.href.startsWith('#') && inline.href.length > 1) {
          return [
            new InternalHyperlink({
              anchor: inline.href.slice(1),
              // No Hyperlink character style: Word's own internal links (tables of contents) look like their text.
              children: this.linkRuns(inline),
            }),
          ];
        }
        return [
          new ExternalHyperlink({
            link: inline.href,
            children: this.linkRuns(inline, { style: 'Hyperlink' }),
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
    const children: unknown[] = this.buildBlocks(block.blocks);
    if (children.length === 0) children.push(this.buildParagraph({}, []));
    this.inNumberedRun = false;
    // The writer's block builders return docx paragraphs and tables, which the section and cell take as children; the wrapper is another XML element of the same kind.
    return this.buildControl(block.control, children) as unknown as Paragraph;
  }

  /** A field character run (`begin`, `separate` or `end`); `dirty` asks the application that opens the file to recompute the field. */
  private fieldChar(type: 'begin' | 'separate' | 'end', dirty = false): TextRun {
    return new TextRun({
      children: [
        new BuilderElement({
          name: 'w:fldChar',
          attributes: {
            type: { key: 'w:fldCharType', value: type },
            ...(type === 'begin' && dirty && { dirty: { key: 'w:dirty', value: 'true' } }),
          },
        }),
      ],
    });
  }

  /** The run that holds a field's instruction. */
  private instructionRun(instruction: string): TextRun {
    return new TextRun({
      children: [
        new BuilderElement({
          name: 'w:instrText',
          attributes: { space: { key: 'xml:space', value: 'preserve' } },
          // The xml formatter writes string children as the element's text.
          children: [` ${instruction} `] as unknown as BuilderElement[],
        }),
      ],
    });
  }

  /** A complex field: begin, instruction, separate, the cached result runs, end, so Word can refresh it. */
  private buildField(field: DocFieldRun): ParagraphChild[] {
    const result: ParagraphChild[] =
      typeof field.result === 'string'
        ? field.result
          ? this.textRuns({}, field.result)
          : []
        : field.result.flatMap((r) =>
            r.type === 'text' ? this.textRuns(r, r.text) : [this.buildImage(r.src, r.width, r.height, r.alt)].filter((c): c is ParagraphChild => c != null)
          );
    return [this.fieldChar('begin', field.dirty), this.instructionRun(field.instruction), this.fieldChar('separate'), ...result, this.fieldChar('end')];
  }

  private runOptions(base: DocRunBase, extra: Partial<IRunOptions>): IRunOptions {
    const style = this.namedStyle(base.namedStyleId, 'character');
    const formatChange = base.revision?.formatChange;
    return {
      ...extra,
      // The formatting the run had before a tracked change of it (`w:rPrChange`).
      ...(formatChange && { revision: { ...this.change('fmt', formatChange), ...this.runOptions(formatChange.previous, {}) } }),
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

  /** The paragraph options a paragraph's formatting gives (also used for the formatting before a tracked change). */
  private paragraphOptions(style: DocParagraphFormat): Record<string, unknown> {
    const custom = style.lineSpacingCustom;
    const spacingLine =
      custom != null
        ? custom.rule === 'multiple'
          ? { line: Math.round(custom.value * 240), lineRule: LineRuleType.AUTO }
          : { line: ptToTwips(custom.value), lineRule: custom.rule === 'exactly' ? LineRuleType.EXACT : LineRuleType.AT_LEAST }
        : { line: style.lineSpacing ? LINE_SPACING_MAP[style.lineSpacing] : undefined };
    const firstLine = style.firstLineIndentPt;
    const namedStyle = this.namedStyle(style.namedStyleId, 'paragraph');
    return {
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
    };
  }

  buildParagraph(style: DocParagraphStyle, children: ParagraphChild[], extra: Record<string, unknown> = {}): Paragraph {
    const revision = style.revision;
    const previous = revision?.formatChange?.previous;
    return new Paragraph({
      children,
      ...this.paragraphOptions(style),
      // The paragraph mark inserted or deleted (`w:rPr` of the paragraph mark), and the formatting before a tracked change of the paragraph (`w:pPrChange`).
      ...((revision?.insertion || revision?.deletion) && {
        run: {
          ...(revision.insertion && { insertion: this.change('pins', revision.insertion) }),
          ...(revision.deletion && { deletion: this.change('pdel', revision.deletion) }),
        },
      }),
      ...(revision?.formatChange && previous && {
        revision: {
          ...this.change('pfmt', revision.formatChange),
          ...this.paragraphOptions(previous),
          ...(previous.headingLevel && !this.namedStyle(previous.namedStyleId, 'paragraph') && previous.headingLevel >= 1 && previous.headingLevel <= 6 && { heading: HEADING_MAP[previous.headingLevel as 1 | 2 | 3 | 4 | 5 | 6] }),
        },
      }),
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

  /** The blocks as docx elements. A table of contents is a run of paragraphs (the field spans them). */
  buildBlocks(blocks: DocBlock[]): Array<Paragraph | Table> {
    return blocks.flatMap((b) => (b.type === 'toc' ? this.tocParagraphs(b) : [this.buildBlock(b)]));
  }

  /** The paragraphs of a table of contents: its entries, with the field starting in the first and ending in a paragraph of its own. */
  private tocParagraphs(block: Extract<DocBlock, { type: 'toc' }>): Paragraph[] {
    const entries = block.blocks.filter((b): b is Extract<DocBlock, { type: 'paragraph' }> => b.type === 'paragraph');
    const start = [this.fieldChar('begin'), this.instructionRun(block.instruction), this.fieldChar('separate')];
    // The page numbers are right-aligned under a dotted leader, as in Word's own tables of contents.
    const tabs = { tabStops: [{ type: TabStopType.RIGHT, position: this.contentWidthTwips, leader: LeaderType.DOT }] };
    const paragraphs = entries.map((entry, i) =>
      this.buildParagraph(entry, [...(i === 0 ? start : []), ...this.buildRuns(withPageReferences(entry.content))], tabs)
    );
    if (entries.length === 0) paragraphs.push(this.buildParagraph({}, start));
    paragraphs.push(new Paragraph({ children: [this.fieldChar('end')] }));
    return paragraphs;
  }

  buildBlock(block: Exclude<DocBlock, { type: 'toc' }>): Paragraph | Table {
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
    const children = this.buildBlocks(cell.blocks);
    this.inNumberedRun = false;
    if (endsWithTable(cell.blocks[cell.blocks.length - 1])) children.push(new Paragraph({}));
    return children;
  }

  private buildTable(block: Extract<DocBlock, { type: 'table' }>): Table {
    const { rows, colWidthsPx } = block;
    const tableRows = rows.map(
      (row) =>
        new TableRow({
          ...(row.revision?.insertion && { insertion: this.change('rins', row.revision.insertion) }),
          ...(row.revision?.deletion && { deletion: this.change('rdel', row.revision.deletion) }),
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
    this.contentWidthTwips = Math.max(1, pxToTwips(section.pageWidthPx - section.marginLeftPx - section.marginRightPx));
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
      children: this.buildBlocks(section.blocks),
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
  return exportDocxPart(input, new Set());
}

/** `reservedRevisionIds`: the `w:id`s of tracked changes already written in other parts, which this part must not reuse (it adds its own). */
async function exportDocxPart(input: DocumentExportInput, reservedRevisionIds: Set<number>): Promise<DocumentExportResult> {
  const comments = planComments(input.comments, input.sections);
  const namedStyles = usableNamedStyles(input.namedStyles);
  const builder = new DocxBuilder(comments, namedStyles);
  for (const id of reservedRevisionIds) builder.usedRevisionIds.add(id);
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
  for (const id of builder.usedRevisionIds) reservedRevisionIds.add(id);
  if (builder.notes.footnote.length > 0 || builder.notes.endnote.length > 0) {
    bytes = await addNoteParts(bytes, builder.notes, async (bodies) => {
      const blocks: DocBlock[] = bodies.map((content) => ({ type: 'paragraph', content }));
      return (await exportDocxPart({ title: '', sections: [{ ...input.sections[0], blocks, header: undefined, footer: undefined, columns: undefined }], namedStyles: input.namedStyles }, reservedRevisionIds)).bytes;
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
