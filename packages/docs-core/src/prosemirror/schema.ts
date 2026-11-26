import { Schema, NodeSpec, MarkSpec, DOMOutputSpec } from 'prosemirror-model';
import {
  sanitizeHref,
  sanitizeImageSrc,
  safeCssColor,
  safeFontFamily,
  safeCssKeyword,
  safeCssNumber,
} from './sanitize';

/**
 * Paragraph formatting that paragraphs and headings carry beyond alignment and spacing: indents in pixels
 * (first-line negative for a hanging indent) and the pagination flags. All optional, absent by default.
 */
const PARAGRAPH_FORMAT_ATTRS = {
  leftIndent: { default: null },
  rightIndent: { default: null },
  firstLineIndent: { default: null },
  keepWithNext: { default: null },
  keepTogether: { default: null },
  pageBreakBefore: { default: null },
  // Named paragraph style (DOCX `w:pStyle`), a key of the document's `namedStyles`.
  namedStyleId: { default: null },
};

/**
 * Properties of a content control, on the block and the inline node alike (null/false when unset). `items` is a list of
 * `{ text, value }` choices; the rest are strings and flags that mirror DOCX's `w:sdtPr`.
 */
const CONTENT_CONTROL_ATTRS = {
  controlType: { default: 'richText' },
  tag: { default: null },
  alias: { default: null },
  lock: { default: null },
  placeholder: { default: null },
  showingPlaceholder: { default: false },
  items: { default: null },
  dateFormat: { default: null },
  dateValue: { default: null },
  checked: { default: false },
  multiline: { default: false },
  gallery: { default: null },
};

const CONTROL_TYPES = new Set(['richText', 'text', 'dropDown', 'comboBox', 'date', 'checkbox', 'picture', 'docPart']);
const CONTROL_LOCKS = new Set(['sdtLocked', 'contentLocked', 'sdtContentLocked']);

/** The attrs of a content control read back from `data-content-control` (pasted or dropped HTML is untrusted): unknown fields and wrongly typed values are dropped. */
function contentControlAttrsFromDom(dom: HTMLElement): Record<string, unknown> | false {
  let raw: unknown;
  try {
    raw = JSON.parse(dom.getAttribute('data-content-control') || '');
  } catch {
    return false;
  }
  if (!raw || typeof raw !== 'object') return false;
  const r = raw as Record<string, unknown>;
  const str = (v: unknown) => (typeof v === 'string' ? v : null);
  const items = Array.isArray(r.items)
    ? r.items.filter((i): i is { text: string; value: string } => !!i && typeof (i as { text?: unknown }).text === 'string' && typeof (i as { value?: unknown }).value === 'string').map((i) => ({ text: i.text, value: i.value }))
    : null;
  return {
    controlType: typeof r.controlType === 'string' && CONTROL_TYPES.has(r.controlType) ? r.controlType : 'richText',
    tag: str(r.tag),
    alias: str(r.alias),
    lock: typeof r.lock === 'string' && CONTROL_LOCKS.has(r.lock) ? r.lock : null,
    placeholder: str(r.placeholder),
    showingPlaceholder: r.showingPlaceholder === true,
    items,
    dateFormat: str(r.dateFormat),
    dateValue: str(r.dateValue),
    checked: r.checked === true,
    multiline: r.multiline === true,
    gallery: str(r.gallery),
  };
}

/** DOM attributes of a content control: the title, and every property as JSON so a copy pasted into the editor keeps them. */
function contentControlDom(attrs: Record<string, unknown>, tag: 'div' | 'span'): DOMOutputSpec {
  const domAttrs: Record<string, string> = {
    class: 'content-control',
    'data-content-control': JSON.stringify(Object.fromEntries(Object.keys(CONTENT_CONTROL_ATTRS).map((k) => [k, attrs[k]]))),
  };
  if (typeof attrs.alias === 'string' && attrs.alias) domAttrs.title = attrs.alias;
  return [tag, domAttrs, 0];
}

/** The displayed text of a field's cached result (a string, or text/image runs stored as JSON in the node attr). */
export function fieldResultText(result: unknown): string {
  if (typeof result === 'string') return result;
  if (!Array.isArray(result)) return '';
  return result.map((r) => (r && typeof r === 'object' && typeof (r as { text?: unknown }).text === 'string' ? (r as { text: string }).text : '')).join('');
}

/** CSS for the indents of a paragraph/heading (sanitised: the attrs come from loaded or collaborative JSON). */
function indentStyle(attrs: Record<string, unknown>): string {
  let css = '';
  const left = safeCssNumber(attrs.leftIndent as number | null);
  const right = safeCssNumber(attrs.rightIndent as number | null);
  const first = safeCssNumber(attrs.firstLineIndent as number | null);
  if (left) css += ` margin-left: ${left}px;`;
  if (right) css += ` margin-right: ${right}px;`;
  if (first) css += ` text-indent: ${first}px;`;
  return css;
}

/**
 * Node specifications matching our document model
 */
const nodes: Record<string, NodeSpec> = {
  doc: {
    // A section break only ends a section of the body, so it is allowed here and not inside cells or list items.
    // A content control wraps blocks; like a section break it is outside the `block` group so lists and quotes cannot hold one.
    content: '(block | content_control | section_break)+',
  },

  paragraph: {
    content: 'inline*',
    group: 'block',
    attrs: {
      alignment: { default: 'left' },
      indent: { default: 0 },
      // 'single', or a LineSpacing object ({ type, value? }) for anything else.
      lineSpacing: { default: 'single' },
      spaceBefore: { default: 0 },
      spaceAfter: { default: 8 },
      ...PARAGRAPH_FORMAT_ATTRS,
    },
    parseDOM: [{ tag: 'p' }],
    toDOM(node): DOMOutputSpec {
      // Attrs come from loaded/collab document JSON (untrusted); sanitize before
      // interpolating into the style string, or a value like
      // `left; background: url(//evil)` injects a CSS declaration.
      const align = safeCssKeyword(node.attrs.alignment) || 'left';
      const before = safeCssNumber(node.attrs.spaceBefore) ?? 0;
      const after = safeCssNumber(node.attrs.spaceAfter) ?? 8;
      const style = `text-align: ${align}; margin-top: ${before}px; margin-bottom: ${after}px;${indentStyle(node.attrs)}`;
      return ['p', { style }, 0];
    },
  },

  heading: {
    content: 'inline*',
    group: 'block',
    attrs: {
      level: { default: 1 },
      alignment: { default: 'left' },
      lineSpacing: { default: 'single' },
      spaceBefore: { default: null },
      spaceAfter: { default: null },
      ...PARAGRAPH_FORMAT_ATTRS,
    },
    parseDOM: [
      { tag: 'h1', attrs: { level: 1 } },
      { tag: 'h2', attrs: { level: 2 } },
      { tag: 'h3', attrs: { level: 3 } },
      { tag: 'h4', attrs: { level: 4 } },
      { tag: 'h5', attrs: { level: 5 } },
      { tag: 'h6', attrs: { level: 6 } },
    ],
    toDOM(node): DOMOutputSpec {
      const align = safeCssKeyword(node.attrs.alignment) || 'left';
      // Clamp to a valid tag number — an untrusted (collab/loaded) level that is
      // a string or out of 1-6 would make `h${level}` an invalid tag name and
      // throw InvalidCharacterError in createElement, aborting the render.
      const lvl = Math.min(6, Math.max(1, Math.trunc(Number(node.attrs.level)) || 1));
      const before = safeCssNumber(node.attrs.spaceBefore);
      const after = safeCssNumber(node.attrs.spaceAfter);
      const extras = `${before != null ? ` margin-top: ${before}px;` : ''}${after != null ? ` margin-bottom: ${after}px;` : ''}${indentStyle(node.attrs)}`;
      return [`h${lvl}`, { style: extras ? `text-align: ${align};${extras}` : `text-align: ${align}` }, 0];
    },
  },

  bullet_list: {
    content: 'list_item+',
    group: 'block',
    parseDOM: [{ tag: 'ul' }],
    toDOM(): DOMOutputSpec {
      return ['ul', 0];
    },
  },

  ordered_list: {
    content: 'list_item+',
    group: 'block',
    attrs: { start: { default: 1 } },
    parseDOM: [
      {
        tag: 'ol',
        getAttrs(dom: HTMLElement) {
          return { start: dom.hasAttribute('start') ? +dom.getAttribute('start')! : 1 };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      return node.attrs.start === 1 ? ['ol', 0] : ['ol', { start: node.attrs.start }, 0];
    },
  },

  list_item: {
    content: 'paragraph block*',
    attrs: { level: { default: 0 } }, // nesting depth; preserved across collab round-trips
    parseDOM: [
      {
        tag: 'li',
        getAttrs(dom: HTMLElement) {
          const lvl = dom.getAttribute('data-level');
          return { level: lvl ? parseInt(lvl, 10) || 0 : 0 };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      return node.attrs.level ? ['li', { 'data-level': node.attrs.level }, 0] : ['li', 0];
    },
    defining: true,
  },

  table: {
    content: 'table_row+',
    group: 'block',
    tableRole: 'table',
    isolating: true,
    // colWidths/styleId are model metadata carried across collab round-trips;
    // toDOM doesn't emit them (the painter reads them from the block model).
    attrs: { colWidths: { default: null }, styleId: { default: null }, borders: { default: null }, look: { default: null }, namedStyleId: { default: null } },
    parseDOM: [{ tag: 'table' }],
    toDOM(): DOMOutputSpec {
      return ['table', { class: 'docs-table' }, ['tbody', 0]];
    },
  },

  table_row: {
    content: '(table_cell | table_header)+',
    tableRole: 'row',
    attrs: { height: { default: null }, repeatHeader: { default: false }, cantSplit: { default: false } },
    parseDOM: [{ tag: 'tr' }],
    toDOM(): DOMOutputSpec {
      return ['tr', 0];
    },
  },

  table_cell: {
    content: '(block | content_control)+',
    attrs: {
      colspan: { default: 1 },
      rowspan: { default: 1 },
      colwidth: { default: null },
      backgroundColor: { default: null },
      styleId: { default: null },
      borders: { default: null },
      verticalAlign: { default: null },
      textDirection: { default: null },
    },
    tableRole: 'cell',
    isolating: true,
    parseDOM: [
      {
        tag: 'td',
        getAttrs(dom: HTMLElement) {
          const colspan = dom.getAttribute('colspan');
          const rowspan = dom.getAttribute('rowspan');
          const bgColor = dom.style.backgroundColor || null;
          return {
            colspan: colspan ? +colspan : 1,
            rowspan: rowspan ? +rowspan : 1,
            backgroundColor: bgColor,
          };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      const attrs: Record<string, string> = {};
      if (node.attrs.colspan !== 1) attrs.colspan = node.attrs.colspan;
      if (node.attrs.rowspan !== 1) attrs.rowspan = node.attrs.rowspan;
      const bg = safeCssColor(node.attrs.backgroundColor as string | null);
      if (bg) attrs.style = `background-color: ${bg}`;
      return ['td', attrs, 0];
    },
  },

  table_header: {
    content: '(block | content_control)+',
    attrs: {
      colspan: { default: 1 },
      rowspan: { default: 1 },
      colwidth: { default: null },
      backgroundColor: { default: null },
      styleId: { default: null },
      borders: { default: null },
      verticalAlign: { default: null },
      textDirection: { default: null },
    },
    tableRole: 'header_cell',
    isolating: true,
    parseDOM: [
      {
        tag: 'th',
        getAttrs(dom: HTMLElement) {
          const colspan = dom.getAttribute('colspan');
          const rowspan = dom.getAttribute('rowspan');
          const bgColor = dom.style.backgroundColor || null;
          return {
            colspan: colspan ? +colspan : 1,
            rowspan: rowspan ? +rowspan : 1,
            backgroundColor: bgColor,
          };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      const attrs: Record<string, string> = {};
      if (node.attrs.colspan !== 1) attrs.colspan = node.attrs.colspan;
      if (node.attrs.rowspan !== 1) attrs.rowspan = node.attrs.rowspan;
      const bg = safeCssColor(node.attrs.backgroundColor as string | null);
      if (bg) attrs.style = `background-color: ${bg}`;
      return ['th', attrs, 0];
    },
  },

  image: {
    group: 'block',
    attrs: {
      src: {},
      alt: { default: '' },
      title: { default: null },
      width: { default: null },
      height: { default: null },
      alignment: { default: 'center' },
    },
    parseDOM: [
      {
        tag: 'img[src]',
        getAttrs(dom: HTMLElement) {
          return {
            src: dom.getAttribute('src'),
            alt: dom.getAttribute('alt') || '',
            title: dom.getAttribute('title'),
            width: dom.getAttribute('width'),
            height: dom.getAttribute('height'),
          };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      const attrs: Record<string, string | null> = {
        src: sanitizeImageSrc(node.attrs.src as string | null),
        alt: node.attrs.alt,
      };
      if (node.attrs.title) attrs.title = node.attrs.title;
      if (node.attrs.width) attrs.width = node.attrs.width;
      if (node.attrs.height) attrs.height = node.attrs.height;
      return ['img', attrs];
    },
  },

  horizontal_rule: {
    group: 'block',
    parseDOM: [{ tag: 'hr' }],
    toDOM(): DOMOutputSpec {
      return ['hr'];
    },
  },

  page_break: {
    group: 'block',
    parseDOM: [{ tag: 'div.page-break' }],
    toDOM(): DOMOutputSpec {
      return ['div', { class: 'page-break' }];
    },
  },

  // The end of a section: the content before it (back to the previous break) belongs to the section whose id it
  // names. The last section of a document has no break. Deliberately has no `parseDOM`, so pasted or dropped HTML
  // can never create a break with a duplicate id; commands insert breaks with a fresh id. Not in the `block` group, so
  // it can only sit at the top level of the document.
  section_break: {
    atom: true,
    selectable: true,
    attrs: { sectionId: { default: '' } },
    toDOM(node): DOMOutputSpec {
      return ['div', { class: 'section-break', 'data-section-id': String(node.attrs.sectionId) }];
    },
  },

  // A content control (a Word structured document tag) around blocks. It only sits at the top level of the document, in
  // table cells and in other controls. Its properties are attrs; `checked` is the state of a checkbox control.
  content_control: {
    content: '(block | content_control)+',
    defining: true,
    attrs: CONTENT_CONTROL_ATTRS,
    parseDOM: [{ tag: 'div[data-content-control]', getAttrs: contentControlAttrsFromDom }],
    toDOM: (node) => contentControlDom(node.attrs, 'div'),
  },

  // A content control inside a paragraph, around inline content.
  content_control_inline: {
    inline: true,
    group: 'inline',
    content: 'inline*',
    attrs: CONTENT_CONTROL_ATTRS,
    parseDOM: [{ tag: 'span[data-content-control]', getAttrs: contentControlAttrsFromDom }],
    toDOM: (node) => contentControlDom(node.attrs, 'span'),
  },

  // A reference to a footnote or endnote. The note's text is its content, which the editor never edits in place: the node
  // is an atom (a single unit for the caret) drawn as the note's number, and the text is shown in the notes area. Only the
  // inline content a note can hold is allowed, so notes do not nest.
  note_ref: {
    inline: true,
    group: 'inline',
    atom: true,
    selectable: true,
    content: '(text | hard_break | inline_image | bookmark | field | dynamic_field)*',
    attrs: { kind: { default: 'footnote' }, noteId: { default: '' }, customMark: { default: null } },
    parseDOM: [
      {
        tag: 'sup[data-note-ref]',
        getAttrs(dom: HTMLElement) {
          const kind = dom.getAttribute('data-note-ref');
          if (kind !== 'footnote' && kind !== 'endnote') return false;
          return { kind, noteId: dom.getAttribute('data-note-id') || '', customMark: dom.getAttribute('data-custom-mark') };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      const attrs: Record<string, string> = { class: 'note-ref', 'data-note-ref': node.attrs.kind === 'endnote' ? 'endnote' : 'footnote', 'data-note-id': String(node.attrs.noteId ?? '') };
      if (typeof node.attrs.customMark === 'string' && node.attrs.customMark) attrs['data-custom-mark'] = node.attrs.customMark;
      return ['sup', attrs, 0];
    },
  },

  hard_break: {
    inline: true,
    group: 'inline',
    selectable: false,
    parseDOM: [{ tag: 'br' }],
    toDOM(): DOMOutputSpec {
      return ['br'];
    },
  },

  // Inline (in-paragraph) image. Distinct from the block-level `image` node so
  // it keeps its place in the text flow and round-trips as an InlineImage.
  inline_image: {
    inline: true,
    group: 'inline',
    atom: true,
    draggable: true,
    attrs: {
      src: {},
      alt: { default: '' },
      width: { default: null },
      height: { default: null },
    },
    parseDOM: [
      {
        tag: 'img[data-inline-image]',
        getAttrs(dom: HTMLElement) {
          const w = dom.getAttribute('width');
          const h = dom.getAttribute('height');
          return {
            src: dom.getAttribute('src'),
            alt: dom.getAttribute('alt') || '',
            width: w ? Number(w) : null,
            height: h ? Number(h) : null,
          };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      const attrs: Record<string, string | number | null> = {
        'data-inline-image': 'true',
        src: sanitizeImageSrc(node.attrs.src as string | null),
        alt: node.attrs.alt,
      };
      if (node.attrs.width) attrs.width = node.attrs.width;
      if (node.attrs.height) attrs.height = node.attrs.height;
      return ['img', attrs];
    },
  },

  // Dynamic field (page number, total pages, date, ...) embedded in the text.
  // An atom: its value is resolved at render time from `fieldType`/`format`.
  dynamic_field: {
    inline: true,
    group: 'inline',
    atom: true,
    selectable: true,
    attrs: {
      fieldType: { default: 'pageNumber' },
      format: { default: null },
    },
    parseDOM: [
      {
        tag: 'span[data-dynamic-field]',
        getAttrs(dom: HTMLElement) {
          return {
            fieldType: dom.getAttribute('data-dynamic-field') || 'pageNumber',
            format: dom.getAttribute('data-format'),
          };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      const attrs: Record<string, string> = {
        class: 'dynamic-field',
        'data-dynamic-field': String(node.attrs.fieldType),
      };
      if (node.attrs.format) attrs['data-format'] = String(node.attrs.format);
      return ['span', attrs, `\u2039${String(node.attrs.fieldType)}\u203A`];
    },
  },

  // Field the editor does not evaluate (REF, SEQ, STYLEREF, ...). An atom that
  // displays its cached `result` (a string, or text/image runs as JSON) and
  // keeps the `instruction` so it can be written back as a field.
  field: {
    inline: true,
    group: 'inline',
    atom: true,
    selectable: true,
    attrs: {
      instruction: { default: '' },
      kind: { default: '' },
      result: { default: '' },
      dirty: { default: false },
    },
    parseDOM: [
      {
        tag: 'span[data-field]',
        getAttrs(dom: HTMLElement) {
          return {
            kind: dom.getAttribute('data-field') || '',
            instruction: dom.getAttribute('data-instruction') || '',
            result: dom.textContent || '',
            dirty: dom.getAttribute('data-dirty') === 'true',
          };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      const attrs: Record<string, string> = {
        class: 'field',
        'data-field': String(node.attrs.kind),
        'data-instruction': String(node.attrs.instruction),
      };
      if (node.attrs.dirty) attrs['data-dirty'] = 'true';
      return ['span', attrs, fieldResultText(node.attrs.result)];
    },
  },

  // A named position (Word bookmark), the target of internal links. Zero-width in the
  // document flow; the editor sees it as one position.
  bookmark: {
    inline: true,
    group: 'inline',
    atom: true,
    selectable: false,
    attrs: { name: { default: '' } },
    parseDOM: [
      {
        tag: 'span[data-bookmark]',
        getAttrs(dom: HTMLElement) {
          return { name: dom.getAttribute('data-bookmark') || '' };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      return ['span', { class: 'bookmark', 'data-bookmark': String(node.attrs.name) }];
    },
  },

  text: {
    group: 'inline',
  },
};

/**
 * Mark specifications for inline formatting
 */
const marks: Record<string, MarkSpec> = {
  bold: {
    parseDOM: [
      { tag: 'strong' },
      { tag: 'b' },
      {
        style: 'font-weight',
        getAttrs: (value: string) => /^(bold(er)?|[5-9]\d{2,})$/.test(value) && null,
      },
    ],
    toDOM(): DOMOutputSpec {
      return ['strong', 0];
    },
  },

  italic: {
    parseDOM: [
      { tag: 'em' },
      { tag: 'i' },
      { style: 'font-style=italic' },
    ],
    toDOM(): DOMOutputSpec {
      return ['em', 0];
    },
  },

  underline: {
    parseDOM: [
      { tag: 'u' },
      {
        style: 'text-decoration',
        getAttrs: (value: string) => value.includes('underline') && null,
      },
    ],
    toDOM(): DOMOutputSpec {
      return ['u', 0];
    },
  },

  strikethrough: {
    parseDOM: [
      { tag: 's' },
      { tag: 'strike' },
      { tag: 'del' },
      {
        style: 'text-decoration',
        getAttrs: (value: string) => value.includes('line-through') && null,
      },
    ],
    toDOM(): DOMOutputSpec {
      return ['s', 0];
    },
  },

  superscript: {
    excludes: 'subscript',
    parseDOM: [{ tag: 'sup' }],
    toDOM(): DOMOutputSpec {
      return ['sup', 0];
    },
  },

  subscript: {
    excludes: 'superscript',
    parseDOM: [{ tag: 'sub' }],
    toDOM(): DOMOutputSpec {
      return ['sub', 0];
    },
  },

  link: {
    attrs: {
      href: {},
      title: { default: null },
      target: { default: '_blank' },
    },
    inclusive: false,
    parseDOM: [
      {
        tag: 'a[href]',
        getAttrs(dom: HTMLElement) {
          return {
            href: sanitizeHref(dom.getAttribute('href')),
            title: dom.getAttribute('title'),
            target: dom.getAttribute('target'),
          };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      const attrs: Record<string, string | null> = {
        href: sanitizeHref(node.attrs.href as string | null),
        rel: 'noopener noreferrer',
      };
      if (node.attrs.title) attrs.title = node.attrs.title;
      if (node.attrs.target) attrs.target = node.attrs.target;
      return ['a', attrs, 0];
    },
  },

  textStyle: {
    attrs: {
      color: { default: null },
      backgroundColor: { default: null },
      fontSize: { default: null },
      fontFamily: { default: null },
      // East Asian and complex-script fonts (only when they differ from fontFamily).
      fontFamilyEastAsia: { default: null },
      fontFamilyComplex: { default: null },
    },
    parseDOM: [
      {
        tag: 'span',
        getAttrs(dom: HTMLElement) {
          const style = dom.style;
          return {
            color: style.color || null,
            backgroundColor: style.backgroundColor || null,
            fontSize: style.fontSize ? parseInt(style.fontSize) : null,
            fontFamily: style.fontFamily || null,
          };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      const styles: string[] = [];
      const color = safeCssColor(node.attrs.color as string | null);
      if (color) styles.push(`color: ${color}`);
      const bg = safeCssColor(node.attrs.backgroundColor as string | null);
      if (bg) styles.push(`background-color: ${bg}`);
      const fontSize = safeCssNumber(node.attrs.fontSize);
      if (fontSize) styles.push(`font-size: ${fontSize}pt`);
      const font = safeFontFamily(node.attrs.fontFamily as string | null);
      if (font) styles.push(`font-family: ${font}`);
      return ['span', { style: styles.join('; ') }, 0];
    },
  },

  smallCaps: {
    parseDOM: [
      {
        style: 'font-variant',
        getAttrs: (value: string) => value === 'small-caps' && null,
      },
    ],
    toDOM(): DOMOutputSpec {
      return ['span', { style: 'font-variant: small-caps' }, 0];
    },
  },

  allCaps: {
    parseDOM: [
      {
        style: 'text-transform',
        getAttrs: (value: string) => value === 'uppercase' && null,
      },
    ],
    toDOM(): DOMOutputSpec {
      return ['span', { style: 'text-transform: uppercase' }, 0];
    },
  },

  // Hidden text (DOCX `vanish`). Kept in the document and exported; the editor shows it dimmed rather than
  // removing it from the layout, so every character stays reachable by the caret.
  hidden: {
    parseDOM: [{ tag: 'span[data-hidden-text]' }],
    toDOM(): DOMOutputSpec {
      return ['span', { 'data-hidden-text': 'true', style: 'opacity: 0.5' }, 0];
    },
  },

  // Named character style (DOCX `w:rStyle`); `id` is a key of the document's `namedStyles`. The run's formatting
  // marks hold the effective formatting; this mark keeps the style's name so it round-trips and shows in Word.
  namedStyle: {
    attrs: { id: {} },
    parseDOM: [
      {
        tag: 'span[data-named-style]',
        getAttrs(dom: HTMLElement) {
          return { id: dom.getAttribute('data-named-style') };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      return ['span', { 'data-named-style': String(node.attrs.id) }, 0];
    },
  },

  // Anchors a comment thread to a span of text. `threadId` points at an
  // entry in the document's comment-thread store. `inclusive: false` so
  // typing at either edge of the range does not extend the comment.
  comment: {
    attrs: {
      threadId: {},
    },
    inclusive: false,
    parseDOM: [
      {
        tag: 'span[data-comment-thread]',
        getAttrs(dom: HTMLElement) {
          return { threadId: dom.getAttribute('data-comment-thread') };
        },
      },
    ],
    toDOM(node): DOMOutputSpec {
      return ['span', { 'data-comment-thread': node.attrs.threadId as string }, 0];
    },
  },
};

/**
 * The main ProseMirror schema for the docs editor
 */
export const docsSchema = new Schema({ nodes, marks });

/**
 * Export individual node and mark specs for extension
 */
export { nodes as nodeSpecs, marks as markSpecs };

