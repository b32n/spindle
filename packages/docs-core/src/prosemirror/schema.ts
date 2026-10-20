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
 * Node specifications matching our document model
 */
const nodes: Record<string, NodeSpec> = {
  doc: {
    content: 'block+',
  },

  paragraph: {
    content: 'inline*',
    group: 'block',
    attrs: {
      alignment: { default: 'left' },
      indent: { default: 0 },
      lineSpacing: { default: 'single' },
      spaceBefore: { default: 0 },
      spaceAfter: { default: 8 },
    },
    parseDOM: [{ tag: 'p' }],
    toDOM(node): DOMOutputSpec {
      // Attrs come from loaded/collab document JSON (untrusted); sanitize before
      // interpolating into the style string, or a value like
      // `left; background: url(//evil)` injects a CSS declaration.
      const align = safeCssKeyword(node.attrs.alignment) || 'left';
      const before = safeCssNumber(node.attrs.spaceBefore) ?? 0;
      const after = safeCssNumber(node.attrs.spaceAfter) ?? 8;
      const style = `text-align: ${align}; margin-top: ${before}px; margin-bottom: ${after}px;`;
      return ['p', { style }, 0];
    },
  },

  heading: {
    content: 'inline*',
    group: 'block',
    attrs: {
      level: { default: 1 },
      alignment: { default: 'left' },
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
      return [`h${lvl}`, { style: `text-align: ${align}` }, 0];
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
    attrs: { colWidths: { default: null }, styleId: { default: null } },
    parseDOM: [{ tag: 'table' }],
    toDOM(): DOMOutputSpec {
      return ['table', { class: 'docs-table' }, ['tbody', 0]];
    },
  },

  table_row: {
    content: '(table_cell | table_header)+',
    tableRole: 'row',
    attrs: { height: { default: null } },
    parseDOM: [{ tag: 'tr' }],
    toDOM(): DOMOutputSpec {
      return ['tr', 0];
    },
  },

  table_cell: {
    content: 'block+',
    attrs: {
      colspan: { default: 1 },
      rowspan: { default: 1 },
      colwidth: { default: null },
      backgroundColor: { default: null },
      styleId: { default: null },
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
    content: 'block+',
    attrs: {
      colspan: { default: 1 },
      rowspan: { default: 1 },
      colwidth: { default: null },
      backgroundColor: { default: null },
      styleId: { default: null },
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

