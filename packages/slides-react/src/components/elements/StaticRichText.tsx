// StaticRichText — renders a RichTextDoc as plain React nodes for idle (not
// being edited) text. Memoized on (doc, theme): a theme switch recolors theme
// slot marks with no data migration. The live ProseMirror editor (Phase 3)
// swaps in only on the actively-edited element.

import React from 'react';
import {
  resolveColor,
  resolveFont,
  type RichTextDoc,
  type RichTextParagraph,
  type RichTextInline,
  type ThemeData,
  type Color,
} from '@b32nio/spindle-slides-core';

const INDENT_PX = 28;

// URL sanitizer for rendered link hrefs. Attacker-controlled hrefs (from pasted
// or imported HTML) must not become `javascript:`/`data:`/`vbscript:` anchors.
// Kept local to avoid a cross-package coupling for one small helper.
const SAFE_HREF_SCHEMES = ['http:', 'https:', 'mailto:', 'tel:'];
// eslint-disable-next-line no-control-regex
const HREF_CONTROL_CHARS = /[\u0000-\u0020]/g;

function sanitizeHref(href: string | null | undefined): string {
  if (!href) return '';
  const value = String(href).trim();
  if (value === '') return '';
  const colon = value.indexOf(':');
  if (colon === -1) return value; // relative / anchor / query — no scheme
  const beforeSlash = value.search(/[/?#]/);
  // A ':' after the first '/', '?' or '#' is part of the path, not a scheme.
  if (beforeSlash !== -1 && beforeSlash < colon) return value;
  const scheme = value.slice(0, colon + 1).replace(HREF_CONTROL_CHARS, '').toLowerCase();
  return SAFE_HREF_SCHEMES.includes(scheme) ? value : '';
}

function inlineStyle(inline: RichTextInline, theme: ThemeData): { style: React.CSSProperties; href?: string } {
  const style: React.CSSProperties = {};
  const decorations: string[] = [];
  let href: string | undefined;

  for (const mark of inline.marks ?? []) {
    switch (mark.type) {
      case 'bold':
        style.fontWeight = 700;
        break;
      case 'italic':
        style.fontStyle = 'italic';
        break;
      case 'underline':
        decorations.push('underline');
        break;
      case 'strikethrough':
        decorations.push('line-through');
        break;
      case 'textColor':
        if (mark.attrs?.color) style.color = resolveColor(mark.attrs.color as Color, theme);
        break;
      case 'fontFamily':
        style.fontFamily = resolveFont(mark.attrs?.family as string | undefined, theme);
        break;
      case 'fontSize':
        if (typeof mark.attrs?.size === 'number') style.fontSize = `${mark.attrs.size}px`;
        break;
      case 'highlight':
        if (mark.attrs?.color) style.backgroundColor = resolveColor(mark.attrs.color as Color, theme);
        break;
      case 'link':
        href = mark.attrs?.href as string | undefined;
        if (!style.color) style.color = resolveColor({ kind: 'theme', slot: 'hlink' }, theme);
        decorations.push('underline');
        break;
    }
  }
  if (decorations.length) style.textDecoration = decorations.join(' ');
  return { style, href };
}

function renderInline(inline: RichTextInline, theme: ThemeData, key: number): React.ReactNode {
  const { style, href } = inlineStyle(inline, theme);
  const span = (
    <span key={key} style={style}>
      {inline.text}
    </span>
  );
  const safeHref = sanitizeHref(href);
  if (safeHref) {
    return (
      <a key={key} href={safeHref} style={style} target="_blank" rel="noreferrer">
        {inline.text}
      </a>
    );
  }
  return span;
}

function renderParagraph(
  p: RichTextParagraph,
  theme: ThemeData,
  key: number,
  ordinal: number
): React.ReactNode {
  const attrs = p.attrs ?? {};
  const indent = attrs.indent ?? 0;
  const style: React.CSSProperties = {
    textAlign: attrs.align ?? 'left',
    lineHeight: attrs.lineHeight ?? 1.2,
    marginTop: attrs.spaceBefore ?? 0,
    marginBottom: attrs.spaceAfter ?? 0,
    paddingLeft: indent * INDENT_PX,
  };

  const inlines = (p.content ?? []).map((inline, i) => renderInline(inline, theme, i));
  const body = inlines.length ? inlines : <br />;

  const listType = attrs.listType ?? 'none';
  if (listType === 'none') {
    return (
      <div key={key} style={style}>
        {body}
      </div>
    );
  }

  // Mirror the live editor's list layout (schema toDOM + editor stylesheet):
  // a (indent+1)*INDENT_PX gutter with the marker hanging in it, so there's no
  // indent jump when entering/leaving edit mode.
  const marker = listType === 'bullet' ? '•' : `${ordinal}.`;
  return (
    <div key={key} style={{ ...style, position: 'relative', paddingLeft: (indent + 1) * INDENT_PX }}>
      <span style={{ position: 'absolute', left: listType === 'bullet' ? 8 : 2, opacity: 0.9 }}>{marker}</span>
      {body}
    </div>
  );
}

export interface StaticRichTextProps {
  doc: RichTextDoc;
  theme: ThemeData;
}

function StaticRichTextImpl({ doc, theme }: StaticRichTextProps): React.ReactElement {
  // Per-nesting-level ordered-list counters. A numbered item continues its own
  // level's sequence and restarts any deeper levels; a non-numbered paragraph
  // breaks the list. (A single flat counter numbered nested lists continuously.)
  const counters: number[] = [];
  const children = doc.content.map((p, i) => {
    const attrs = p.attrs ?? {};
    const level = (attrs.indent as number) ?? 0;
    let ordinal = 0;
    if ((attrs.listType ?? 'none') === 'number') {
      counters[level] = (counters[level] ?? 0) + 1;
      counters.length = level + 1; // entering this level ends deeper sub-lists
      ordinal = counters[level];
    } else {
      counters.length = 0;
    }
    return renderParagraph(p, theme, i, ordinal);
  });
  return <>{children}</>;
}

export const StaticRichText = React.memo(StaticRichTextImpl);
