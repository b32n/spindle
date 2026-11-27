// StaticRichText — renders a RichTextDoc as plain React nodes for idle (not
// being edited) text. Memoized on (doc, theme): a theme switch recolors theme
// slot marks with no data migration. The live ProseMirror editor (Phase 3)
// swaps in only on the actively-edited element.

import React from 'react';
import {
  resolveColor,
  resolveFont,
  resolveParagraph,
  formatOrdinal,
  resolveRun,
  runCss,
  baselineCss,
  resolveFieldText,
  fieldMarks,
  safeFieldKind,
  inlineText,
  type FieldContext,
  type RichTextDoc,
  type RichTextParagraph,
  type RichTextInline,
  type ThemeData,
  type Color,
} from '@b32nio/spindle-slides-core';

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

/** What field runs (slide number, date) show: the slide being drawn and the moment. SlideView provides it. */
export const FieldContextValue = React.createContext<FieldContext | null>(null);

function inlineStyle(marks: RichTextInline['marks'], theme: ThemeData, scaled: boolean): { style: React.CSSProperties; href?: string } {
  const style: React.CSSProperties = {};
  const decorations: string[] = [];
  let href: string | undefined;

  for (const mark of marks ?? []) {
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
        if (typeof mark.attrs?.size === 'number') {
          // Under autofit the size routes through the body's font scale variable (same expression as the editor schema).
          style.fontSize = scaled ? `calc(${mark.attrs.size}px * var(--spindle-fs, 1))` : `${mark.attrs.size}px`;
        }
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
  Object.assign(style, runCss(resolveRun(marks)));
  return { style, href };
}

function renderInline(inline: RichTextInline, theme: ThemeData, key: number, scaled: boolean, fields: FieldContext): React.ReactNode {
  const marks = inline.type === 'field' ? fieldMarks(inline) : inline.marks;
  const { style, href } = inlineStyle(marks, theme, scaled);
  const run = resolveRun(marks);
  let content: React.ReactNode = inline.type === 'field' ? resolveFieldText(inline.attrs, fields, inlineText(inline)) : inline.text;
  // A raised or lowered run is an inner span (smaller, shifted) so the outer span's size stays the base for its em.
  if (run.baselinePct !== null) content = <span style={baselineCss(run.baselinePct)}>{content}</span>;
  const safeHref = sanitizeHref(href);
  if (safeHref) {
    return (
      <a key={key} href={safeHref} style={style} lang={run.lang ?? undefined} target="_blank" rel="noreferrer">
        {content}
      </a>
    );
  }
  return (
    <span key={key} style={style} lang={run.lang ?? undefined} {...(inline.type === 'field' ? { 'data-field': safeFieldKind(inline.attrs?.kind) } : {})}>
      {content}
    </span>
  );
}

function renderParagraph(
  p: RichTextParagraph,
  theme: ThemeData,
  key: number,
  ordinal: number,
  scaled: boolean,
  fields: FieldContext
): React.ReactNode {
  const attrs = p.attrs ?? {};
  const r = resolveParagraph(attrs);
  const lh = 'px' in r.lineHeight ? `${r.lineHeight.px}px` : r.lineHeight.mult;
  const space = (v: typeof r.spaceBefore): string | number => (v ? ('px' in v ? v.px : `${v.em}em`) : 0);
  const style: React.CSSProperties = {
    textAlign: attrs.align ?? 'left',
    lineHeight: scaled ? `calc(${lh} * var(--spindle-ls, 1))` : lh,
    marginTop: space(r.spaceBefore),
    marginBottom: space(r.spaceAfter),
    paddingLeft: r.paddingLeft,
    ...(r.textIndent ? { textIndent: r.textIndent } : {}),
  };

  const inlines = (p.content ?? []).map((inline, i) => renderInline(inline, theme, i, scaled, fields));
  const body = inlines.length ? inlines : <br />;

  const marker = r.marker;
  if (!marker) {
    return (
      <div key={key} style={style}>
        {body}
      </div>
    );
  }

  // Mirror the live editor's list layout (schema toDOM + editor stylesheet): the marker hangs in the left gutter, so
  // there's no indent jump when entering/leaving edit mode.
  const label = marker.kind === 'bullet' ? marker.char : `${formatOrdinal(ordinal, marker.scheme)}.`;
  const markerStyle: React.CSSProperties = {
    position: 'absolute',
    left: r.markerLeft,
    opacity: marker.color ? 1 : 0.9,
    ...(marker.color ? { color: resolveColor(marker.color, theme) } : {}),
    ...(marker.sizePct !== 100 ? { fontSize: `${marker.sizePct}%` } : {}),
  };
  return (
    <div key={key} style={{ ...style, position: 'relative' }}>
      <span style={markerStyle}>{label}</span>
      {body}
    </div>
  );
}

export interface StaticRichTextProps {
  doc: RichTextDoc;
  theme: ThemeData;
  /** Route font sizes and line heights through the `--spindle-fs` / `--spindle-ls` variables of an ancestor (autofit). */
  scaled?: boolean;
}

function StaticRichTextImpl({ doc, theme, scaled = false }: StaticRichTextProps): React.ReactElement {
  // Per-nesting-level ordered-list counters. A numbered item continues its own
  // level's sequence and restarts any deeper levels; a non-numbered paragraph
  // breaks the list. (A single flat counter numbered nested lists continuously.)
  const fields = React.useContext(FieldContextValue) ?? { slideNumber: null, now: new Date() };
  const counters: number[] = [];
  const children = doc.content.map((p, i) => {
    const attrs = p.attrs ?? {};
    const level = (attrs.indent as number) ?? 0;
    const marker = resolveParagraph(attrs).marker;
    let ordinal = 0;
    if (marker?.kind === 'number') {
      // A start value applies where a list begins; later items of the run keep counting (PPTX repeats it on each item).
      counters[level] = counters[level] === undefined ? (marker.startAt ?? 1) : counters[level] + 1;
      counters.length = level + 1; // entering this level ends deeper sub-lists
      ordinal = counters[level];
    } else {
      counters.length = 0;
    }
    return renderParagraph(p, theme, i, ordinal, scaled, fields);
  });
  return <>{children}</>;
}

export const StaticRichText = React.memo(StaticRichTextImpl);
