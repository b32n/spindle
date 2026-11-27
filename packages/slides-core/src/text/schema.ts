// The ProseMirror schema for slide text bodies. Deliberately flat —
// doc > paragraph+ > inline* — because PPTX text bodies are flat paragraphs
// with bullet attrs, not nested lists. The schema's serialized JSON is exactly
// the RichTextDoc shape in model.ts, so the live editor and the static
// renderer agree.
//
// Marks that carry a theme-aware Color (textColor) or a symbolic font
// (fontFamily 'major'/'minor') store the symbolic value in attrs; toDOM does a
// best-effort literal render for the editor (theme resolution happens in the
// static renderer, which has the theme). Editing a theme-colored run therefore
// shows a fallback color, by design.

import { Schema } from 'prosemirror-model';
import type { Color } from '../scene/types';
import type { ParagraphAttrs } from './model';
import { resolveParagraph, type ResolvedParagraph } from './paragraph';
import {
  sanitizeHref,
  safeCssColor,
  safeFontFamily,
  safeThemeSlot,
  safeCssKeyword,
} from './sanitize';

// Resolve a model Color to a sanitized CSS value for the editor's best-effort
// render. A literal hex goes through the color allowlist; a theme color becomes
// a `var(--slot-…)` reference, but only for a recognized slot — both guard
// against a crafted `hex`/`slot` in imported or collaborator JSON injecting CSS.
function colorToCss(c: Color | null): string | undefined {
  if (!c) return undefined;
  if (c.kind === 'rgb') return safeCssColor(c.hex);
  const slot = safeThemeSlot(c.slot);
  return slot ? `var(--slot-${slot})` : undefined;
}

export const slidesSchema = new Schema({
  nodes: {
    doc: { content: 'paragraph+' },
    paragraph: {
      content: 'inline*',
      group: 'block',
      attrs: {
        align: { default: 'left' },
        listType: { default: 'none' },
        indent: { default: 0 },
        lineHeight: { default: null },
        spaceBefore: { default: 0 },
        spaceAfter: { default: 0 },
        // K3b: absent (null) attrs are not written to the collaboration layout, so existing text is unchanged.
        bullet: { default: null },
        lineHeightPx: { default: null },
        spaceBeforePct: { default: null },
        spaceAfterPct: { default: null },
        marginLeft: { default: null },
        firstLineIndent: { default: null },
        tabStops: { default: null },
      },
      parseDOM: [{ tag: 'p' }],
      toDOM(node) {
        const a = node.attrs as ParagraphAttrs;
        // Attrs come from imported/collaborator JSON (untrusted); resolveParagraph coerces every number and allowlists
        // every keyword, and the colour goes through colorToCss, or a value like `left;background:url(//evil)` would
        // inject a CSS declaration into the style string.
        const r = resolveParagraph(a);
        const align = safeCssKeyword(a.align) || 'left';
        const m = r.marker;
        const lineHeight = 'px' in r.lineHeight ? `${r.lineHeight.px}px` : String(r.lineHeight.mult);
        const space = (v: ResolvedParagraph['spaceBefore']) => (v ? ('px' in v ? `${v.px}px` : `${v.em}em`) : '');
        const style =
          `text-align:${align};position:relative;` +
          // Always emit line-height (default 1.2) so it matches StaticRichText — otherwise the editor inherits the
          // browser's `normal` and the text shifts a hair when entering/leaving edit.
          `line-height:calc(${lineHeight} * var(--spindle-ls, 1));` +
          (r.spaceBefore ? `margin-top:${space(r.spaceBefore)};` : '') +
          (r.spaceAfter ? `margin-bottom:${space(r.spaceAfter)};` : '') +
          (r.paddingLeft ? `padding-left:${r.paddingLeft}px;` : '') +
          (r.textIndent ? `text-indent:${r.textIndent}px;` : '') +
          (m ? `--pm-marker-left:${r.markerLeft}px;--pm-marker-size:${m.sizePct}%;` : '') +
          (m && m.color && colorToCss(m.color) ? `--pm-marker-color:${colorToCss(m.color)};` : '') +
          (m && m.kind === 'number' && m.startAt !== null ? `--pm-start:${m.startAt};` : '');
        const attrs: Record<string, string> = { style, 'data-list': m ? m.kind : 'none' };
        // The glyph and scheme travel as attributes (the editor stylesheet reads them), never inside CSS text.
        if (m && m.kind === 'bullet') attrs['data-bullet-char'] = m.char;
        if (m && m.kind === 'number') attrs['data-scheme'] = m.scheme;
        return ['p', attrs, 0];
      },
    },
    text: { group: 'inline' },
  },
  marks: {
    bold: {
      parseDOM: [{ tag: 'strong' }, { tag: 'b' }, { style: 'font-weight=bold' }],
      toDOM: () => ['strong', 0],
    },
    italic: {
      parseDOM: [{ tag: 'em' }, { tag: 'i' }, { style: 'font-style=italic' }],
      toDOM: () => ['em', 0],
    },
    underline: {
      parseDOM: [{ tag: 'u' }, { style: 'text-decoration=underline' }],
      toDOM: () => ['span', { style: 'text-decoration:underline' }, 0],
    },
    strikethrough: {
      parseDOM: [{ tag: 's' }, { tag: 'del' }, { style: 'text-decoration=line-through' }],
      toDOM: () => ['span', { style: 'text-decoration:line-through' }, 0],
    },
    link: {
      attrs: { href: { default: '' } },
      inclusive: false,
      parseDOM: [{ tag: 'a[href]', getAttrs: (dom) => ({ href: sanitizeHref((dom as HTMLElement).getAttribute('href')) }) }],
      // hlink slot resolved via the editor's CSS variable (see RichTextEditor).
      toDOM: (mark) => ['a', { href: sanitizeHref(mark.attrs.href as string), rel: 'noreferrer', target: '_blank', style: 'color:var(--slot-hlink);text-decoration:underline' }, 0],
    },
    // Theme slots and 'major'/'minor' fonts can't be resolved in a
    // theme-agnostic schema, so toDOM emits CSS variables that the editor
    // container binds to the current theme — keeping edit mode's colors/fonts
    // consistent with StaticRichText.
    textColor: {
      attrs: { color: { default: null } },
      parseDOM: [{ style: 'color', getAttrs: (v) => ({ color: { kind: 'rgb', hex: v as string } }) }],
      toDOM: (mark) => {
        const css = colorToCss(mark.attrs.color as Color | null);
        if (!css) return ['span', {}, 0];
        return ['span', { style: `color:${css}` }, 0];
      },
    },
    fontFamily: {
      attrs: { family: { default: 'minor' } },
      parseDOM: [{ style: 'font-family', getAttrs: (v) => ({ family: v as string }) }],
      toDOM: (mark) => {
        const fam = mark.attrs.family as string;
        const css =
          fam === 'major' ? 'var(--font-major)'
          : fam === 'minor' ? 'var(--font-minor)'
          : safeFontFamily(fam);
        if (!css) return ['span', {}, 0];
        return ['span', { style: `font-family:${css}` }, 0];
      },
    },
    fontSize: {
      attrs: { size: { default: 18 } },
      parseDOM: [{ style: 'font-size', getAttrs: (v) => ({ size: parseInt(v as string, 10) || 18 }) }],
      toDOM: (mark) => {
        // Coerce to a finite number so a crafted attr can't inject CSS.
        const size = Number(mark.attrs.size);
        // The editor container sets --spindle-fs to the body's autofit font scale (1 when none).
        return ['span', { style: `font-size:calc(${Number.isFinite(size) ? size : 18}px * var(--spindle-fs, 1))` }, 0];
      },
    },
    highlight: {
      attrs: { color: { default: null } },
      parseDOM: [{ style: 'background-color', getAttrs: (v) => ({ color: { kind: 'rgb', hex: v as string } }) }],
      toDOM: (mark) => {
        const css = colorToCss(mark.attrs.color as Color | null);
        if (!css) return ['span', {}, 0];
        return ['span', { style: `background-color:${css}` }, 0];
      },
    },
  },
});
