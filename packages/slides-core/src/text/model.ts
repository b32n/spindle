// Rich-text model — the ProseMirror-JSON shape stored on text-bearing
// elements, plus pure helpers shared by the static renderer, the toolbar's
// idle-formatting path, and tests. Deliberately flat (doc > paragraph+ >
// inline*): PPTX text bodies are flat paragraphs with bullet attrs, not nested
// lists, which keeps static rendering and future PPTX export honest.
//
// This module is pure JSON — it does not import ProseMirror. The live editor
// (Phase 3) builds a PM schema whose serialized form matches these types.

import type { Color } from '../scene/types';

export type TextAlign = 'left' | 'center' | 'right' | 'justify';
export type ListType = 'none' | 'bullet' | 'number';

/** How a paragraph's bullet is drawn. `none` suppresses a marker, `char` is a literal glyph, `auto` is a numbered scheme. */
export type BulletKind = 'none' | 'char' | 'auto';
/** Auto-number schemes (PPTX `buAutoNum` types that map to a CSS counter style). */
export type BulletScheme = 'arabic' | 'alphaLower' | 'alphaUpper' | 'romanLower' | 'romanUpper';

/**
 * A paragraph bullet (PPTX `a:buChar` / `a:buAutoNum` / `a:buNone` with colour and size). When present it takes
 * precedence over the older `listType`.
 */
export interface Bullet {
  kind: BulletKind;
  /** The glyph, for `char` (default •). */
  char?: string;
  /** First number, for `auto` (default 1). */
  startAt?: number;
  /** Number style, for `auto` (default arabic). */
  scheme?: BulletScheme;
  /** Bullet colour (default: the text colour). */
  color?: Color;
  /** Bullet size as a percent of the text size (default 100). */
  sizePct?: number;
}

export interface TabStop {
  /** Position from the text box's left edge, in px. */
  pos: number;
  align: 'l' | 'ctr' | 'r' | 'dec';
}

export interface ParagraphAttrs {
  align?: TextAlign;
  listType?: ListType;
  /** Indent level, 0–8. */
  indent?: number;
  /** Multiplier on the font's line height, e.g. 1.2. */
  lineHeight?: number;
  /** Space before the paragraph, in px. */
  spaceBefore?: number;
  /** Space after the paragraph, in px. */
  spaceAfter?: number;
  /** Bullet (kind, glyph, numbering, colour, size). Overrides `listType` when set. */
  bullet?: Bullet | null;
  /** Exact line height in px; wins over the `lineHeight` multiplier (PPTX `spcPts` line spacing). */
  lineHeightPx?: number | null;
  /** Space before as a percent of the text size; wins over `spaceBefore` (PPTX `spcPct`). */
  spaceBeforePct?: number | null;
  /** Space after as a percent of the text size; wins over `spaceAfter`. */
  spaceAfterPct?: number | null;
  /** Left margin of the paragraph in px (PPTX `marL`); wins over the indent level's 28px steps. */
  marginLeft?: number | null;
  /** First-line indent in px relative to `marginLeft`; negative is a hanging indent (PPTX `indent`). */
  firstLineIndent?: number | null;
  /** Custom tab stops (stored and synced; the editor lays out default tabs). */
  tabStops?: TabStop[] | null;
}

/**
 * A text mark. `type` is one of the schema mark names (bold, italic,
 * underline, strikethrough, link, textColor, fontFamily, fontSize); `attrs`
 * carries the mark's parameters (e.g. a Color union for textColor). Kept
 * loosely typed here so the JSON model does not depend on the scene layer.
 */
export interface RichTextMark {
  type: string;
  attrs?: Record<string, unknown>;
}

export interface RichTextInline {
  type: 'text';
  text: string;
  marks?: RichTextMark[];
}

export interface RichTextParagraph {
  type: 'paragraph';
  attrs?: ParagraphAttrs;
  content?: RichTextInline[];
}

export interface RichTextDoc {
  type: 'doc';
  content: RichTextParagraph[];
}

/** An empty text body — a single empty paragraph (a valid PM doc). */
export function emptyRichText(): RichTextDoc {
  return { type: 'doc', content: [{ type: 'paragraph' }] };
}

/** Build a text body from plain text; each line becomes a paragraph. */
export function richTextFromPlainText(text: string, attrs?: ParagraphAttrs): RichTextDoc {
  const lines = text.split('\n');
  return {
    type: 'doc',
    content: lines.map((line) => ({
      type: 'paragraph',
      ...(attrs ? { attrs } : {}),
      ...(line.length > 0 ? { content: [{ type: 'text', text: line }] } : {}),
    })),
  };
}

/** Flatten a text body to plain text; paragraphs join with newlines. */
export function richTextToPlainText(doc: RichTextDoc): string {
  return doc.content
    .map((p) => (p.content ?? []).map((inline) => inline.text).join(''))
    .join('\n');
}

/** True when the body has no visible characters (used to show placeholder prompts). */
export function isRichTextEmpty(doc: RichTextDoc): boolean {
  return richTextToPlainText(doc).trim().length === 0;
}

// ── Pure formatting helpers (idle-path + tests) ──────────────────────────────
//
// Idle formatting has no text-range selection (the element is selected, not a
// caret range), so these apply across the whole body. The live editor uses
// ProseMirror transactions for range formatting.

function mapInlines(doc: RichTextDoc, fn: (inline: RichTextInline) => RichTextInline): RichTextDoc {
  return {
    type: 'doc',
    content: doc.content.map((p) => ({
      ...p,
      ...(p.content ? { content: p.content.map(fn) } : {}),
    })),
  };
}

/** Whether every text run in the body carries a mark of `type`. */
export function docHasMark(doc: RichTextDoc, type: string): boolean {
  let sawText = false;
  for (const p of doc.content) {
    for (const inline of p.content ?? []) {
      sawText = true;
      if (!(inline.marks ?? []).some((m) => m.type === type)) return false;
    }
  }
  return sawText;
}

/** Add (or replace) a mark across the whole body. */
export function applyMarkToDoc(doc: RichTextDoc, mark: RichTextMark): RichTextDoc {
  return mapInlines(doc, (inline) => ({
    ...inline,
    marks: [...(inline.marks ?? []).filter((m) => m.type !== mark.type), mark],
  }));
}

/** Remove a mark type across the whole body. */
export function removeMarkFromDoc(doc: RichTextDoc, type: string): RichTextDoc {
  return mapInlines(doc, (inline) => {
    const marks = (inline.marks ?? []).filter((m) => m.type !== type);
    return marks.length ? { ...inline, marks } : { type: 'text', text: inline.text };
  });
}

/** Toggle a boolean mark across the whole body. */
export function toggleMarkInDoc(doc: RichTextDoc, type: string, attrs?: Record<string, unknown>): RichTextDoc {
  return docHasMark(doc, type) ? removeMarkFromDoc(doc, type) : applyMarkToDoc(doc, { type, attrs });
}

/** Merge paragraph attrs into every paragraph. */
export function setParagraphAttrs(doc: RichTextDoc, attrs: Partial<ParagraphAttrs>): RichTextDoc {
  return {
    type: 'doc',
    content: doc.content.map((p) => ({ ...p, attrs: { ...p.attrs, ...attrs } })),
  };
}

export interface TextFormatSpec {
  /** Toggle a boolean mark (bold/italic/underline/strikethrough). */
  toggleMark?: string;
  /** Set a parameterized mark (textColor/fontFamily/fontSize/link). */
  setMark?: RichTextMark;
  /** Remove a mark type. */
  removeMark?: string;
  /** Merge paragraph attrs (align/listType/indent/…). */
  paragraph?: Partial<ParagraphAttrs>;
}

/** Apply an idle-path formatting spec to a body. */
export function applyTextFormat(doc: RichTextDoc, spec: TextFormatSpec): RichTextDoc {
  let out = doc;
  if (spec.toggleMark) out = toggleMarkInDoc(out, spec.toggleMark);
  if (spec.setMark) out = applyMarkToDoc(out, spec.setMark);
  if (spec.removeMark) out = removeMarkFromDoc(out, spec.removeMark);
  if (spec.paragraph) out = setParagraphAttrs(out, spec.paragraph);
  return out;
}
