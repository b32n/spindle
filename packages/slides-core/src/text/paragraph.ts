// Pure paragraph layout resolution shared by the editor schema (toDOM) and the static renderer, so edit mode and idle
// mode agree. Every value here may come from imported or collaborator JSON, so numbers are coerced to finite numbers,
// enums are allowlisted and the bullet glyph is never interpolated into CSS (renderers expose it as a data attribute).

import type { Color } from '../scene/types';
import type { Bullet, BulletScheme, ParagraphAttrs, RichTextDoc, TabStop } from './model';
import { safeCssNumber } from './sanitize';

/** Pixels per indent level (the legacy `indent` attr). */
export const INDENT_PX = 28;
export const DEFAULT_BULLET_CHAR = '•';
const MAX_BULLET_CHAR = 4;

const SCHEMES: ReadonlySet<string> = new Set<BulletScheme>(['arabic', 'alphaLower', 'alphaUpper', 'romanLower', 'romanUpper']);

/** The CSS `counter-style` for each numbering scheme. */
export const SCHEME_COUNTER_STYLE: Record<BulletScheme, string> = {
  arabic: 'decimal',
  alphaLower: 'lower-alpha',
  alphaUpper: 'upper-alpha',
  romanLower: 'lower-roman',
  romanUpper: 'upper-roman',
};

export interface ResolvedMarker {
  kind: 'bullet' | 'number';
  /** Glyph for a bullet. */
  char: string;
  scheme: BulletScheme;
  startAt: number | null;
  color: Color | null;
  /** Percent of the text size. */
  sizePct: number;
}

export interface ResolvedParagraph {
  /** null when the paragraph has no marker. */
  marker: ResolvedMarker | null;
  /** Exact line height in px, or the multiplier. */
  lineHeight: { px: number } | { mult: number };
  spaceBefore: { px: number } | { em: number } | null;
  spaceAfter: { px: number } | { em: number } | null;
  /** Left padding of the paragraph block. */
  paddingLeft: number;
  /** text-indent in px (0 when none). */
  textIndent: number;
  /** Where a marker hangs, from the block's left edge. */
  markerLeft: number;
}

// Numeric strings are accepted (parseFloat), as the editor schema always did for line height and spacing.
const finite = safeCssNumber;

/** Whether a value is a bullet this module understands (a crafted kind is treated as absent). */
function validBullet(b: unknown): b is Bullet {
  return !!b && typeof b === 'object' && ((b as Bullet).kind === 'none' || (b as Bullet).kind === 'char' || (b as Bullet).kind === 'auto');
}

/** Normalise the bullet glyph: a short printable string, else the default bullet. */
export function safeBulletChar(c: unknown): string {
  if (typeof c !== 'string') return DEFAULT_BULLET_CHAR;
  const chars = Array.from(c).slice(0, MAX_BULLET_CHAR).join('').trim();
  // eslint-disable-next-line no-control-regex -- rejecting control characters is the point
  return chars === '' || /[\u0000-\u001f\u007f]/.test(chars) ? DEFAULT_BULLET_CHAR : chars;
}

function resolveMarker(a: ParagraphAttrs): ResolvedMarker | null {
  const b = a.bullet;
  if (validBullet(b)) {
    if (b.kind === 'none') return null;
    const size = finite(b.sizePct);
    const startAt = finite(b.startAt);
    return {
      kind: b.kind === 'char' ? 'bullet' : 'number',
      char: safeBulletChar(b.char),
      scheme: b.scheme && SCHEMES.has(b.scheme) ? b.scheme : 'arabic',
      startAt: startAt === undefined ? null : Math.trunc(startAt),
      color: b.color && typeof b.color === 'object' ? b.color : null,
      sizePct: size !== undefined && size > 0 ? Math.min(size, 400) : 100,
    };
  }
  if (a.listType === 'bullet' || a.listType === 'number') {
    return { kind: a.listType, char: DEFAULT_BULLET_CHAR, scheme: 'arabic', startAt: null, color: null, sizePct: 100 };
  }
  return null;
}

/** Resolve a paragraph's attrs to the layout values both renderers use. */
export function resolveParagraph(a: ParagraphAttrs): ResolvedParagraph {
  const marker = resolveMarker(a);
  const indent = finite(a.indent) ?? 0;
  const marginLeft = finite(a.marginLeft);
  const first = finite(a.firstLineIndent) ?? 0;
  const px = finite(a.lineHeightPx);
  const mult = finite(a.lineHeight) ?? 1.2;
  const space = (pct: unknown, abs: unknown): { px: number } | { em: number } | null => {
    const p = finite(pct);
    if (p !== undefined && p > 0) return { em: p / 100 };
    const v = finite(abs);
    return v ? { px: v } : null;
  };
  // Legacy layout: a 28px step per level, plus a one-step gutter for the marker.
  const legacyPad = marker ? (indent + 1) * INDENT_PX : indent * INDENT_PX;
  const paddingLeft = marginLeft === undefined ? legacyPad : Math.max(0, marginLeft);
  const hang = marginLeft === undefined ? undefined : Math.max(0, paddingLeft + first);
  return {
    marker,
    lineHeight: px !== undefined && px > 0 ? { px } : { mult },
    spaceBefore: space(a.spaceBeforePct, a.spaceBefore),
    spaceAfter: space(a.spaceAfterPct, a.spaceAfter),
    paddingLeft,
    // With a marker the first-line offset places the marker; without one it indents the text.
    textIndent: marker ? 0 : first,
    markerLeft: hang ?? (marker?.kind === 'bullet' ? 8 : 2),
  };
}

/** A sanitised copy of tab stops (finite positions, allowlisted alignment), sorted by position. */
export function normalizeTabStops(stops: unknown): TabStop[] {
  if (!Array.isArray(stops)) return [];
  const out: TabStop[] = [];
  for (const s of stops as Array<Partial<TabStop>>) {
    const pos = finite(s?.pos);
    if (pos === undefined || pos < 0) continue;
    const align = s.align === 'ctr' || s.align === 'r' || s.align === 'dec' ? s.align : 'l';
    out.push({ pos, align });
  }
  return out.sort((x, y) => x.pos - y.pos);
}

/** Per-level paragraph defaults (PPTX `a:lstStyle` / `lvl1pPr`..`lvl9pPr`), keyed by indent level 0 to 8. */
export type LevelStyles = Partial<Record<number, ParagraphAttrs>>;

/**
 * Bake level styles into the paragraphs of a body: each paragraph takes its level's defaults for every attr it does not
 * set itself. Importers call this so editor and static render need no inheritance logic of their own.
 */
export function flattenLevelStyles(doc: RichTextDoc, levels: LevelStyles | undefined): RichTextDoc {
  if (!levels) return doc;
  return {
    ...doc,
    content: doc.content.map((p) => {
      const level = Math.max(0, Math.min(8, Math.trunc(finite(p.attrs?.indent) ?? 0)));
      const base = levels[level];
      return base ? { ...p, attrs: { ...base, ...p.attrs } } : p;
    }),
  };
}

const ROMAN: Array<[number, string]> = [
  [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
];

/** The label for the nth item (1-based) of a numbered list in a scheme, without the trailing punctuation. */
export function formatOrdinal(n: number, scheme: BulletScheme): string {
  const v = Math.trunc(n);
  if (!Number.isFinite(v) || v < 1 || scheme === 'arabic') return String(Number.isFinite(v) ? v : 1);
  if (scheme === 'alphaLower' || scheme === 'alphaUpper') {
    let out = '';
    for (let x = v; x > 0; x = Math.floor((x - 1) / 26)) out = String.fromCharCode(97 + ((x - 1) % 26)) + out;
    return scheme === 'alphaUpper' ? out.toUpperCase() : out;
  }
  let out = '';
  let rest = Math.min(v, 3999);
  for (const [value, glyph] of ROMAN) {
    while (rest >= value) {
      out += glyph;
      rest -= value;
    }
  }
  return scheme === 'romanUpper' ? out.toUpperCase() : out;
}
