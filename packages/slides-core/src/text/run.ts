// Pure run-level resolution shared by the editor schema (toDOM) and the static renderer: the K3c run marks (baseline,
// caps, character spacing, language) and the inline field runs (slide number, date/time, footer). Every value may come
// from imported or collaborator JSON, so numbers are coerced and clamped, enums are allowlisted and the language tag is
// matched against a strict pattern before it reaches the DOM.

import type { DateFieldFormat, FieldAttrs, FieldKind, RichTextField, RichTextMark } from './model';
import { safeCssNumber } from './sanitize';

/** A raised or lowered run is drawn at this fraction of its size (PowerPoint's superscript/subscript look). */
export const BASELINE_SCALE = 0.65;
export const MAX_BASELINE_PCT = 100;
export const MAX_SPACING_PX = 200;

export type CapsKind = 'all' | 'small';

export interface ResolvedRun {
  /** Percent of the font size the run is raised (negative: lowered); null when on the baseline. */
  baselinePct: number | null;
  caps: CapsKind | null;
  /** Extra letter spacing in px (negative: tighter); null when none. */
  spacingPx: number | null;
  /** A well-formed BCP 47 tag, or null. */
  lang: string | null;
}

// Language subtag, then up to three more alphanumeric subtags (script, region, variant): `en-US`, `zh-Hans-CN`.
const LANG_TAG = /^[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8}){0,3}$/;

/** The language tag if it is well formed, else null. */
export function safeLang(v: unknown): string | null {
  return typeof v === 'string' && LANG_TAG.test(v) ? v : null;
}

const clamp = (n: number, lo: number, hi: number): number => Math.min(hi, Math.max(lo, n));

/** Resolve the run marks of one inline into render values. */
export function resolveRun(marks: readonly RichTextMark[] | undefined): ResolvedRun {
  const out: ResolvedRun = { baselinePct: null, caps: null, spacingPx: null, lang: null };
  for (const m of marks ?? []) {
    const a = m.attrs ?? {};
    if (m.type === 'baseline') {
      const n = safeCssNumber(a.pct);
      if (n !== undefined && n !== 0) out.baselinePct = clamp(n, -MAX_BASELINE_PCT, MAX_BASELINE_PCT);
    } else if (m.type === 'caps') {
      out.caps = a.kind === 'all' || a.kind === 'small' ? a.kind : null;
    } else if (m.type === 'charSpacing') {
      const n = safeCssNumber(a.px);
      if (n !== undefined && n !== 0) out.spacingPx = clamp(n, -MAX_SPACING_PX, MAX_SPACING_PX);
    } else if (m.type === 'lang') {
      out.lang = safeLang(a.lang);
    }
  }
  return out;
}

/** CSS for the text-level run marks (caps and spacing), keyed in kebab-case for the schema and camelCase for React. */
export function runCss(r: ResolvedRun): { textTransform?: string; fontVariant?: string; letterSpacing?: string } {
  return {
    ...(r.caps === 'all' ? { textTransform: 'uppercase' } : {}),
    ...(r.caps === 'small' ? { fontVariant: 'small-caps' } : {}),
    ...(r.spacingPx !== null ? { letterSpacing: `${r.spacingPx}px` } : {}),
  };
}

/** CSS for the raised/lowered inner span: smaller, shifted by `top` (in em of the smaller size), no effect on the line box. */
export function baselineCss(pct: number): { position: 'relative'; top: string; fontSize: string; lineHeight: number } {
  return { position: 'relative', top: `${Math.round((-pct / 100 / BASELINE_SCALE) * 1000) / 1000}em`, fontSize: `${BASELINE_SCALE}em`, lineHeight: 0 };
}

const FIELD_KINDS: ReadonlySet<string> = new Set<FieldKind>(['slidenum', 'datetime', 'footer']);
const DATE_FORMATS: ReadonlySet<string> = new Set<DateFieldFormat>(['date', 'time', 'datetime']);
/** What an unresolved slide-number field shows (PowerPoint's `‹#›`). */
export const SLIDENUM_PLACEHOLDER = '‹#›';

export interface FieldContext {
  /** 1-based number of the slide being drawn, or null when it is not known (then the field shows its stored text). */
  slideNumber: number | null;
  /** The moment date/time fields show. */
  now: Date;
  /** BCP 47 locale for date/time formatting (default: the viewer's). */
  locale?: string;
}

export function safeFieldKind(v: unknown): FieldKind {
  return typeof v === 'string' && FIELD_KINDS.has(v) ? (v as FieldKind) : 'footer';
}

/**
 * The text a field shows in the given context; `stored` is its cached text, shown when nothing better can be computed.
 * A crafted kind or format degrades to the stored text.
 */
export function resolveFieldText(attrs: Partial<FieldAttrs> | undefined, ctx: FieldContext, stored = ''): string {
  const a = attrs ?? {};
  const kind = safeFieldKind(a.kind);
  if (kind === 'slidenum') {
    return ctx.slideNumber !== null && Number.isFinite(ctx.slideNumber) ? String(ctx.slideNumber) : stored || SLIDENUM_PLACEHOLDER;
  }
  if (kind === 'datetime') {
    const format = typeof a.format === 'string' && DATE_FORMATS.has(a.format) ? a.format : 'date';
    const opts: Intl.DateTimeFormatOptions =
      format === 'time' ? { timeStyle: 'short' } : format === 'datetime' ? { dateStyle: 'medium', timeStyle: 'short' } : { dateStyle: 'medium' };
    try {
      return new Intl.DateTimeFormat(ctx.locale, opts).format(ctx.now);
    } catch {
      return stored;
    }
  }
  return stored;
}

/** The run marks of a field: those on its cached text (where the Y layout keeps them), then any on the node itself. */
export function fieldMarks(field: RichTextField): RichTextMark[] {
  return [...(field.content?.[0]?.marks ?? []), ...(field.marks ?? [])];
}

/** The text a freshly inserted field starts with: what it shows now (a slide number reads ‹#› until drawn on a slide). */
export function initialFieldText(attrs: Partial<FieldAttrs>, now: Date = new Date()): string {
  return resolveFieldText(attrs, { slideNumber: null, now }, '');
}
