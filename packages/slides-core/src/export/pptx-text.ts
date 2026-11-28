// Text mappings between the PPTX interchange and the slides model: a `PptxTextBody` (paragraphs of effective level style and runs)
// and a `RichTextDoc` with the element's `BodyStyle`. The interchange states every paragraph's merged style; the model keeps
// paragraph attributes on the paragraph, run attributes as marks, and the size, colour, weight and family the text shares as the
// body's defaults, so a run carries a mark only where it differs from them.

import type { PptxBodyProps, PptxBullet, PptxColor, PptxLevelStyle, PptxParagraph, PptxRun, PptxTextBody } from '@b32nio/spindle-shared/export/pptx';
import type { BodyStyle, Color } from '../scene/types';
import type { Bullet, BulletScheme, FieldAttrs, ParagraphAttrs, RichTextDoc, RichTextInline, RichTextMark, RichTextParagraph, RichTextText, TabStop } from '../text/model';
import { colorFromPptx, colorToPptx, finite, ptToPx, pxToPt } from './pptx-values';

/** What an import changed or could not keep in text, for the importer's warnings. */
export interface TextTally {
  lineBreaks: number;
  verticalText: number;
}

export const newTextTally = (): TextTally => ({ lineBreaks: 0, verticalText: 0 });

// ── Body properties ──────────────────────────────────────────────────────────

/** `a:bodyPr` insets when the file states none: 0.1 in left and right, 0.05 in top and bottom. */
const DEFAULT_INSETS = { l: 9.6, t: 4.8, r: 9.6, b: 4.8 };
const VERTS = new Set(['horz', 'vert', 'eaVert']);

export function bodyStyleFromPptx(props: PptxBodyProps, tally: TextTally, anchor?: PptxBodyProps['anchor']): BodyStyle {
  const style: BodyStyle = {
    vAlign: props.anchor ?? anchor ?? 'top',
    insets: { ...DEFAULT_INSETS, ...definedOnly(props.insets) },
  };
  if (props.anchorCtr !== undefined) style.anchorCtr = props.anchorCtr;
  if (props.wrap === 'none') style.wrap = false;
  if (props.vert !== undefined && props.vert !== 'horz') {
    if (VERTS.has(props.vert)) style.vert = props.vert as BodyStyle['vert'];
    else tally.verticalText++;
  }
  if (props.autofit) {
    style.autofit = {
      mode: props.autofit.kind,
      ...(props.autofit.fontScale !== undefined ? { fontScale: props.autofit.fontScale } : {}),
      ...(props.autofit.lnSpcReduction !== undefined ? { lnSpcReduction: props.autofit.lnSpcReduction } : {}),
    };
  }
  if (props.columns !== undefined && props.columns > 1) {
    style.columns = { count: props.columns, ...(props.columnSpacing !== undefined ? { gap: props.columnSpacing } : {}) };
  }
  return style;
}

function definedOnly<T extends object>(o: T | undefined): Partial<T> {
  return Object.fromEntries(Object.entries(o ?? {}).filter(([, v]) => v !== undefined)) as Partial<T>;
}

export function bodyPropsToPptx(style: BodyStyle | undefined): PptxBodyProps {
  const props: PptxBodyProps = {};
  if (!style) return props;
  const pad = finite(style.padding);
  const insets = style.insets ?? (pad !== undefined ? { l: pad, t: pad, r: pad, b: pad } : undefined);
  if (insets) props.insets = { l: insets.l, t: insets.t, r: insets.r, b: insets.b };
  if (style.vAlign) props.anchor = style.vAlign;
  if (style.anchorCtr !== undefined) props.anchorCtr = style.anchorCtr;
  if (style.vert) props.vert = style.vert;
  if (style.wrap === false) props.wrap = 'none';
  if (style.autofit) {
    props.autofit = {
      kind: style.autofit.mode,
      ...(style.autofit.fontScale !== undefined ? { fontScale: style.autofit.fontScale } : {}),
      ...(style.autofit.lnSpcReduction !== undefined ? { lnSpcReduction: style.autofit.lnSpcReduction } : {}),
    };
  }
  if (style.columns && style.columns.count > 1) {
    props.columns = style.columns.count;
    if (style.columns.gap !== undefined) props.columnSpacing = style.columns.gap;
  }
  return props;
}

// ── Bullets ──────────────────────────────────────────────────────────────────

const SCHEME_OF: Record<string, BulletScheme> = {
  arabicPeriod: 'arabic', arabicParenR: 'arabic', arabicParenBoth: 'arabic', arabicPlain: 'arabic',
  alphaLcPeriod: 'alphaLower', alphaLcParenR: 'alphaLower', alphaLcParenBoth: 'alphaLower',
  alphaUcPeriod: 'alphaUpper', alphaUcParenR: 'alphaUpper', alphaUcParenBoth: 'alphaUpper',
  romanLcPeriod: 'romanLower', romanLcParenR: 'romanLower', romanLcParenBoth: 'romanLower',
  romanUcPeriod: 'romanUpper', romanUcParenR: 'romanUpper', romanUcParenBoth: 'romanUpper',
};
const AUTO_NUM_OF: Record<BulletScheme, string> = { arabic: 'arabicPeriod', alphaLower: 'alphaLcPeriod', alphaUpper: 'alphaUcPeriod', romanLower: 'romanLcPeriod', romanUpper: 'romanUcPeriod' };

export function bulletFromPptx(b: PptxBullet, style: PptxLevelStyle): Bullet {
  const out: Bullet =
    b.kind === 'none' ? { kind: 'none' }
    : b.kind === 'char' ? { kind: 'char', char: b.char }
    : { kind: 'auto', scheme: SCHEME_OF[b.scheme] ?? 'arabic', ...(b.startAt !== undefined ? { startAt: b.startAt } : {}) };
  if (b.kind !== 'none') {
    if (style.bulletColor) out.color = colorFromPptx(style.bulletColor);
    if (style.bulletSizePct !== undefined) out.sizePct = Math.round(style.bulletSizePct * 100);
  }
  return out;
}

// ── Paragraph attributes ─────────────────────────────────────────────────────

/** The model's paragraph attributes for a level style; only what the style states. */
export function paragraphAttrsFromPptx(level: number, style: PptxLevelStyle): ParagraphAttrs {
  const a: ParagraphAttrs = {};
  if (style.align) a.align = style.align;
  if (level > 0) a.indent = level;
  if (style.lineSpacingPts !== undefined) a.lineHeightPx = ptToPx(style.lineSpacingPts);
  else if (style.lineHeight !== undefined) a.lineHeight = style.lineHeight;
  if (style.spaceBeforePct !== undefined) a.spaceBeforePct = Math.round(style.spaceBeforePct * 100);
  else if (style.spaceBefore !== undefined) a.spaceBefore = ptToPx(style.spaceBefore);
  if (style.spaceAfterPct !== undefined) a.spaceAfterPct = Math.round(style.spaceAfterPct * 100);
  else if (style.spaceAfter !== undefined) a.spaceAfter = ptToPx(style.spaceAfter);
  if (style.bullet) a.bullet = bulletFromPptx(style.bullet, style);
  if (style.marginLeft !== undefined) a.marginLeft = style.marginLeft;
  if (style.indent !== undefined) a.firstLineIndent = style.indent;
  const tabs = (style.tabStops ?? []).filter((t) => t.align === 'l' || t.align === 'ctr' || t.align === 'r' || t.align === 'dec');
  if (tabs.length > 0) a.tabStops = tabs.map((t): TabStop => ({ pos: t.pos, align: t.align as TabStop['align'] }));
  return a;
}

const LEGACY_STEP = 28;

/** The level style a paragraph's attributes say; `defaults` are the body's shared size, weight, colour and family. */
export function levelStyleFromAttrs(a: ParagraphAttrs | undefined, defaults: PptxLevelStyle): PptxLevelStyle {
  const s: PptxLevelStyle = { ...defaults };
  if (!a) return s;
  if (a.align) s.align = a.align;
  const lineHeightPx = finite(a.lineHeightPx);
  const lineHeight = finite(a.lineHeight);
  if (lineHeightPx !== undefined && lineHeightPx > 0) s.lineSpacingPts = pxToPt(lineHeightPx);
  else if (lineHeight !== undefined) s.lineHeight = lineHeight;
  const before = finite(a.spaceBeforePct);
  const after = finite(a.spaceAfterPct);
  if (before !== undefined && before >= 0) s.spaceBeforePct = before / 100;
  else if (finite(a.spaceBefore)) s.spaceBefore = pxToPt(a.spaceBefore as number);
  if (after !== undefined && after >= 0) s.spaceAfterPct = after / 100;
  else if (finite(a.spaceAfter)) s.spaceAfter = pxToPt(a.spaceAfter as number);
  const bullet = a.bullet ?? undefined;
  if (bullet && typeof bullet === 'object') {
    s.bullet = bullet.kind === 'none' ? { kind: 'none' } : bullet.kind === 'char' ? { kind: 'char', char: bullet.char ?? '•' } : { kind: 'auto', scheme: AUTO_NUM_OF[bullet.scheme ?? 'arabic'], ...(finite(bullet.startAt) !== undefined ? { startAt: bullet.startAt } : {}) };
    if (bullet.kind !== 'none') {
      if (bullet.color) s.bulletColor = colorToPptx(bullet.color);
      const size = finite(bullet.sizePct);
      if (size !== undefined) s.bulletSizePct = size / 100;
    }
  } else if (a.listType === 'bullet' || a.listType === 'number') {
    s.bullet = a.listType === 'bullet' ? { kind: 'char', char: '•' } : { kind: 'auto', scheme: 'arabicPeriod' };
  }
  const margin = finite(a.marginLeft);
  const first = finite(a.firstLineIndent);
  const step = Math.min(8, Math.max(0, Math.trunc(finite(a.indent) ?? 0))) * LEGACY_STEP;
  if (margin !== undefined) {
    s.marginLeft = margin;
    if (first !== undefined) s.indent = first;
  } else if (s.bullet && s.bullet.kind !== 'none') {
    // the legacy layout (no margin stated) puts the marker one step left of the text
    s.marginLeft = step + LEGACY_STEP;
    s.indent = first ?? -LEGACY_STEP;
  } else if (step > 0) s.marginLeft = step;
  else if (first !== undefined) s.indent = first;
  const tabs = (a.tabStops ?? []).filter((t) => finite(t?.pos) !== undefined);
  if (tabs.length > 0) s.tabStops = tabs.map((t) => ({ pos: t.pos, align: t.align }));
  return s;
}

// ── Runs ─────────────────────────────────────────────────────────────────────

const colorKey = (c: Color | PptxColor | undefined): string | undefined => JSON.stringify(c);

/** Marks for one run: what it states over the paragraph's style, and where that differs from the body's defaults. */
function marksFromPptx(run: PptxRun, style: PptxLevelStyle, defaults: BodyStyle): RichTextMark[] {
  const marks: RichTextMark[] = [];
  const bold = run.bold ?? style.bold;
  if (bold && defaults.bold !== true) marks.push({ type: 'bold' });
  if (run.italic ?? style.italic) marks.push({ type: 'italic' });
  if (run.underline) marks.push({ type: 'underline' });
  if (run.strike) marks.push({ type: 'strikethrough' });
  if (run.link?.kind === 'url') marks.push({ type: 'link', attrs: { href: run.link.href } });
  const color = run.color ?? style.color;
  if (color && colorKey(color) !== colorKey(defaults.color)) marks.push({ type: 'textColor', attrs: { color: colorFromPptx(color) } });
  const font = run.font ?? style.font;
  if (font !== undefined && font !== defaults.fontFamily) marks.push({ type: 'fontFamily', attrs: { family: font } });
  const size = run.size ?? style.fontSize;
  if (size !== undefined && ptToPx(size) !== defaults.fontSize) marks.push({ type: 'fontSize', attrs: { size: ptToPx(size) } });
  if (run.highlight) marks.push({ type: 'highlight', attrs: { color: colorFromPptx(run.highlight) } });
  if (run.baseline !== undefined && run.baseline !== 0) marks.push({ type: 'baseline', attrs: { pct: Math.round(run.baseline * 100) } });
  if (run.caps) marks.push({ type: 'caps', attrs: { kind: run.caps } });
  if (run.spacing !== undefined && run.spacing !== 0) marks.push({ type: 'charSpacing', attrs: { px: ptToPx(run.spacing) } });
  if (run.lang) marks.push({ type: 'lang', attrs: { lang: run.lang } });
  return marks;
}

const textNode = (text: string, marks: RichTextMark[]): RichTextText => ({ type: 'text', text, ...(marks.length > 0 ? { marks } : {}) });

/** A field the editor models (slide number, date or time) from `a:fld/@type`, or `undefined` for any other field (kept as its cached text). */
function fieldFromPptx(type: string | undefined): FieldAttrs | undefined {
  if (type === 'slidenum') return { kind: 'slidenum' };
  const m = /^datetime(\d*)$/.exec(type ?? '');
  if (!m) return undefined;
  const n = Number(m[1]);
  return { kind: 'datetime', format: n >= 10 && n <= 13 ? 'time' : 'date' };
}

function inlinesFromPptx(runs: PptxRun[], style: PptxLevelStyle, defaults: BodyStyle): RichTextInline[][] {
  const lines: RichTextInline[][] = [[]];
  for (const run of runs) {
    if (run.kind === 'break') {
      lines.push([]);
      continue;
    }
    if (run.text === '') continue;
    const marks = marksFromPptx(run, style, defaults);
    const field = run.kind === 'field' ? fieldFromPptx(run.fieldType) : undefined;
    lines[lines.length - 1].push(field ? { type: 'field', attrs: field, content: [textNode(run.text, marks)] } : textNode(run.text, marks));
  }
  return lines;
}

/** The size, colour and family of a body's first paragraph, and bold only when every run is bold (a run cannot un-bold a default). */
function bodyDefaults(paragraphs: PptxParagraph[]): Pick<BodyStyle, 'fontSize' | 'color' | 'fontFamily' | 'bold'> {
  const first = paragraphs[0]?.style;
  const out: Pick<BodyStyle, 'fontSize' | 'color' | 'fontFamily' | 'bold'> = {};
  if (first?.fontSize !== undefined) out.fontSize = ptToPx(first.fontSize);
  if (first?.color) out.color = colorFromPptx(first.color);
  if (first?.font !== undefined) out.fontFamily = first.font;
  const runs = paragraphs.flatMap((p) => p.runs.filter((r) => r.kind !== 'break' && r.text !== '').map((r) => r.bold ?? p.style.bold));
  if (runs.length > 0 && runs.every(Boolean)) out.bold = true;
  return out;
}

/** A text body as rich text and body style. A line break (`a:br`) starts a new paragraph that continues the first's indent without its bullet. */
export function textFromPptx(body: PptxTextBody, tally: TextTally, anchor?: PptxBodyProps['anchor']): { richText: RichTextDoc; bodyStyle: BodyStyle } {
  const bodyStyle: BodyStyle = { ...bodyStyleFromPptx(body.props, tally, anchor), ...bodyDefaults(body.paragraphs) };
  const content: RichTextParagraph[] = [];
  for (const p of body.paragraphs) {
    const attrs = paragraphAttrsFromPptx(p.level, p.style);
    const lines = inlinesFromPptx(p.runs, p.style, bodyStyle);
    lines.forEach((inlines, i) => {
      if (i > 0) tally.lineBreaks++;
      const lineAttrs: ParagraphAttrs =
        i === 0
          ? attrs
          : { ...attrs, bullet: { kind: 'none' }, spaceBefore: 0, spaceBeforePct: null, ...(attrs.firstLineIndent !== undefined ? { firstLineIndent: 0 } : {}) };
      content.push({ type: 'paragraph', ...(Object.keys(lineAttrs).length > 0 ? { attrs: lineAttrs } : {}), ...(inlines.length > 0 ? { content: inlines } : {}) });
    });
  }
  if (content.length === 0) content.push({ type: 'paragraph' });
  return { richText: { type: 'doc', content }, bodyStyle };
}

// ── Model to interchange ─────────────────────────────────────────────────────

const str = (v: unknown): string | undefined => (typeof v === 'string' && v !== '' ? v : undefined);
const colorOf = (v: unknown): Color | undefined => {
  const c = v as Partial<Color> | null | undefined;
  return c && typeof c === 'object' && (c.kind === 'rgb' || c.kind === 'theme') ? (c as Color) : undefined;
};

const FIELD_TYPES = { slidenum: 'slidenum', date: 'datetime1', time: 'datetime11', datetime: 'datetime1' } as const;

function runFromMarks(text: string, marks: readonly RichTextMark[]): PptxRun {
  const run: PptxRun = { kind: 'text', text };
  for (const m of marks) {
    const a = m.attrs ?? {};
    switch (m.type) {
      case 'bold': run.bold = true; break;
      case 'italic': run.italic = true; break;
      case 'underline': run.underline = 'sng'; break;
      case 'strikethrough': run.strike = 'single'; break;
      case 'link': {
        const href = str(a.href);
        if (href !== undefined) run.link = { kind: 'url', href };
        break;
      }
      case 'textColor': {
        const c = colorOf(a.color);
        if (c) run.color = colorToPptx(c);
        break;
      }
      case 'fontFamily': {
        const f = str(a.family);
        if (f !== undefined) run.font = f;
        break;
      }
      case 'fontSize': {
        const n = finite(a.size);
        if (n !== undefined && n > 0) run.size = pxToPt(n);
        break;
      }
      case 'highlight': {
        const c = colorOf(a.color);
        if (c) run.highlight = colorToPptx(c);
        break;
      }
      case 'baseline': {
        const n = finite(a.pct);
        if (n !== undefined && n !== 0) run.baseline = n / 100;
        break;
      }
      case 'caps':
        if (a.kind === 'all' || a.kind === 'small') run.caps = a.kind;
        break;
      case 'charSpacing': {
        const n = finite(a.px);
        if (n !== undefined && n !== 0) run.spacing = pxToPt(n);
        break;
      }
      case 'lang': {
        const l = str(a.lang);
        if (l !== undefined) run.lang = l;
        break;
      }
    }
  }
  return run;
}

function runsFromInlines(inlines: RichTextInline[] | undefined): PptxRun[] {
  const runs: PptxRun[] = [];
  for (const inline of inlines ?? []) {
    if (inline.type === 'text') {
      if (inline.text !== '') runs.push(runFromMarks(inline.text, inline.marks ?? []));
      continue;
    }
    const text = (inline.content ?? []).map((t) => t.text).join('');
    const marks = [...(inline.content?.[0]?.marks ?? []), ...(inline.marks ?? [])];
    const run = runFromMarks(text, marks);
    if (inline.attrs.kind === 'footer') runs.push(run);
    else runs.push({ ...run, kind: 'field', fieldType: FIELD_TYPES[inline.attrs.kind === 'slidenum' ? 'slidenum' : (inline.attrs.format ?? 'date')] ?? 'datetime1' });
  }
  return runs;
}

/** The effective defaults a body states for its paragraphs: size (pt), weight, colour and family. */
function defaultsToPptx(style: BodyStyle | undefined): PptxLevelStyle {
  const s: PptxLevelStyle = {};
  const size = finite(style?.fontSize);
  if (size !== undefined && size > 0) s.fontSize = pxToPt(size);
  if (style?.bold !== undefined) s.bold = style.bold;
  const color = colorOf(style?.color);
  if (color) s.color = colorToPptx(color);
  const font = str(style?.fontFamily);
  if (font !== undefined) s.font = font;
  return s;
}

export function textToPptx(doc: RichTextDoc, style: BodyStyle | undefined): PptxTextBody {
  const defaults = defaultsToPptx(style);
  const paragraphs = doc.content.map((p): PptxParagraph => {
    const level = Math.min(8, Math.max(0, Math.trunc(finite(p.attrs?.indent) ?? 0)));
    const runs = runsFromInlines(p.content);
    return {
      level,
      style: levelStyleFromAttrs(p.attrs, defaults),
      runs,
      ...(runs.length === 0 && defaults.fontSize !== undefined ? { endSize: defaults.fontSize } : {}),
    };
  });
  return { props: bodyPropsToPptx(style), paragraphs };
}
