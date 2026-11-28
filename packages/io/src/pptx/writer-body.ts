import type { PptxLevelStyle, PptxLevelStyles, PptxParagraph, PptxRun, PptxTextBody } from '@b32nio/spindle-shared/export/pptx';
import { hlinkXml, type WriteContext } from './writer-context';
import { bodyPrXml, levelPropsXml, lstStyleXml, typefaceOf } from './writer-text';
import { attrs, esc, highlightXml, pct, solidFillXml } from './writer-xml';

// Text bodies of shapes (plan K12b, K3): `a:bodyPr`, a list style, paragraphs and runs. The inverse of `readTextBody`.
//
// The model keeps each paragraph's *effective* level style. PowerPoint reads a run's default size, weight, colour and
// font from the shape's list style and ignores them on a paragraph, so the style every paragraph of a level shares goes into
// `a:lstStyle`, and a paragraph's `a:pPr` states only how it differs. Read back, the two merge to the same effective style.

/** `ST_TextUnderlineType`, apart from `none` (the reader does not keep it). */
const UNDERLINES = new Set(['words', 'sng', 'dbl', 'heavy', 'dotted', 'dottedHeavy', 'dash', 'dashHeavy', 'dashLong', 'dashLongHeavy', 'dotDash', 'dotDashHeavy', 'dotDotDash', 'dotDotDashHeavy', 'wavy', 'wavyHeavy', 'wavyDbl']);
const CAPS = { all: 'all', small: 'small' } as const;
const STRIKES = { single: 'sngStrike', double: 'dblStrike' } as const;

/** The properties every style of the list states with the same value. */
function commonStyle(styles: PptxLevelStyle[]): PptxLevelStyle {
  const [first, ...rest] = styles;
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(first)) if (rest.every((s) => JSON.stringify((s as Record<string, unknown>)[k]) === JSON.stringify(v))) out[k] = v;
  return out as PptxLevelStyle;
}

/** The properties of `style` that `base` does not already state with the same value. */
function styleDifference(base: PptxLevelStyle, style: PptxLevelStyle): PptxLevelStyle {
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(style)) if (JSON.stringify((base as Record<string, unknown>)[k]) !== JSON.stringify(v)) out[k] = v;
  return out as PptxLevelStyle;
}

/** The level a paragraph is written at: `lvl` is 0 to 8. */
const levelOf = (p: PptxParagraph): number => Math.min(8, Math.max(0, Math.floor(p.level)));

/** What each level of a body has in common, as a list style: levels no paragraph uses are empty. */
export function sharedLevels(paragraphs: PptxParagraph[]): PptxLevelStyles {
  return Array.from({ length: 9 }, (_, level) => {
    const styles = paragraphs.filter((p) => levelOf(p) === level).map((p) => p.style);
    return styles.length === 0 ? {} : commonStyle(styles);
  });
}

/** `a:rPr`, `a:endParaRPr` or the properties of a field, with only what the run states. `hlink` is the run's `a:hlinkClick`. */
function runPropsXml(tag: string, run: Partial<PptxRun>, size: number | undefined, hlink: string, ctx: WriteContext): string {
  let underline = run.underline;
  if (underline !== undefined && !UNDERLINES.has(underline)) {
    ctx.lossy('underline');
    underline = 'sng';
  }
  const a = attrs({
    lang: run.lang,
    sz: size !== undefined ? Math.round(size * 100) : undefined,
    b: run.bold === undefined ? undefined : run.bold ? 1 : 0,
    i: run.italic === undefined ? undefined : run.italic ? 1 : 0,
    u: underline,
    strike: run.strike ? STRIKES[run.strike] : undefined,
    cap: run.caps ? CAPS[run.caps] : undefined,
    spc: run.spacing !== undefined ? Math.round(run.spacing * 100) : undefined,
    baseline: run.baseline !== undefined ? pct(run.baseline) : undefined,
  });
  const children =
    (run.color ? solidFillXml(run.color) : '') +
    (run.highlight ? highlightXml(run.highlight) : '') +
    (run.font !== undefined ? `<a:latin${attrs({ typeface: typefaceOf(run.font) })}/>` : '') +
    (run.fontEastAsian !== undefined ? `<a:ea${attrs({ typeface: typefaceOf(run.fontEastAsian, 'ea') })}/>` : '') +
    (run.fontComplex !== undefined ? `<a:cs${attrs({ typeface: typefaceOf(run.fontComplex, 'cs') })}/>` : '') +
    hlink;
  return children === '' ? (a === '' ? '' : `<${tag}${a}/>`) : `<${tag}${a}>${children}</${tag}>`;
}

function runXml(run: PptxRun, ctx: WriteContext): string {
  const props = (tag = 'a:rPr') => runPropsXml(tag, run, run.size, hlinkXml(run.link, ctx), ctx);
  if (run.kind === 'break') return `<a:br>${props()}</a:br>`;
  if (run.kind === 'field') return `<a:fld${attrs({ id: ctx.nextFieldId(), type: run.fieldType })}>${props()}<a:t>${esc(run.text)}</a:t></a:fld>`;
  return `<a:r>${props()}<a:t>${esc(run.text)}</a:t></a:r>`;
}

/** Known limit: a paragraph property that differs from its level's shared style is written in `a:pPr/a:defRPr`, which PowerPoint ignores, so such a difference may not render there (it reads back equal); check with the real-app run (D8). */
function paragraphXml(p: PptxParagraph, shared: PptxLevelStyles, ctx: WriteContext): string {
  const level = levelOf(p);
  const pPr = levelPropsXml('a:pPr', styleDifference(shared[level], p.style), { lvl: level > 0 ? level : undefined, rtl: p.rtl ? 1 : undefined });
  const end = p.endSize !== undefined ? `<a:endParaRPr${attrs({ sz: Math.round(p.endSize * 100) })}/>` : '';
  return `<a:p>${pPr}${p.runs.map((r) => runXml(r, ctx)).join('')}${end}</a:p>`;
}

/** The children of `p:txBody`: `a:bodyPr`, `a:lstStyle`, then at least one paragraph. */
export function txBodyXml(text: PptxTextBody, ctx: WriteContext): string {
  const shared = sharedLevels(text.paragraphs);
  const paragraphs = text.paragraphs.length === 0 ? '<a:p/>' : text.paragraphs.map((p) => paragraphXml(p, shared, ctx)).join('');
  return `${bodyPrXml(text.props)}${lstStyleXml(shared)}${paragraphs}`;
}
