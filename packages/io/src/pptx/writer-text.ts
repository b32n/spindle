import type { PptxBodyProps, PptxBullet, PptxLevelStyle, PptxLevelStyles } from '@b32nio/spindle-shared/export/pptx';
import { angle, attrs, colorXml, emu, esc, pct, solidFillXml } from './writer-xml';

// Text properties as DrawingML (plan K12, K3): a level style as `a:lvlNpPr`, body properties as `a:bodyPr`. The inverse of
// `parseLevel` and `parseBodyProps`: what the reader states comes back the same, and what it leaves out stays out.

const ALGN: Record<NonNullable<PptxLevelStyle['align']>, string> = { left: 'l', center: 'ctr', right: 'r', justify: 'just' };
const ANCHOR: Record<NonNullable<PptxBodyProps['anchor']>, string> = { top: 't', middle: 'ctr', bottom: 'b' };

/** `typeface` of a font as the editor names it: `major` and `minor` are the theme fonts. */
export const typefaceOf = (font: string): string => (font === 'major' ? '+mj-lt' : font === 'minor' ? '+mn-lt' : font);

function bulletXml(b: PptxBullet): string {
  if (b.kind === 'none') return '<a:buNone/>';
  if (b.kind === 'char') return `<a:buChar${attrs({ char: b.char })}/>`;
  return `<a:buAutoNum${attrs({ type: b.scheme, startAt: b.startAt })}/>`;
}

const spacingXml = (tag: string, pts: number | undefined, fraction: number | undefined): string =>
  pts !== undefined ? `<a:${tag}><a:spcPts val="${Math.round(pts * 100)}"/></a:${tag}>` : fraction !== undefined ? `<a:${tag}><a:spcPct val="${pct(fraction)}"/></a:${tag}>` : '';

/** The `a:defRPr` of a level: size, weight, colour and family. Empty string when the level states none of them. */
function defaultRunXml(s: PptxLevelStyle): string {
  const a = attrs({ sz: s.fontSize !== undefined ? Math.round(s.fontSize * 100) : undefined, b: s.bold === undefined ? undefined : s.bold ? 1 : 0, i: s.italic === undefined ? undefined : s.italic ? 1 : 0 });
  const children = (s.color ? solidFillXml(s.color) : '') + (s.font !== undefined ? `<a:latin${attrs({ typeface: typefaceOf(s.font) })}/>` : '');
  return a === '' && children === '' ? '' : `<a:defRPr${a}>${children}</a:defRPr>`;
}

/**
 * One level as `a:lvl<n>pPr` (`index` is 0 to 8), its children in the order of `CT_TextParagraphProperties`: line and paragraph
 * spacing, bullet colour, size, font and kind, tab stops, then the default run properties. An empty level is an empty string.
 */
export function levelXml(index: number, s: PptxLevelStyle): string {
  const a = attrs({
    marL: s.marginLeft !== undefined ? emu(s.marginLeft) : undefined,
    indent: s.indent !== undefined ? emu(s.indent) : undefined,
    algn: s.align !== undefined ? ALGN[s.align] : undefined,
  });
  const body =
    (s.lineHeight !== undefined ? `<a:lnSpc><a:spcPct val="${pct(s.lineHeight)}"/></a:lnSpc>` : s.lineSpacingPts !== undefined ? `<a:lnSpc><a:spcPts val="${Math.round(s.lineSpacingPts * 100)}"/></a:lnSpc>` : '') +
    spacingXml('spcBef', s.spaceBefore, s.spaceBeforePct) +
    spacingXml('spcAft', s.spaceAfter, s.spaceAfterPct) +
    (s.bulletColor ? `<a:buClr>${colorXml(s.bulletColor)}</a:buClr>` : '') +
    (s.bulletSizePct !== undefined ? `<a:buSzPct val="${pct(s.bulletSizePct)}"/>` : '') +
    (s.bulletFont !== undefined ? `<a:buFont${attrs({ typeface: s.bulletFont })}/>` : '') +
    (s.bullet ? bulletXml(s.bullet) : '') +
    (s.tabStops?.length ? `<a:tabLst>${s.tabStops.map((t) => `<a:tab${attrs({ pos: emu(t.pos), algn: t.align })}/>`).join('')}</a:tabLst>` : '') +
    defaultRunXml(s);
  if (a === '' && body === '') return '';
  const tag = `a:lvl${index + 1}pPr`;
  return body === '' ? `<${tag}${a}/>` : `<${tag}${a}>${body}</${tag}>`;
}

/** `a:lstStyle`: the levels that state anything (an empty list is `<a:lstStyle/>`). */
export function lstStyleXml(levels: PptxLevelStyles): string {
  const body = levels.slice(0, 9).map((l, i) => levelXml(i, l)).join('');
  return body === '' ? '<a:lstStyle/>' : `<a:lstStyle>${body}</a:lstStyle>`;
}

/** `a:bodyPr` with the properties the model carries: insets, anchor, vertical text, wrap, columns, rotation and the autofit mode. */
export function bodyPrXml(props: PptxBodyProps | undefined, anchor?: PptxBodyProps['anchor']): string {
  const p = props ?? {};
  const a = attrs({
    rot: p.rot !== undefined ? angle(p.rot) : undefined,
    vert: p.vert,
    wrap: p.wrap,
    lIns: p.insets?.l !== undefined ? emu(p.insets.l) : undefined,
    tIns: p.insets?.t !== undefined ? emu(p.insets.t) : undefined,
    rIns: p.insets?.r !== undefined ? emu(p.insets.r) : undefined,
    bIns: p.insets?.b !== undefined ? emu(p.insets.b) : undefined,
    numCol: p.columns,
    spcCol: p.columnSpacing !== undefined ? emu(p.columnSpacing) : undefined,
    anchor: p.anchor ?? anchor ? ANCHOR[(p.anchor ?? anchor) as keyof typeof ANCHOR] : undefined,
    anchorCtr: p.anchorCtr === undefined ? undefined : p.anchorCtr ? 1 : 0,
  });
  const fit = p.autofit;
  const child = !fit
    ? ''
    : fit.kind === 'none'
      ? '<a:noAutofit/>'
      : fit.kind === 'resize'
        ? '<a:spAutoFit/>'
        : `<a:normAutofit${attrs({ fontScale: fit.fontScale !== undefined ? pct(fit.fontScale) : undefined, lnSpcReduction: fit.lnSpcReduction !== undefined ? pct(fit.lnSpcReduction) : undefined })}/>`;
  return child === '' ? `<a:bodyPr${a}/>` : `<a:bodyPr${a}>${child}</a:bodyPr>`;
}

/** Plain paragraphs (one per line of `text`) for a prompt or a label. */
export function plainParagraphsXml(text: string): string {
  const lines = text === '' ? [''] : text.split('\n');
  return lines.map((l) => (l === '' ? '<a:p><a:endParaRPr lang="en-US"/></a:p>' : `<a:p><a:r><a:rPr lang="en-US"/><a:t>${esc(l)}</a:t></a:r></a:p>`)).join('');
}
