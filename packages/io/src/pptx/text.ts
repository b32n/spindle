import type { PptxBodyProps, PptxColorMap, PptxLevelStyle, PptxLevelStyles, PptxParagraph, PptxRun, PptxTextBody } from '@b32nio/spindle-shared/export/pptx';
import { isTrue, kid, type XmlElement } from '../xlsx/xml';
import { colorElement, parseColor } from './color';
import { readLink, type LinkContext } from './links';
import { mergeLevel, parseLevel, parseLevels } from './levels';
import { emuToPx, num, pct } from './units';

// Text bodies (plan K3): `a:bodyPr`, paragraphs with the effective level style, and runs with their own attributes.

const ANCHOR: Record<string, PptxBodyProps['anchor']> = { t: 'top', ctr: 'middle', b: 'bottom' };

/** `a:bodyPr`: only what it states. */
export function parseBodyProps(bodyPr: XmlElement | undefined): PptxBodyProps {
  const out: PptxBodyProps = {};
  if (!bodyPr) return out;
  const insets: NonNullable<PptxBodyProps['insets']> = {};
  for (const [key, attr] of [['l', 'lIns'], ['t', 'tIns'], ['r', 'rIns'], ['b', 'bIns']] as const) {
    const v = emuToPx(bodyPr.attrs[attr]);
    if (v !== undefined) insets[key] = v;
  }
  if (Object.keys(insets).length > 0) out.insets = insets;
  const anchor = ANCHOR[bodyPr.attrs.anchor ?? ''];
  if (anchor) out.anchor = anchor;
  if (bodyPr.attrs.anchorCtr !== undefined) out.anchorCtr = isTrue(bodyPr.attrs.anchorCtr);
  if (bodyPr.attrs.vert) out.vert = bodyPr.attrs.vert;
  if (bodyPr.attrs.wrap === 'none' || bodyPr.attrs.wrap === 'square') out.wrap = bodyPr.attrs.wrap;
  const columns = num(bodyPr.attrs.numCol);
  if (columns !== undefined && columns >= 1) out.columns = Math.floor(columns);
  const gap = emuToPx(bodyPr.attrs.spcCol);
  if (gap !== undefined && gap >= 0) out.columnSpacing = gap;
  const rot = num(bodyPr.attrs.rot);
  if (rot !== undefined && rot !== 0) out.rot = rot / 60000;
  if (kid(bodyPr, 'noAutofit')) out.autofit = { kind: 'none' };
  else if (kid(bodyPr, 'spAutoFit')) out.autofit = { kind: 'resize' };
  else {
    const norm = kid(bodyPr, 'normAutofit');
    if (norm) {
      const scale = pct(norm.attrs.fontScale);
      const reduction = pct(norm.attrs.lnSpcReduction);
      out.autofit = { kind: 'shrink', ...(scale !== undefined ? { fontScale: scale } : {}), ...(reduction !== undefined ? { lnSpcReduction: reduction } : {}) };
    }
  }
  return out;
}

/** Body properties over another set: insets merge side by side, every other property is replaced when stated. */
export function mergeBodyProps(base: PptxBodyProps | undefined, over: PptxBodyProps): PptxBodyProps {
  const insets = { ...base?.insets, ...over.insets };
  return { ...base, ...over, ...(Object.keys(insets).length > 0 ? { insets } : {}) };
}

/** A typeface attribute as the editor names it: `major` / `minor` for the theme fonts, else the family; `undefined` for nothing usable. */
function typeface(parent: XmlElement | undefined, script: string): string | undefined {
  const face = (kid(parent, script)?.attrs.typeface ?? '').trim();
  if (/^\+mj-/.test(face)) return 'major';
  if (/^\+mn-/.test(face)) return 'minor';
  return face === '' || face.startsWith('+') ? undefined : face;
}

/** `a:rPr` (or `a:endParaRPr`): the attributes the run states. */
export function parseRunProps(r: XmlElement | undefined, map: PptxColorMap): Omit<PptxRun, 'kind' | 'text'> {
  const out: Omit<PptxRun, 'kind' | 'text'> = {};
  if (!r) return out;
  if (r.attrs.b !== undefined) out.bold = isTrue(r.attrs.b);
  if (r.attrs.i !== undefined) out.italic = isTrue(r.attrs.i);
  if (r.attrs.u && r.attrs.u !== 'none') out.underline = r.attrs.u;
  if (r.attrs.strike === 'sngStrike') out.strike = 'single';
  else if (r.attrs.strike === 'dblStrike') out.strike = 'double';
  const size = num(r.attrs.sz);
  if (size !== undefined && size > 0) out.size = size / 100;
  const color = parseColor(colorElement(kid(r, 'solidFill')), map);
  if (color) out.color = color;
  const highlight = parseColor(colorElement(kid(r, 'highlight')), map);
  if (highlight) out.highlight = highlight;
  const font = typeface(r, 'latin');
  if (font) out.font = font;
  const ea = typeface(r, 'ea');
  if (ea) out.fontEastAsian = ea;
  const cs = typeface(r, 'cs');
  if (cs) out.fontComplex = cs;
  const baseline = pct(r.attrs.baseline);
  if (baseline !== undefined && baseline !== 0) out.baseline = baseline;
  if (r.attrs.cap === 'all' || r.attrs.cap === 'small') out.caps = r.attrs.cap;
  const spacing = num(r.attrs.spc);
  if (spacing !== undefined && spacing !== 0) out.spacing = spacing / 100;
  if (r.attrs.lang) out.lang = r.attrs.lang;
  return out;
}

function readParagraph(p: XmlElement, base: PptxLevelStyles, shapeList: PptxLevelStyles, map: PptxColorMap, links: LinkContext): PptxParagraph {
  const pPr = kid(p, 'pPr');
  const lvl = num(pPr?.attrs.lvl);
  const level = lvl !== undefined && lvl >= 0 && lvl <= 8 ? Math.floor(lvl) : 0;
  const style = mergeLevel(mergeLevel(base[level], shapeList[level]), parseLevel(pPr, map));
  const runs: PptxRun[] = [];
  const linkOf = (rPr: XmlElement | undefined): Pick<PptxRun, 'link'> => {
    const link = readLink(kid(rPr, 'hlinkClick'), links);
    return link ? { link } : {};
  };
  for (const c of p.children) {
    if (c.name === 'r') runs.push({ kind: 'text', text: kid(c, 't')?.text ?? '', ...parseRunProps(kid(c, 'rPr'), map), ...linkOf(kid(c, 'rPr')) });
    else if (c.name === 'fld') runs.push({ kind: 'field', text: kid(c, 't')?.text ?? '', ...(c.attrs.type ? { fieldType: c.attrs.type } : {}), ...parseRunProps(kid(c, 'rPr'), map), ...linkOf(kid(c, 'rPr')) });
    else if (c.name === 'br') runs.push({ kind: 'break', text: '', ...parseRunProps(kid(c, 'rPr'), map) });
  }
  const endSize = parseRunProps(kid(p, 'endParaRPr'), map).size;
  return { level, style, ...(isTrue(pPr?.attrs.rtl) ? { rtl: true } : {}), runs, ...(endSize !== undefined ? { endSize } : {}) };
}

/**
 * `p:txBody`. `base` is what a paragraph of each level starts from (the layout placeholder's levels, or the presentation's
 * default text style), `own` the body properties the placeholder chain gives. `fontRefColor` is the colour of the shape style's `fontRef`, which sits between them and the shape's own list style. `links` (the part's relationships) is what a run's hyperlink is read against.
 */
export function readTextBody(body: XmlElement, base: PptxLevelStyles, own: PptxBodyProps | undefined, map: PptxColorMap, fontRefColor: PptxLevelStyle['color'] | undefined, links: LinkContext): PptxTextBody {
  const withFont = fontRefColor ? base.map((l) => mergeLevel(l, { color: fontRefColor })) : base;
  const shapeList = parseLevels(kid(body, 'lstStyle'), map);
  return {
    props: mergeBodyProps(own, parseBodyProps(kid(body, 'bodyPr'))),
    paragraphs: body.children.filter((c) => c.name === 'p').map((p) => readParagraph(p, withFont, shapeList, map, links)),
  };
}
