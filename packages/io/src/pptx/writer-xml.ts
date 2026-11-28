import type { PptxColor, PptxFrame } from '@b32nio/spindle-shared/export/pptx';
import { EMU_PER_PX } from './units';

// String builders shared by the PPTX writer's parts (plan K12). Every value that reaches the markup goes through `esc` or a
// number formatter, so a deck's text can never break out of an attribute or an element.

export const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
/** The three namespaces every presentation part declares on its root. */
export const ROOT_NS = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;
export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

// eslint-disable-next-line no-control-regex -- the characters XML 1.0 cannot carry
const INVALID_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;
const ESCAPES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };

/** Text for an element or an attribute: XML-invalid control characters are dropped, the markup characters escaped. */
export const esc = (s: string): string => s.replace(INVALID_XML, '').replace(/[&<>"]/g, (c) => ESCAPES[c]);

/** Attributes as ` name="value"` pairs; `undefined` values are left out and numbers are written as integers or plain decimals. */
export function attrs(values: Record<string, string | number | undefined>): string {
  let out = '';
  for (const [k, v] of Object.entries(values)) if (v !== undefined) out += ` ${k}="${typeof v === 'number' ? String(v) : esc(v)}"`;
  return out;
}

/** px to EMU, rounded to a whole number (the reader divides by the same 9525, so a value read from a file comes back exact). */
export const emu = (px: number): number => Math.round(px * EMU_PER_PX);
/** A fraction to `ST_Percentage` thousandths of a percent (`0.75` is `75000`). */
export const pct = (fraction: number): number => Math.round(fraction * 100000);
/** Degrees to the 60000ths DrawingML uses for angles. */
export const angle = (degrees: number): number => Math.round(degrees * 60000);

/** The `hex` of a `#RRGGBB` string in the upper case DrawingML uses; anything that is not a 6-digit hex is black. */
export const hex6 = (hex: string): string => (/^#?[0-9a-fA-F]{6}$/.test(hex) ? hex.replace('#', '').toUpperCase() : '000000');

/** A colour element: `a:schemeClr` for a theme slot, `a:srgbClr` for a literal, with `lumMod`, `lumOff` and `alpha` as children. */
export function colorXml(color: PptxColor): string {
  const transforms =
    (color.kind === 'theme' && color.lumMod !== undefined ? `<a:lumMod val="${pct(color.lumMod)}"/>` : '') +
    (color.kind === 'theme' && color.lumOff !== undefined ? `<a:lumOff val="${pct(color.lumOff)}"/>` : '') +
    (color.alpha !== undefined ? `<a:alpha val="${pct(color.alpha)}"/>` : '');
  const tag = color.kind === 'theme' ? 'a:schemeClr' : 'a:srgbClr';
  const val = color.kind === 'theme' ? color.slot : hex6(color.hex);
  return transforms === '' ? `<${tag} val="${val}"/>` : `<${tag} val="${val}">${transforms}</${tag}>`;
}

export const highlightXml = (color: PptxColor): string => `<a:highlight>${colorXml(color)}</a:highlight>`;

export const solidFillXml = (color: PptxColor): string => `<a:solidFill>${colorXml(color)}</a:solidFill>`;

/** A relationship of a `.rels` part. */
export interface RelEntry {
  id: string;
  /** The last segment of the relationship type (`slide`, `slideLayout`, `theme`...). */
  type: string;
  target: string;
  external?: boolean;
  /** The whole type URI, for the package-level relationships that are not under the officeDocument namespace. */
  typeUri?: string;
}

export const relsXml = (list: RelEntry[]): string =>
  `${XML_DECLARATION}<Relationships xmlns="${NS_PKG_REL}">${list
    .map((r) => `<Relationship${attrs({ Id: r.id, Type: r.typeUri ?? `${NS_R}/${r.type}`, Target: r.target, TargetMode: r.external ? 'External' : undefined })}/>`)
    .join('')}</Relationships>`;

/** `a:xfrm` (or `p:xfrm`, `tag`) of a frame: offset and size in EMU, rotation and flips when set. */
export const xfrmXml = (f: PptxFrame, tag = 'a:xfrm'): string =>
  `<${tag}${attrs({ rot: f.rot ? angle(f.rot) : undefined, flipH: f.flipH ? 1 : undefined, flipV: f.flipV ? 1 : undefined })}><a:off x="${emu(f.x)}" y="${emu(f.y)}"/><a:ext cx="${emu(f.w)}" cy="${emu(f.h)}"/></${tag}>`;

/** The `p:nvGrpSpPr` and `p:grpSpPr` that open every shape tree: the group transform of the root is the identity. */
export const TREE_HEADER =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
