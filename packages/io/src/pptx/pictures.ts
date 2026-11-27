import type { PptxPicture } from '@b32nio/spindle-shared/export/pptx';
import type { OoxmlPackage } from '../common/ooxml-package';
import { isTrue, kid, type XmlElement } from '../xlsx/xml';
import { parseGeometry, parseLine } from './fill';
import { frameOf, matchLayoutPlaceholder } from './levels';
import { shapeLink } from './links';
import { readPlaceholder, type ShapeContext } from './shapes';
import { pct } from './units';

// Pictures (plan K10): `p:pic` with its embedded image as a data URL, crop, frame and outline. Linked pictures are never fetched.

export interface PictureContext extends ShapeContext {
  pkg: OoxmlPackage;
  /** Images already read, by part: a picture used on several slides is encoded once. */
  images: Map<string, { mime: string; dataUrl: string } | { unsupported: Unsupported }>;
}

const NO_RENDERER_WHY = 'is in EMF or WMF format, which browsers cannot show';
/** Why a picture is a placeholder, as the warning says it. */
const WHY = {
  emf: NO_RENDERER_WHY,
  wmf: NO_RENDERER_WHY,
  linked: 'is linked from outside the file, which is never fetched',
  missing: 'has no readable image',
  format: 'is in a format browsers cannot show',
} as const;
type Unsupported = keyof typeof WHY;

/** Formats with no browser renderer, by extension (judged before the bytes are read). */
const NO_RENDERER = /\.(emf|wmf|emz|wmz)$/i;

/** The MIME type of an image from its first bytes (the extension is not trusted), or `undefined` when it is none a browser shows. */
export function sniffImage(b: Uint8Array): string | undefined {
  const at = (i: number, s: string): boolean => [...s].every((c, k) => b[i + k] === c.charCodeAt(0));
  if (b[0] === 0x89 && at(1, 'PNG')) return 'image/png';
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'image/jpeg';
  if (at(0, 'GIF8')) return 'image/gif';
  if (at(0, 'RIFF') && at(8, 'WEBP')) return 'image/webp';
  // a BMP file header is 14 bytes then a DIB header of at least 12; the reserved fields (offsets 6 to 9) are zero
  if (at(0, 'BM') && b.length >= 26 && b[6] === 0 && b[7] === 0 && b[8] === 0 && b[9] === 0 && b[14] >= 12) return 'image/bmp';
  if (at(4, 'ftypavif')) return 'image/avif';
  // an SVG is text: skip a byte-order mark and an XML declaration or comment before the root
  const head = new TextDecoder().decode(b.subarray(0, 512)).replace(/^\uFEFF/, '').replace(/^\s*(?:<\?xml[^>]*\?>|<!--[\s\S]*?-->|<!DOCTYPE[^>]*>|\s)*/i, '');
  return /^<svg[\s>/]/i.test(head) ? 'image/svg+xml' : undefined;
}

function base64(bytes: Uint8Array): string {
  let out = '';
  for (let i = 0; i < bytes.length; i += 0x8000) out += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(out);
}

async function readImage(part: string, ctx: PictureContext): Promise<{ mime: string; dataUrl: string } | { unsupported: Unsupported }> {
  const cached = ctx.images.get(part);
  if (cached) return cached;
  let result: { mime: string; dataUrl: string } | { unsupported: Unsupported };
  if (NO_RENDERER.test(part)) result = { unsupported: /\.(emf|emz)$/i.test(part) ? 'emf' : 'wmf' };
  else {
    const bytes = await ctx.pkg.bytes(part);
    const mime = bytes === undefined ? undefined : sniffImage(bytes);
    result = bytes === undefined ? { unsupported: 'missing' } : mime === undefined ? { unsupported: 'format' } : { mime, dataUrl: `data:${mime};base64,${base64(bytes)}` };
  }
  ctx.images.set(part, result);
  return result;
}

/** An embedded image as a data URL, or `undefined` when it is missing, in a format browsers cannot show or has no renderer (cached per part, like pictures). */
export async function imageDataUrl(part: string, ctx: PictureContext): Promise<string | undefined> {
  const image = await readImage(part, ctx);
  return 'dataUrl' in image ? image.dataUrl : undefined;
}

/** `a:srcRect` in percent of the source (`l="25000"` is 25), only when it trims something; an outset (a negative edge) counts as none, and so does a crop whose opposite edges meet or overlap. */
function readCrop(rect: XmlElement | undefined): PptxPicture['crop'] {
  const edge = (a: string): number => Math.min(100, Math.max(0, (pct(rect?.attrs[a]) ?? 0) * 100));
  const crop = { l: edge('l'), t: edge('t'), r: edge('r'), b: edge('b') };
  // edges that together trim the whole image (or more) leave nothing to show: ignored
  if (crop.l + crop.r >= 100 || crop.t + crop.b >= 100) return undefined;
  return crop.l || crop.t || crop.r || crop.b ? crop : undefined;
}

export async function readPicture(e: XmlElement, ctx: PictureContext): Promise<PptxPicture> {
  const nv = kid(e, 'nvPicPr');
  const cNvPr = kid(nv, 'cNvPr');
  const placeholder = readPlaceholder(kid(kid(nv, 'nvPr'), 'ph'));
  const match = placeholder ? matchLayoutPlaceholder(ctx.layout, placeholder) : undefined;
  const spPr = kid(e, 'spPr');
  const blipFill = kid(e, 'blipFill');
  const blip = kid(blipFill, 'blip');
  const rel = ctx.rels.get(blip?.rattrs?.embed ?? '');
  const linked = (!!blip?.rattrs?.link && !rel) || !!rel?.external;
  const image: Awaited<ReturnType<typeof readImage>> = rel && !rel.external ? await readImage(rel.target, ctx) : { unsupported: linked ? 'linked' : 'missing' };
  if ('unsupported' in image) {
    ctx.warn(`A picture ${WHY[image.unsupported]}; it is kept as a placeholder.`);
  }
  const altText = (cNvPr?.attrs.descr ?? '').trim();
  const frame = frameOf(kid(spPr, 'xfrm')) ?? match?.frame;
  const geometry = parseGeometry(spPr);
  const line = parseLine(kid(spPr, 'ln'), ctx.map, ctx.rels);
  const crop = readCrop(kid(blipFill, 'srcRect'));
  const alpha = pct(kid(blip, 'alphaModFix')?.attrs.amt);
  const link = shapeLink(cNvPr, { rels: ctx.rels, warn: ctx.warn });
  return {
    kind: 'picture',
    pptxId: cNvPr?.attrs.id ?? '',
    name: (cNvPr?.attrs.name ?? '').trim(),
    ...(altText !== '' ? { altText } : {}),
    ...(isTrue(cNvPr?.attrs.hidden) ? { hidden: true } : {}),
    ...(placeholder ? { placeholder } : {}),
    ...(frame ? { frame } : {}),
    ...(geometry ? { geometry } : {}),
    ...(line ? { line } : {}),
    ...(rel && !rel.external ? { part: rel.target } : {}),
    ...('mime' in image ? { mime: image.mime, dataUrl: image.dataUrl } : { unsupported: image.unsupported }),
    ...(crop ? { crop } : {}),
    ...(alpha !== undefined && alpha < 1 ? { alpha: Math.max(0, alpha) } : {}),
    ...(link ? { link } : {}),
  };
}

