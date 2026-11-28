import type { PptxLink } from '@b32nio/spindle-shared/export/pptx';
import { LINK_SCHEMES, schemeOf } from './links';
import { attrs, type RelEntry } from './writer-xml';

// What the slide and layout writers share (plan K12b): the relationships of the part being written, the count of what the
// writer could not state (so `warnings` stays honest) and hyperlinks.

/** The relationships of one part: the first is the one the part always has (a slide's layout, a layout's master); the others are added as needed and shared by target. */
export class PartRels {
  private readonly list: RelEntry[];
  private readonly byKey = new Map<string, string>();
  constructor(first: RelEntry) {
    this.list = [{ ...first, id: 'rId1' }];
  }
  /** The `rId` of a relationship, adding it on first use. */
  add(type: string, target: string, external = false): string {
    const key = `${type}\u0000${external}\u0000${target}`;
    const known = this.byKey.get(key);
    if (known !== undefined) return known;
    const id = `rId${this.list.length + 1}`;
    this.list.push({ id, type, target, ...(external ? { external: true } : {}) });
    this.byKey.set(key, id);
    return id;
  }
  get entries(): RelEntry[] {
    return this.list;
  }
}

/** What a writer can leave out or change within an element it writes, as keys of `LOSSY`. */
export type LossyKey = 'customGeometry' | 'pictureFill' | 'customDash' | 'link' | 'geometryName' | 'underline' | 'picture' | 'modernComment';

export const LOSSY: Record<LossyKey, string> = {
  customGeometry: 'shapes with custom geometry the editor could not keep written without their outline',
  pictureFill: 'picture fills left out (the image of a fill is not kept)',
  customDash: 'custom dash patterns written as solid lines',
  link: 'hyperlinks with an address the editor does not allow left out',
  geometryName: 'shapes with an unusable preset name written as rectangles',
  underline: 'underline styles outside the format written as single',
  picture: 'pictures with no embedded image PowerPoint can show (linked, missing, EMF, WMF, WebP, AVIF or SVG without a PNG) left out',
  modernComment: 'modern comments written as legacy comments on the slide (status and the shape a comment was anchored to are not kept)',
};

export interface WriteContext {
  rels: PartRels;
  /** The relationship target (from the part being written) of the slide that was the part `partName` in the source, when it is in the deck. */
  slideTarget: (partName: string) => string | undefined;
  lossy: (key: LossyKey) => void;
  /** The relationship target (from the part being written) of the media part holding an image given as a data URL (it keeps the media part name `part` when that is usable), which is registered once per distinct image; `undefined` when the data URL is not a base64 image PowerPoint can show. */
  image: (dataUrl: string, part?: string) => string | undefined;
  /** Names a table style (`a:tableStyleId`) so the package lists it in `ppt/tableStyles.xml`. */
  tableStyle: (id: string, name: string | undefined) => void;
  /** A GUID for the next `a:fld`: fields need one, the model does not keep it. */
  nextFieldId: () => string;
}

/** The actions the editor understands; others (macros, programs) are not written, and a `file:` target is never made a live link. */
const ACTIONS = /^ppaction:\/\/(hlinksldjump|hlinkshowjump|hlinkfile)(\?|$)/i;

/** The `a:hlinkClick` of a link, or an empty string when the link cannot be written (an address the editor's link rules do not allow). */
export function hlinkXml(link: PptxLink | undefined, ctx: WriteContext): string {
  if (!link) return '';
  const tooltip = link.tooltip;
  if (link.kind === 'url') {
    const scheme = schemeOf(link.href);
    if (scheme === undefined || !LINK_SCHEMES.has(scheme)) {
      ctx.lossy('link');
      return '';
    }
    return `<a:hlinkClick${attrs({ 'r:id': ctx.rels.add('hyperlink', link.href, true), tooltip })}/>`;
  }
  if (ACTIONS.test(link.target)) {
    const slide = link.slidePart !== undefined ? ctx.slideTarget(link.slidePart) : undefined;
    return `<a:hlinkClick${attrs({ 'r:id': slide !== undefined ? ctx.rels.add('slide', slide) : undefined, action: link.target, tooltip })}/>`;
  }
  ctx.lossy('link');
  return '';
}
