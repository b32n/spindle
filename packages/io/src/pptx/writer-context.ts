import type { PptxLink } from '@b32nio/spindle-shared/export/pptx';
import { LINK_SCHEMES, schemeOf } from './links';
import { attrs, type RelEntry } from './writer-xml';
import type { OpaqueEmitter } from './writer-opaque';

// What the slide and layout writers share (plan K12b): the relationships of the part being written, the count of what the
// writer could not state (so `warnings` stays honest) and hyperlinks.

/** The relationships of one part: the first is the one the part always has (a slide's layout, a layout's master); the others are added as needed and shared by target. */
export class PartRels {
  private readonly list: RelEntry[];
  private readonly byKey = new Map<string, string>();
  private readonly used = new Set<string>(['rId1']);
  /** `movable`: nothing refers to the first relationship by id (a slide's layout), so another may take `rId1` and the first moves to a free id. */
  constructor(
    first: RelEntry,
    private readonly movable = false
  ) {
    this.list = [{ ...first, id: 'rId1' }];
  }
  private free(): string {
    let n = this.list.length + 1;
    while (this.used.has(`rId${n}`)) n++;
    return `rId${n}`;
  }
  /**
   * The `rId` of a relationship, adding it on first use. A `type` that is a whole URI (a Microsoft extension relationship) is
   * written as it is, any other is the last segment of an officeDocument relationship type. `wanted` asks for that id when it is
   * free (a part names a relationship of the slide by id, as the diagram data does for its drawing); else the next free one is used.
   */
  add(type: string, target: string, external = false, wanted?: string): string {
    const key = `${type}\u0000${external}\u0000${target}`;
    const known = this.byKey.get(key);
    if (known !== undefined) return known;
    if (this.movable && wanted === 'rId1' && this.list[0].id === 'rId1') {
      this.list[0] = { ...this.list[0], id: this.free() };
      this.used.add(this.list[0].id);
      this.used.delete('rId1');
    }
    const id = wanted !== undefined && !this.used.has(wanted) ? wanted : this.free();
    this.used.add(id);
    const whole = /^[a-z][a-z0-9+.-]*:\/\//i.test(type);
    this.list.push({ id, type: whole ? type.slice(type.lastIndexOf('/') + 1) : type, target, ...(whole ? { typeUri: type } : {}), ...(external ? { external: true } : {}) });
    this.byKey.set(key, id);
    return id;
  }
  get entries(): RelEntry[] {
    return this.list;
  }
}

/** What a writer can leave out or change within an element it writes, as keys of `LOSSY`. */
export type LossyKey = 'opaque' | 'opaqueLinked' | 'opaqueControl' | 'opaquePart' | 'transition' | 'timing' | 'customGeometry' | 'pictureFill' | 'customDash' | 'link' | 'geometryName' | 'underline' | 'picture' | 'modernComment';

export const LOSSY: Record<LossyKey, string> = {
  opaque: 'charts, diagrams, objects, media and ink left out because the parts they are made of are missing',
  opaqueLinked: 'media and objects linked from outside the file left out (the link is never written)',
  opaqueControl: 'ActiveX controls left out (their binary parts are not written)',
  opaquePart: 'relationships of kept content to parts that are missing left out',
  transition: 'transitions left out (the XML is malformed or holds a sound the file does not carry)',
  timing: 'animation timelines left out (they name shapes the slide does not hold, or hold sounds the file does not carry)',
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
  /** Writes the parts of opaque elements (charts, diagrams, objects, media, ink); only the context of a slide has it (plan K14). */
  opaque?: OpaqueEmitter;
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
