import { XMLValidator } from 'fast-xml-parser';
import type { OpaquePart, OpaqueRef } from '@b32nio/spindle-shared/opaque';
import type { PptxElement, PptxFrame, PptxOpaque, PptxSlideOpaque } from '@b32nio/spindle-shared/export/pptx';
import { relsPart, resolveTarget } from '../common/ooxml-package';
import type { LossyKey, PartRels, WriteContext } from './writer-context';
import { attrs, esc, NS_R, xfrmXml } from './writer-xml';

// Re-emitting the opaque content (plan K14, decision D5): charts, SmartArt, OLE objects, audio and video, and ink come back as
// the frame that names their parts and as the parts themselves, byte for byte from the opaque part store (X5). The part names are
// kept when the package has no use for them (so a deck read and written again states the same part references); a part named like
// one of the writer's own moves under `ppt/opaque/`, a taken name gets a number. Relationship parts are rewritten only where a
// target changed name or points at a part that was not kept. Transitions and the timing tree of a slide are written as they were
// read, when they are well formed, need no part the file does not carry and name only shapes the slide holds under the same id.

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const URI = {
  chart: NS_C,
  diagram: NS_DGM,
  ole: 'http://schemas.openxmlformats.org/presentationml/2006/ole',
};
const MEDIA_EXT = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}';
/** The relationship types that live outside the officeDocument namespace, by the last segment the store keeps as the role. */
const MS_RELS = 'http://schemas.microsoft.com/office/2007/relationships';
const EXTENSION_RELS = new Set(['diagramDrawing', 'media', 'ink']);
const CLIPS = new Set(['audioFile', 'videoFile', 'quickTimeFile']);
/** Names the writer gives its own parts (and `[Content_Types].xml`): an opaque part with one moves to `ppt/opaque/`. */
const RESERVED = /^(?:\[content_types\]\.xml$|_rels\/|docprops\/|ppt\/(?:_rels\/|slides\/|slidelayouts\/|slidemasters\/|notesslides\/|notesmasters\/|comments\/|theme\/theme\d+\.xml$|(?:presentation|presprops|viewprops|tablestyles|commentauthors)\.xml$))/i;
/** What a part name may hold to be written as it is: no space, no percent escape, no climbing out of the package. */
const SAFE_NAME = /^(?!\/)(?!.*(?:^|\/)\.\.?(?:\/|$))[A-Za-z0-9._\-/]+$/;
const REL_TAG = /<Relationship\b(?:"[^"]*"|'[^']*'|[^>"'])*>(?:\s*<\/Relationship\s*>)?/g;
const ATTR = (name: string): RegExp => new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`);
const DECODER = new TextDecoder();
const ENCODER = new TextEncoder();

/** A part to write: `type` is the content type of its `Override` (a relationship part has none, `.rels` is a `Default`). */
export interface PartEntry {
  name: string;
  data: Uint8Array;
  type?: string;
}

export interface PlacedRoot {
  role: string;
  /** The relationship target from the part that holds the element. */
  target: string;
}

export interface PlacedOpaque {
  roots: PlacedRoot[];
  /** The id the diagram data gives the relationship of its cached drawing (`dsp:dataModelExt relId`), which the slide should use for it. */
  drawingRelId?: string;
}

export interface OpaqueEmitter {
  /** The relationship targets of an element's root parts, with its parts queued for the package; `undefined` (and reported) when the element cannot be written. */
  place(e: PptxOpaque): PlacedOpaque | undefined;
  /** Claims the relationship ids the diagram data of the slide's elements names, before the other relationships take them. */
  reserve(elements: PptxElement[], rels: PartRels): void;
}

const attr = (tag: string, name: string): string | undefined => {
  const m = ATTR(name).exec(tag);
  return m ? (m[1] ?? m[2]) : undefined;
};
const unescapeXml = (s: string): string => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** `to` as a relationship target from the part `from`. */
export function relativeTarget(from: string, to: string): string {
  const base = from.split('/').slice(0, -1);
  const dest = to.split('/');
  let common = 0;
  while (common < base.length && common < dest.length - 1 && base[common] === dest[common]) common++;
  return [...base.slice(common).map(() => '..'), ...dest.slice(common)].join('/');
}

/** The part names the internal relationships of a relationship part (owned by `owner`) point at. */
function relTargets(bytes: Uint8Array, owner: string): string[] {
  const out: string[] = [];
  for (const tag of DECODER.decode(bytes).match(REL_TAG) ?? []) {
    const target = attr(tag, 'Target');
    if (target !== undefined && !/^external$/i.test(attr(tag, 'TargetMode') ?? '')) out.push(resolveTarget(owner, unescapeXml(target)));
  }
  return out;
}

/**
 * The relationship part with its targets rewritten: `owner` is the name the owning part had, `placed` the name it has now,
 * `names` maps the names of the element's parts to the names they have now. A relationship to a part that was not kept is left out.
 * The bytes are returned as they were when every target still resolves to the same part.
 */
function rewriteRels(bytes: Uint8Array, owner: string, placed: string, names: Map<string, string>): { bytes: Uint8Array; dropped: number } {
  let dropped = 0;
  let changed = false;
  const text = DECODER.decode(bytes).replace(REL_TAG, (tag) => {
    const raw = attr(tag, 'Target');
    if (raw === undefined || /^external$/i.test(attr(tag, 'TargetMode') ?? '')) return tag;
    const now = names.get(resolveTarget(owner, unescapeXml(raw)));
    if (now === undefined) {
      dropped++;
      changed = true;
      return '';
    }
    if (resolveTarget(placed, unescapeXml(raw)) === now) return tag;
    changed = true;
    return tag.replace(ATTR('Target'), () => ` Target="${esc(relativeTarget(placed, now))}"`);
  });
  return { bytes: changed ? ENCODER.encode(text) : bytes, dropped };
}

const stem = (name: string): { dir: string; base: string; ext: string } => {
  const slash = name.lastIndexOf('/');
  const file = name.slice(slash + 1);
  const dot = file.lastIndexOf('.');
  return { dir: name.slice(0, slash + 1), base: dot > 0 ? file.slice(0, dot) : file, ext: dot > 0 ? file.slice(dot) : '' };
};

/**
 * The opaque parts of a deck on their way into the package: names them (once per distinct part) and keeps the entries to write.
 * `taken` holds the lower-case names the package already uses for media; it is shared with the image writer so neither takes the other's.
 */
export class OpaqueParts {
  readonly entries: PartEntry[] = [];
  private readonly store = new Map<string, OpaquePart>();
  private readonly names = new Map<string, string>();
  private readonly written = new Set<string>();
  private generic = 0;
  constructor(
    parts: OpaquePart[],
    private readonly taken: Set<string>
  ) {
    for (const p of parts) this.store.set(p.hash, p);
  }

  private nameOf(ref: OpaqueRef): string {
    const key = `${ref.hash}\u0000${ref.partName ?? ''}`;
    const known = this.names.get(key);
    if (known !== undefined) return known;
    let wanted = ref.partName !== undefined && SAFE_NAME.test(ref.partName) ? ref.partName : `ppt/opaque/part${++this.generic}${stem(ref.partName ?? '').ext.replace(/[^A-Za-z0-9.]/g, '')}`;
    if (RESERVED.test(wanted)) wanted = `ppt/opaque/${stem(wanted).base}${stem(wanted).ext}`;
    const { dir, base, ext } = stem(wanted);
    let name = wanted;
    for (let n = 2; this.taken.has(name.toLowerCase()); n++) name = `${dir}${base}_${n}${ext}`;
    this.taken.add(name.toLowerCase());
    this.names.set(key, name);
    return name;
  }

  private emit(entry: PartEntry): void {
    if (this.written.has(entry.name.toLowerCase())) return;
    this.written.add(entry.name.toLowerCase());
    this.entries.push(entry);
  }

  /** The emitter of the part `owner` (a slide): the targets it computes are relative to it. */
  forPart(owner: string, lossy: (key: LossyKey) => void): OpaqueEmitter {
    const cache = new Map<PptxOpaque, PlacedOpaque | undefined>();
    const place = (e: PptxOpaque): PlacedOpaque | undefined => {
      if (cache.has(e)) return cache.get(e);
      const placed = this.placeElement(e, owner, lossy);
      cache.set(e, placed);
      return placed;
    };
    const reserve = (elements: PptxElement[], rels: PartRels): void => {
      for (const e of elements) {
        if (e.kind === 'group') reserve(e.children, rels);
        else if (e.kind === 'opaque' && e.opaqueKind === 'smartArt') {
          const placed = place(e);
          const drawing = placed?.roots.find((r) => r.role === 'diagramDrawing');
          if (drawing) rels.add(relType('diagramDrawing'), drawing.target, false, placed?.drawingRelId);
        }
      }
    };
    return { place, reserve };
  }

  private placeElement(e: PptxOpaque, owner: string, lossy: (key: LossyKey) => void): PlacedOpaque | undefined {
    const refused = e.props?.linkedTarget !== undefined ? 'opaqueLinked' : e.props?.control === 'true' ? 'opaqueControl' : undefined;
    if (refused) {
      lossy(refused);
      return undefined;
    }
    const present: Array<{ ref: OpaqueRef; part: OpaquePart }> = [];
    let missing = false;
    for (const ref of e.parts) {
      const part = this.store.get(ref.hash);
      if (part) present.push({ ref, part });
      else missing = true;
    }
    const content = present.filter((p) => p.ref.role !== 'rels');
    const names = new Map<string, string>();
    for (const { ref } of content) if (ref.partName !== undefined) names.set(ref.partName, this.nameOf(ref));
    // roots: the parts no other part of the element points at (the frame's own relationships), in the order the element lists them
    const pointedAt = new Set<string>();
    for (const { ref, part } of present) {
      if (ref.role === 'rels' && ref.partName !== undefined) for (const t of relTargets(part.bytes, ref.partName.replace(/_rels\/([^/]+)\.rels$/, '$1'))) pointedAt.add(t);
    }
    const roots = content.filter(({ ref }) => ref.partName === undefined || !pointedAt.has(ref.partName));
    if (roots.length === 0 && content.length > 0) roots.push(content[0]);
    const placed: PlacedRoot[] = roots.map(({ ref }) => ({ role: ref.role, target: relativeTarget(owner, this.nameOf(ref)) }));
    const has = (role: string): boolean => placed.some((r) => r.role === role);
    const ok =
      e.opaqueKind === 'chart' ? has('chart') : e.opaqueKind === 'smartArt' ? has('diagramData') : e.opaqueKind === 'ink' ? placed.length > 0 : placed.some((r) => r.role !== 'image');
    if (!ok) {
      lossy(e.opaqueKind === 'chart' && has('chartEx') ? 'opaqueChartEx' : 'opaque');
      return undefined;
    }
    let dropped = missing;
    for (const { ref, part } of present) {
      const name = this.nameOf(ref);
      if (ref.role !== 'rels') this.emit({ name, data: part.bytes, type: part.contentType || 'application/octet-stream' });
      else if (ref.partName === undefined) dropped = true;
      else {
        const was = ref.partName.replace(/_rels\/([^/]+)\.rels$/, '$1');
        const now = names.get(was);
        if (now === undefined) dropped = true;
        else {
          const rewritten = rewriteRels(part.bytes, was, now, names);
          if (rewritten.dropped > 0) dropped = true;
          this.emit({ name: relsPart(now), data: rewritten.bytes });
        }
      }
    }
    if (dropped) lossy('opaquePart');
    const data = present.find(({ ref }) => ref.role === 'diagramData');
    const drawingRelId = data ? /dataModelExt\b[^>]*?\srelId\s*=\s*["']([^"']+)["']/.exec(DECODER.decode(data.part.bytes))?.[1] : undefined;
    return { roots: placed, ...(drawingRelId !== undefined ? { drawingRelId } : {}) };
  }
}

/** The relationship type of a root part, from the role the store keeps. */
export const relType = (role: string): string => (EXTENSION_RELS.has(role) ? `${MS_RELS}/${role}` : role);

const cNvPrXml = (e: PptxOpaque, id: number, tag = 'p:cNvPr'): string => `<${tag}${attrs({ id, name: e.name, descr: e.altText, hidden: e.hidden ? 1 : undefined })}/>`;
const ZERO: PptxFrame = { x: 0, y: 0, w: 0, h: 0 };

function graphicFrame(e: PptxOpaque, frame: PptxFrame | undefined, id: number, uri: string, data: string): string {
  return (
    `<p:graphicFrame><p:nvGraphicFramePr>${cNvPrXml(e, id)}<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>${xfrmXml(frame ?? ZERO, 'p:xfrm')}` +
    `<a:graphic><a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic></p:graphicFrame>`
  );
}

/** `p:blipFill` of a picture of the image part related as `rId`, a stretch of the whole picture. */
const blipFill = (rId: string | undefined): string => `<p:blipFill><a:blip${attrs({ 'r:embed': rId })}/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`;
const rect = (frame: PptxFrame | undefined): string => `<p:spPr>${xfrmXml(frame ?? ZERO)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`;

/**
 * An opaque element as its frame: `p:graphicFrame` for a chart, a diagram or an OLE object, `p:pic` for audio and video,
 * `mc:AlternateContent` holding `p:contentPart` for ink. The element's parts are queued by `place`; the relationships of the
 * root parts are added to the slide here; `frame` is the element's frame in the space of the part it is written in. An element that cannot be written gives an empty string (and is counted).
 */
export function opaqueXml(e: PptxOpaque, frame: PptxFrame | undefined, id: number, ctx: WriteContext): string {
  const placed = ctx.opaque?.place(e);
  if (!placed) {
    if (!ctx.opaque) ctx.lossy('opaque');
    return '';
  }
  const rel = (role: string, wanted?: string): string | undefined => {
    const r = placed.roots.find((x) => x.role === role);
    return r ? ctx.rels.add(relType(role), r.target, false, wanted) : undefined;
  };
  const main = placed.roots.find((r) => r.role !== 'image') as PlacedRoot;
  const image = placed.roots.find((r) => r.role === 'image');
  switch (e.opaqueKind) {
    case 'chart':
      return graphicFrame(e, frame, id, URI.chart, `<c:chart${attrs({ 'xmlns:c': NS_C, 'r:id': rel('chart') })}/>`);
    case 'smartArt':
      rel('diagramDrawing', placed.drawingRelId);
      return graphicFrame(
        e,
        frame,
        id,
        URI.diagram,
        `<dgm:relIds${attrs({ 'xmlns:dgm': NS_DGM, 'r:dm': rel('diagramData'), 'r:lo': rel('diagramLayout'), 'r:qs': rel('diagramQuickStyle'), 'r:cs': rel('diagramColors') })}/>`
      );
    case 'ole': {
      const preview = image
        ? `<p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>${blipFill(ctx.rels.add('image', image.target))}${rect(frame)}</p:pic>`
        : '';
      return graphicFrame(e, frame, id, URI.ole, `<p:oleObj${attrs({ name: e.props?.oleName, progId: e.props?.progId, 'r:id': ctx.rels.add(relType(main.role), main.target) })}><p:embed/>${preview}</p:oleObj>`);
    }
    case 'media': {
      const clip = CLIPS.has(e.props?.clip ?? '') ? (e.props?.clip as string) : main.role === 'audio' ? 'audioFile' : 'videoFile';
      const link = ctx.rels.add(clip === 'audioFile' ? 'audio' : 'video', main.target);
      const embed = ctx.rels.add(relType('media'), main.target);
      return (
        `<p:pic><p:nvPicPr>${cNvPrXml(e, id)}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:${clip} r:link="${link}"/>` +
        `<p:extLst><p:ext uri="${MEDIA_EXT}"><p14:media xmlns:p14="${NS_P14}" r:embed="${embed}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>` +
        `${blipFill(image ? ctx.rels.add('image', image.target) : undefined)}${rect(frame)}</p:pic>`
      );
    }
    case 'ink':
      return (
        `<mc:AlternateContent xmlns:mc="${NS_MC}"><mc:Choice xmlns:p14="${NS_P14}" Requires="p14"><p:contentPart p14:bwMode="auto" r:id="${ctx.rels.add(relType(main.role), main.target)}">` +
        `<p14:nvContentPartPr>${cNvPrXml(e, id, 'p14:cNvPr')}<p14:cNvContentPartPr/><p14:nvPr/></p14:nvContentPartPr>${xfrmXml(frame ?? ZERO, 'p14:xfrm')}</p:contentPart></mc:Choice></mc:AlternateContent>`
      );
  }
}

const ROOT_LOCAL = /^\s*<(?:[\w.-]+:)?([\w.-]+)/;
const FORBIDDEN = /<!DOCTYPE|<!ENTITY|<\?/i;
const SOUND = /<((?:[\w.-]+:)?)sndAc\b[\s\S]*?<\/\1sndAc>/g;
const SPID = /\sspid\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

/** Whether the fragment can be put into a slide: a single well-formed element of one of the named kinds, with no declaration or instruction. */
function usable(xml: string, kinds: string[]): boolean {
  const root = ROOT_LOCAL.exec(xml)?.[1];
  return root !== undefined && kinds.includes(root) && !FORBIDDEN.test(xml) && XMLValidator.validate(xml) === true;
}

/** Whether the fragment still relates to a part (an attribute in the relationships namespace): such a part is not in the package. */
function relatesToParts(xml: string): boolean {
  const prefixes = [...xml.matchAll(/xmlns:([\w.-]+)\s*=\s*["']([^"']*)["']/g)].filter((m) => m[2] === NS_R).map((m) => m[1]);
  return prefixes.some((p) => new RegExp(`\\s${p.replace(/\./g, '\\.')}:[\\w.-]+\\s*=`).test(xml));
}

/**
 * The transition and the timing XML of a slide as the slide root's children, in schema order (`transition`, then `timing`).
 * `ids` gives the id each source shape id was written with: a timing that names a shape under another id or not at all is left
 * out (PowerPoint reports such a file as damaged), as is any fragment that is malformed or relates to a part the file does not carry (a sound);
 * a transition's sound alone is cut out of it, the transition stays.
 */
export function slideFragmentsXml(opaque: PptxSlideOpaque | undefined, ids: Map<string, number>, lossy: (key: LossyKey) => void): string {
  let out = '';
  const transition = opaque?.transition;
  if (transition !== undefined) {
    const silent = usable(transition, ['transition', 'AlternateContent']) ? transition.replace(SOUND, '') : undefined;
    if (silent === undefined || relatesToParts(silent)) lossy('transition');
    else {
      if (silent !== transition) lossy('transition');
      out += silent;
    }
  }
  const timing = opaque?.timing;
  if (timing !== undefined) {
    const named = [...timing.matchAll(SPID)].map((m) => m[1] ?? m[2]);
    if (usable(timing, ['timing']) && !relatesToParts(timing) && named.every((id) => ids.get(id) === Number(id))) out += timing;
    else lossy('timing');
  }
  return out;
}
