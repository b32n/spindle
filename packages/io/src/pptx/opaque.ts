import type { OpaqueRef, OpaqueStore } from '@b32nio/spindle-shared/opaque';
import type { PptxFrame, PptxOpaque, PptxOpaqueKind } from '@b32nio/spindle-shared/export/pptx';
import { relsPart, type OoxmlPackage } from '../common/ooxml-package';
import { descendants, isTrue, kid, kids, type XmlElement } from '../xlsx/xml';
import { frameOf } from './levels';
import { imageDataUrl, readPicture, type PictureContext } from './pictures';

// Content the editor neither edits nor draws (plan K11, decision D5): charts, SmartArt, OLE objects and controls, audio and video, ink.
// Each becomes an opaque element holding references to its parts, which are kept byte for byte in the opaque part store (X5).

/** The part types of the document's own structure: a relationship to one of these is never part of an opaque element. */
const STRUCTURAL = new Set(['slide', 'slideLayout', 'slideMaster', 'notesSlide', 'notesMaster', 'handoutMaster', 'presentation', 'comments', 'hyperlink']);
/** How far relationships are followed from the root part of an element (a chart reaches its workbook in one step), and how many parts one element may gather. */
const MAX_DEPTH = 8;
const MAX_PARTS = 256;

/** `[Content_Types].xml`: the content type of a part, by override then by extension. */
export interface ContentTypes {
  of(part: string): string;
}

export function readContentTypes(root: XmlElement | undefined): ContentTypes {
  const overrides = new Map<string, string>();
  const defaults = new Map<string, string>();
  for (const o of kids(root, 'Override')) overrides.set((o.attrs.PartName ?? '').replace(/^\//, '').toLowerCase(), o.attrs.ContentType ?? '');
  for (const d of kids(root, 'Default')) defaults.set((d.attrs.Extension ?? '').toLowerCase(), d.attrs.ContentType ?? '');
  return {
    of: (part) => overrides.get(part.toLowerCase()) ?? defaults.get(part.slice(part.lastIndexOf('.') + 1).toLowerCase()) ?? 'application/octet-stream',
  };
}

export interface OpaqueContext {
  pkg: OoxmlPackage;
  store: OpaqueStore;
  types: ContentTypes;
  /** Parts already offered to the store, by name: the reference, or `null` when the part is missing or the store refused it. */
  stored: Map<string, OpaqueRef | null>;
}

export interface OpaqueElementContext extends PictureContext {
  opaque: OpaqueContext;
}

/** Offer one part to the store (once per import, however many elements refer to it). `null`: missing, or refused (a warning says which). */
async function storePart(o: OpaqueContext, part: string, role: string, warn: (m: string) => void): Promise<OpaqueRef | null> {
  const known = o.stored.get(part);
  if (known !== undefined) return known;
  const bytes = await o.pkg.bytes(part);
  let ref: OpaqueRef | null = null;
  if (bytes === undefined) warn(`${part} could not be read (it is missing or too large to read) and is not kept.`);
  else {
    const put = await o.store.put({ bytes, contentType: o.types.of(part), role, partName: part });
    if (put.ok) ref = put.ref;
    else warn(`${part} is ${put.reason === 'too-large' ? 'over the size limit for a kept part' : 'not kept: the deck’s kept parts are over the size limit'} and is dropped.`);
  }
  o.stored.set(part, ref);
  return ref;
}

/**
 * The parts reachable from the root parts through relationships, each followed by its own relationship part (role `rels`, so an
 * exporter can rewrite the targets). A relationship to the document's own structure, an external one and one to a missing part is not followed.
 */
async function collectParts(o: OpaqueContext, roots: Array<{ part: string; role: string }>, warn: (m: string) => void): Promise<OpaqueRef[]> {
  const refs: OpaqueRef[] = [];
  const seen = new Set<string>();
  let queue = roots.map((r) => ({ ...r, depth: 0 }));
  let truncated = false;
  while (queue.length > 0) {
    const next: typeof queue = [];
    for (const { part, role, depth } of queue) {
      if (seen.has(part)) continue;
      if (seen.size >= MAX_PARTS) {
        truncated = true;
        continue;
      }
      seen.add(part);
      const ref = await storePart(o, part, role, warn);
      if (!ref) continue;
      refs.push(ref);
      const rels = relsPart(part);
      const relsRef = o.pkg.has(rels) ? await storePart(o, rels, 'rels', warn) : null;
      if (relsRef) refs.push(relsRef);
      const links = (await o.pkg.rels(part)).filter((r) => !r.external && !STRUCTURAL.has(r.type) && o.pkg.has(r.target));
      if (depth < MAX_DEPTH) for (const r of links) next.push({ part: r.target, role: r.type, depth: depth + 1 });
      else if (links.some((r) => !seen.has(r.target))) truncated = true;
    }
    queue = next;
  }
  if (truncated) warn(`An element refers to more than ${MAX_PARTS} parts or nests them more than ${MAX_DEPTH} levels deep: the rest is not kept.`);
  return refs;
}

const nameOf = (cNvPr: XmlElement | undefined): { pptxId: string; name: string; altText?: string; hidden?: true } => {
  const altText = (cNvPr?.attrs.descr ?? '').trim();
  return { pptxId: cNvPr?.attrs.id ?? '', name: (cNvPr?.attrs.name ?? '').trim(), ...(altText !== '' ? { altText } : {}), ...(isTrue(cNvPr?.attrs.hidden) ? { hidden: true as const } : {}) };
};

/** A chart's title as plain text (`c:chart/c:title`), when it states one. */
async function chartTitle(o: OpaqueContext, part: string): Promise<string | undefined> {
  const title = kid(kid(await o.pkg.xml(part), 'chart'), 'title');
  const text = descendants(title, 't').map((t) => t.text).join('').trim();
  return text === '' ? undefined : text;
}

/** What kind of opaque content a `p:graphicFrame` holds, or `undefined` for one this reader does not keep (a table is read elsewhere). */
export function opaqueFrameKind(frame: XmlElement): PptxOpaqueKind | undefined {
  const uri = kid(kid(frame, 'graphic'), 'graphicData')?.attrs.uri ?? '';
  if (/chart/i.test(uri)) return 'chart';
  if (/\/diagram$/i.test(uri)) return 'smartArt';
  if (/\/ole$/i.test(uri)) return 'ole';
  return undefined;
}

const roleOf = (rels: OpaqueElementContext['rels'], id: string | undefined): { part: string; role: string } | undefined => {
  const r = rels.get(id ?? '');
  return r && !r.external && !STRUCTURAL.has(r.type) ? { part: r.target, role: r.type } : undefined;
};

/** A `p:graphicFrame` holding a chart, a SmartArt diagram or an OLE object. */
export async function readOpaqueFrame(e: XmlElement, kind: PptxOpaqueKind, ctx: OpaqueElementContext): Promise<PptxOpaque> {
  const data = kid(kid(e, 'graphic'), 'graphicData') as XmlElement;
  const base = { kind: 'opaque' as const, opaqueKind: kind, ...nameOf(kid(kid(e, 'nvGraphicFramePr'), 'cNvPr')) };
  const frame = frameOf(kid(e, 'xfrm'));
  const roots: Array<{ part: string; role: string }> = [];
  const props: Record<string, string> = {};
  let poster: string | undefined;
  let title: string | undefined;
  const add = (r: { part: string; role: string } | undefined): void => {
    if (r) roots.push(r);
  };
  if (kind === 'chart') {
    const root = roleOf(ctx.rels, descendants(data, 'chart').find((c) => c.rattrs?.id)?.rattrs?.id);
    add(root);
    if (root) title = await chartTitle(ctx.opaque, root.part);
  } else if (kind === 'smartArt') {
    const ids = descendants(data, 'relIds')[0]?.rattrs ?? {};
    for (const id of [ids.dm, ids.lo, ids.qs, ids.cs]) add(roleOf(ctx.rels, id));
    // the cached drawing is related from the slide; the data part names its relationship id
    const dm = roleOf(ctx.rels, ids.dm);
    const drawing = dm ? descendants(await ctx.opaque.pkg.xml(dm.part), 'dataModelExt')[0]?.attrs.relId : undefined;
    add(roleOf(ctx.rels, drawing));
  } else {
    const ole = descendants(data, 'oleObj')[0];
    const target = ctx.rels.get(ole?.rattrs?.id ?? '');
    if (target?.external) props.linkedTarget = target.raw;
    else add(roleOf(ctx.rels, ole?.rattrs?.id));
    if (ole?.attrs.progId) props.progId = ole.attrs.progId;
    if (ole?.attrs.name) props.oleName = ole.attrs.name;
    const preview = roleOf(ctx.rels, descendants(ole, 'blip').find((b) => b.rattrs?.embed)?.rattrs?.embed);
    if (preview) {
      add(preview);
      poster = await imageDataUrl(preview.part, ctx);
    }
  }
  return finish(base, frame, roots, props, ctx, { poster, title });
}

async function finish(
  base: Omit<PptxOpaque, 'parts'>,
  frame: PptxFrame | undefined,
  roots: Array<{ part: string; role: string }>,
  props: Record<string, string>,
  ctx: OpaqueElementContext,
  extra: { poster?: string; title?: string }
): Promise<PptxOpaque> {
  const parts = await collectParts(ctx.opaque, roots, ctx.warn);
  return {
    ...base,
    ...(frame ? { frame } : {}),
    ...(extra.title ? { title: extra.title } : {}),
    parts,
    ...(extra.poster ? { poster: extra.poster } : {}),
    ...(Object.keys(props).length > 0 ? { props } : {}),
  };
}

const MEDIA_FILES = new Set(['audioFile', 'videoFile', 'quickTimeFile']);

/** The clip element (`a:audioFile`, `a:videoFile`, `a:quickTimeFile`) of a `p:pic` that is an audio or video clip. */
export const mediaClip = (e: XmlElement): XmlElement | undefined => kid(kid(e, 'nvPicPr'), 'nvPr')?.children.find((k) => MEDIA_FILES.has(k.name));

/**
 * An audio or video clip (`p:pic` with a clip element): the media part (named by the clip's own relationship and by the newer
 * `p14:media`), with the poster frame, which is the picture, as the placeholder's image. A clip linked from outside the package is never fetched: its address is kept inert in `props`.
 */
export async function readMedia(e: XmlElement, clip: XmlElement, ctx: OpaqueElementContext): Promise<PptxOpaque> {
  const pic = await readPicture(e, ctx);
  const roots: Array<{ part: string; role: string }> = [];
  const props: Record<string, string> = { clip: clip.name };
  const link = ctx.rels.get(clip.rattrs?.link ?? '');
  if (link?.external) props.linkedTarget = link.raw;
  const ids = [clip.rattrs?.link, descendants(kid(kid(e, 'nvPicPr'), 'nvPr'), 'media').find((m) => m.rattrs?.embed)?.rattrs?.embed];
  for (const id of ids) {
    const r = roleOf(ctx.rels, id);
    if (r) roots.push(r);
  }
  const poster = roleOf(ctx.rels, kid(kid(e, 'blipFill'), 'blip')?.rattrs?.embed);
  if (poster) roots.push({ part: poster.part, role: 'image' });
  const { pptxId, name, altText, hidden, frame } = pic;
  return finish({ kind: 'opaque', opaqueKind: 'media', pptxId, name, ...(altText ? { altText } : {}), ...(hidden ? { hidden } : {}) }, frame, roots, props, ctx, { poster: pic.dataUrl });
}

/** Ink (`p:contentPart`, the `p14` choice of a `mc:AlternateContent`): the ink part its relationship names. */
export async function readInk(e: XmlElement, ctx: OpaqueElementContext): Promise<PptxOpaque> {
  const root = roleOf(ctx.rels, e.rattrs?.id);
  return finish({ kind: 'opaque', opaqueKind: 'ink', ...nameOf(kid(kid(e, 'nvContentPartPr'), 'cNvPr')) }, frameOf(kid(e, 'xfrm')), root ? [root] : [], {}, ctx, {});
}

/** An ActiveX control (`p:control` in `p:cSld/p:controls`): the control part and its picture, drawn in the placeholder. */
export async function readControl(e: XmlElement, ctx: OpaqueElementContext): Promise<PptxOpaque> {
  const pic = kid(e, 'pic');
  const root = roleOf(ctx.rels, e.rattrs?.id);
  const preview = roleOf(ctx.rels, descendants(pic, 'blip').find((b) => b.rattrs?.embed)?.rattrs?.embed);
  const { pptxId, name, altText, hidden } = nameOf(kid(kid(pic, 'nvPicPr'), 'cNvPr'));
  return finish(
    { kind: 'opaque', opaqueKind: 'ole', pptxId: e.attrs.spid ?? pptxId, name: (e.attrs.name ?? name).trim(), ...(altText ? { altText } : {}), ...(hidden ? { hidden } : {}) },
    frameOf(kid(kid(pic, 'spPr'), 'xfrm')),
    [...(root ? [root] : []), ...(preview ? [preview] : [])],
    { control: 'true' },
    ctx,
    { poster: preview ? await imageDataUrl(preview.part, ctx) : undefined }
  );
}
