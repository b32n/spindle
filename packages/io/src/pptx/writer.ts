import JSZip from 'jszip';
import { relsPart } from '../common/ooxml-package';
import { baseName } from './parts';
import type { PptxElement, PptxLayout, PptxSlide, PresentationImportResult } from '@b32nio/spindle-shared/export/pptx';
import { BLANK_LAYOUT, layoutXml, MASTER_ID, masterXml } from './writer-layouts';
import { BUILT_IN_TABLE_STYLES } from './table-styles';
import { LOSSY, PartRels, type LossyKey, type WriteContext } from './writer-context';
import { CommentAuthors, commentsXml } from './writer-comments';
import { elementIds, elementsXml } from './writer-elements';
import { notesMasterXml, notesSlideXml } from './writer-notes';
import { allocateElementIds, fillXml } from './writer-shapes';
import { themeXml } from './writer-theme';
import { attrs, emu, esc, NS_CT, NS_P, NS_PKG_REL, relsXml, ROOT_NS, TREE_HEADER, XML_DECLARATION, type RelEntry } from './writer-xml';

// The PPTX writer (plan K12): `jszip` and string builders, no generator library (decision SD3). It takes what `importPptx`
// returns and writes a package: content types, relationships, `presentation.xml` (size, slide list, hidden slides, sections),
// the theme, one master with the deck's layouts, one part per slide and `docProps`. Parts are written in one fixed order with a
// fixed zip timestamp, so the same deck gives the same bytes.
//
// K12a wrote the package, theme, master and layouts; K12b adds the slides' text boxes and shapes (text, fills, outlines, hyperlinks),
// slide backgrounds and the layouts' static shapes; K13 adds pictures (media parts), tables, groups, connectors, custom geometry, notes
// (the notes master and a notes slide per slide) and comments (legacy format). Opaque parts, transitions and timing arrive with K14;
// until then `exportPptx` says in `warnings` what it left out, so nothing is dropped silently.

const CT_ROOT = 'application/vnd.openxmlformats-officedocument.presentationml';
const TYPE = {
  presentation: `${CT_ROOT}.presentation.main+xml`,
  slide: `${CT_ROOT}.slide+xml`,
  layout: `${CT_ROOT}.slideLayout+xml`,
  master: `${CT_ROOT}.slideMaster+xml`,
  presProps: `${CT_ROOT}.presProps+xml`,
  viewProps: `${CT_ROOT}.viewProps+xml`,
  tableStyles: `${CT_ROOT}.tableStyles+xml`,
  notesMaster: `${CT_ROOT}.notesMaster+xml`,
  notesSlide: `${CT_ROOT}.notesSlide+xml`,
  comments: `${CT_ROOT}.comments+xml`,
  commentAuthors: `${CT_ROOT}.commentAuthors+xml`,
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
};
const NS_CORE = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const NS_APP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const SECTION_EXT = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';
/** DOS time of every entry: the zip format cannot go earlier, and a fixed value keeps the bytes the same from one run to the next. */
const ZIP_DATE = new Date(Date.UTC(1980, 0, 1));
const MAX_SLIDE_ID = 2147483647;
const LAYOUT_PART = /^ppt\/slideLayouts\/slideLayout\d+\.xml$/;

export interface PptxExportMetadata {
  title?: string;
  creator?: string;
  description?: string;
  /** Written as the creation and modification time; left out of `docProps/core.xml` when not given, so the package carries no clock. */
  modified?: Date;
}

export interface PptxExportOptions {
  metadata?: PptxExportMetadata;
}

export interface PptxExportResult {
  bytes: Uint8Array;
  /** What the writer left out or changed, as plain sentences. Empty when the whole deck was written. */
  warnings: string[];
}

/** The file extension of each image format a data URL can carry that every PowerPoint version displays; SVG (PowerPoint needs a PNG beside it), WebP and AVIF are left out and reported as a dropped picture. */
const IMAGE_TYPES: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp' };
const MEDIA_PART = /^ppt\/media\/[A-Za-z0-9._-]+\.(?:png|jpe?g|gif|bmp)$/i;
const DATA_URL = /^data:([^;,]+);base64,([A-Za-z0-9+/]*={0,2})$/;

/** The bytes of base64 text, or `undefined` when it is not valid base64 (a length no encoder writes). */
function decodeBase64(text: string): Uint8Array | undefined {
  try {
    return Uint8Array.from(atob(text), (c) => c.charCodeAt(0));
  } catch {
    return undefined;
  }
}

interface Entry {
  name: string;
  data: string | Uint8Array;
  /** Content type of an `Override`; parts with a `Default` (`.rels`, `.xml`) have none. */
  type?: string;
}

/** The slide ids: a slide keeps its `pptxId` when it is a usable `p:sldId` no earlier slide has, else it gets the next unused number from 256. */
function allocateSlideIds(slides: PptxSlide[]): string[] {
  const usable = (id: string): boolean => /^\d+$/.test(id) && Number(id) >= 256 && Number(id) <= MAX_SLIDE_ID;
  const taken = new Set<string>();
  const kept = slides.map((s) => {
    if (!usable(s.pptxId) || taken.has(String(Number(s.pptxId)))) return undefined;
    taken.add(String(Number(s.pptxId)));
    return String(Number(s.pptxId));
  });
  let next = 255;
  return kept.map((id) => {
    if (id !== undefined) return id;
    do next++;
    while (taken.has(String(next)));
    taken.add(String(next));
    return String(next);
  });
}

/** The layout part names: a layout keeps its own (`slideLayoutN.xml`) when that is unique, else the next free number. */
function allocateLayoutParts(layouts: PptxLayout[]): string[] {
  const taken = new Set<string>();
  const kept = layouts.map((l) => {
    if (!LAYOUT_PART.test(l.partName) || taken.has(l.partName)) return undefined;
    taken.add(l.partName);
    return l.partName;
  });
  let n = 0;
  return kept.map((part) => {
    if (part !== undefined) return part;
    let name: string;
    do name = `ppt/slideLayouts/slideLayout${++n}.xml`;
    while (taken.has(name));
    taken.add(name);
    return name;
  });
}

const plural = (n: number, one: string, many = `${one}s`): string => `${n} ${n === 1 ? one : many}`;

/** What this writer does not export yet, one sentence per kind, so the report never says "clean" about a deck it emptied. */
function notExported(p: PresentationImportResult): string[] {
  const counts: Record<string, number> = {};
  const bump = (k: string, n = 1) => {
    counts[k] = (counts[k] ?? 0) + n;
  };
  const count = (list: PptxElement[]) => {
    for (const e of list) {
      if (e.kind === 'opaque') bump('opaque');
      else if (e.kind === 'group') count(e.children);
    }
  };
  for (const s of p.slides) {
    count(s.elements ?? []);
    if (s.opaque?.transition) bump('transition');
    if (s.opaque?.timing) bump('timing');
  }
  const failed = p.slides.filter((s) => s.error !== undefined).length;
  const lines: [string, string][] = [
    ['opaque', 'charts, diagrams, objects, media and ink (K14)'],
    ['transition', 'transitions (K14)'],
    ['timing', 'animation timelines (K14)'],
  ];
  const out = lines.filter(([k]) => counts[k]).map(([k, label]) => `Not exported yet: ${counts[k]} ${label}.`);
  if (failed > 0) out.unshift(`${plural(failed, 'slide')} could not be read at import and ${failed === 1 ? 'is' : 'are'} written empty.`);
  const masters = new Set(p.layouts.map((l) => l.masterPart)).size;
  if (masters > 1) out.unshift(`${masters} slide masters were merged into one (their layouts keep their placeholders).`);
  return out;
}

function presentationXml(p: PresentationImportResult, slideIds: string[], slideRels: string[], masterRel: string, notesMasterRel: string | undefined): string {
  const ids = slideIds.map((id, i) => `<p:sldId id="${id}" r:id="${slideRels[i]}"/>`).join('');
  // a repeated source id (malformed deck) keeps its first slide
  const byOld = new Map<string, string>();
  p.slides.forEach((s, i) => byOld.has(s.pptxId) || byOld.set(s.pptxId, slideIds[i]));
  const sections = (p.sections ?? []).map((s, i) => {
    const members = s.slideIds.flatMap((id) => (byOld.has(id) ? [`<p14:sldId id="${byOld.get(id)}"/>`] : []));
    return `<p14:section${attrs({ name: s.name, id: `{00000000-0000-4000-8000-${String(i + 1).padStart(12, '0')}}` })}><p14:sldIdLst>${members.join('')}</p14:sldIdLst></p14:section>`;
  });
  const ext = sections.length === 0 ? '' : `<p:extLst><p:ext uri="${SECTION_EXT}"><p14:sectionLst xmlns:p14="${NS_P14}">${sections.join('')}</p14:sectionLst></p:ext></p:extLst>`;
  return (
    `${XML_DECLARATION}<p:presentation ${ROOT_NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="${MASTER_ID}" r:id="${masterRel}"/></p:sldMasterIdLst>` +
    `${notesMasterRel !== undefined ? `<p:notesMasterIdLst><p:notesMasterId r:id="${notesMasterRel}"/></p:notesMasterIdLst>` : ''}<p:sldIdLst>${ids}</p:sldIdLst><p:sldSz cx="${emu(p.slideSize.w)}" cy="${emu(p.slideSize.h)}"/><p:notesSz cx="6858000" cy="9144000"/>${ext}</p:presentation>`
  );
}

function slideXml(slide: PptxSlide, ctx: WriteContext): string {
  const ids = allocateElementIds(slide.elements ?? []);
  const bg = slide.background ? fillXml(slide.background, ctx) : '';
  const background = bg === '' ? '' : `<p:bg><p:bgPr>${bg}<a:effectLst/></p:bgPr></p:bg>`;
  return (
    `${XML_DECLARATION}<p:sld ${ROOT_NS}${attrs({ show: slide.hidden ? 0 : undefined })}><p:cSld${attrs({ name: slide.name })}>${background}<p:spTree>${TREE_HEADER}${elementsXml(slide.elements ?? [], ctx, elementIds(ids, ctx))}</p:spTree></p:cSld>` +
    '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
  );
}

function corePropertiesXml(m: PptxExportMetadata | undefined): string {
  const time = m?.modified ? m.modified.toISOString().replace(/\.\d{3}Z$/, 'Z') : undefined;
  const stamped = (tag: string) => (time ? `<dcterms:${tag} xsi:type="dcterms:W3CDTF">${time}</dcterms:${tag}>` : '');
  return (
    `${XML_DECLARATION}<cp:coreProperties xmlns:cp="${NS_CORE}" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `${m?.title ? `<dc:title>${esc(m.title)}</dc:title>` : ''}${m?.creator ? `<dc:creator>${esc(m.creator)}</dc:creator>` : ''}${m?.description ? `<dc:description>${esc(m.description)}</dc:description>` : ''}` +
    `${stamped('created')}${stamped('modified')}</cp:coreProperties>`
  );
}

/**
 * Write a presentation as a .pptx. Slide `i` is `ppt/slides/slide<i+1>.xml`; the layouts keep their part names where those are the usual
 * `slideLayoutN.xml`. A slide whose layout is not in the deck uses the first one (a blank layout is written when the deck has none).
 */
export async function exportPptx(presentation: PresentationImportResult, options: PptxExportOptions = {}): Promise<PptxExportResult> {
  const warnings = notExported(presentation);
  const layouts = presentation.layouts.length > 0 ? presentation.layouts : [BLANK_LAYOUT];
  const layoutParts = allocateLayoutParts(layouts);
  const slideIds = allocateSlideIds(presentation.slides);
  const entries: Entry[] = [];
  const lossy = new Map<LossyKey, number>();
  // images by data URL (one media part per distinct image, named in order of first use) and the table styles the slides name
  const media = new Map<string, { name: string; bytes: Uint8Array } | null>();
  const mediaNames = new Set<string>();
  const mediaTypes = new Map<string, string>();
  const tableStyles = new Map<string, string | undefined>();
  const registerImage = (dataUrl: string, preferred?: string): string | undefined => {
    if (!media.has(dataUrl)) {
      const m = DATA_URL.exec(dataUrl);
      const ext = m ? IMAGE_TYPES[m[1]] : undefined;
      const bytes = m && ext ? decodeBase64(m[2]) : undefined;
      if (!m || !ext || !bytes) media.set(dataUrl, null);
      else {
        // a picture keeps the part name it was read from when that is a media file of the right type that no other has (part names do not differ by case)
        const partExt = preferred?.slice(preferred.lastIndexOf('.') + 1).toLowerCase().replace(/^jpg$/, 'jpeg');
        const kept = preferred !== undefined && MEDIA_PART.test(preferred) && partExt === ext && !mediaNames.has(preferred.toLowerCase()) ? preferred : undefined;
        let name = kept;
        for (let n = 1; name === undefined; n++) name = mediaNames.has(`ppt/media/image${n}.${ext}`.toLowerCase()) ? undefined : `ppt/media/image${n}.${ext}`;
        mediaNames.add(name.toLowerCase());
        mediaTypes.set(name.slice(name.lastIndexOf('.') + 1).toLowerCase(), m[1]);
        media.set(dataUrl, { name, bytes });
      }
    }
    const image = media.get(dataUrl);
    return image ? `../media/${baseName(image.name)}` : undefined;
  };
  const context = (rels: PartRels, slideTarget: (part: string) => string | undefined): WriteContext => {
    let field = 0;
    return {
      rels,
      slideTarget,
      lossy: (key) => lossy.set(key, (lossy.get(key) ?? 0) + 1),
      image: registerImage,
      tableStyle: (id, name) => {
        if (tableStyles.get(id) === undefined) tableStyles.set(id, name);
      },
      nextFieldId: () => `{F1E1D000-0000-4000-8000-${String(++field).padStart(12, '0')}}`,
    };
  };
  let notesCount = 0;
  let commentCount = 0;
  const slideParts = new Map<string, number>();
  presentation.slides.forEach((slide, i) => slideParts.has(slide.partName) || slideParts.set(slide.partName, i + 1));
  const slideTarget = (prefix: string) => (part: string): string | undefined => (slideParts.has(part) ? `${prefix}slide${slideParts.get(part)}.xml` : undefined);
  const add = (name: string, data: string, type?: string) => entries.push({ name, data, type });

  const presentationRels: RelEntry[] = [{ id: 'rId1', type: 'slideMaster', target: 'slideMasters/slideMaster1.xml' }];
  const slideRels = presentation.slides.map((_, i) => `rId${i + 2}`);
  presentation.slides.forEach((_, i) => presentationRels.push({ id: slideRels[i], type: 'slide', target: `slides/slide${i + 1}.xml` }));
  let n = presentation.slides.length + 2;
  const extra: [string, string][] = [['theme', 'theme/theme1.xml'], ['presProps', 'presProps.xml'], ['viewProps', 'viewProps.xml'], ['tableStyles', 'tableStyles.xml']];
  for (const [type, target] of extra) presentationRels.push({ id: `rId${n++}`, type, target });
  const hasNotes = presentation.slides.some((s) => s.notes);
  const notesMasterRel = hasNotes ? `rId${n++}` : undefined;
  if (notesMasterRel) presentationRels.push({ id: notesMasterRel, type: 'notesMaster', target: 'notesMasters/notesMaster1.xml' });
  const authors = new CommentAuthors();
  const authorsRel = presentation.slides.some((s) => s.threads?.length) ? `rId${n++}` : undefined;
  if (authorsRel) presentationRels.push({ id: authorsRel, type: 'commentAuthors', target: 'commentAuthors.xml' });

  add('_rels/.rels', relsXml([
    { id: 'rId1', type: 'officeDocument', target: 'ppt/presentation.xml' },
    { id: 'rId2', type: 'extended-properties', target: 'docProps/app.xml' },
    { id: 'rId3', type: 'core-properties', target: 'docProps/core.xml', typeUri: `${NS_PKG_REL}/metadata/core-properties` },
  ]));
  add('ppt/presentation.xml', presentationXml(presentation, slideIds, slideRels, 'rId1', notesMasterRel), TYPE.presentation);
  add('ppt/_rels/presentation.xml.rels', relsXml(presentationRels));
  add('ppt/presProps.xml', `${XML_DECLARATION}<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="${NS_P}"/>`, TYPE.presProps);
  add('ppt/viewProps.xml', `${XML_DECLARATION}<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="${NS_P}"/>`, TYPE.viewProps);
  // the table styles the slides name are listed once all slides are written (their parts are not kept: each cell states its own formatting)
  const tableStylesEntry: Entry = { name: 'ppt/tableStyles.xml', data: '', type: TYPE.tableStyles };
  entries.push(tableStylesEntry);
  add('ppt/theme/theme1.xml', themeXml(presentation.theme), TYPE.theme);

  const layoutRelIds = layouts.map((_, i) => `rId${i + 1}`);
  add('ppt/slideMasters/slideMaster1.xml', masterXml(presentation.slideSize, presentation.colorMap, layoutRelIds, layouts.find((l) => l.masterName)?.masterName), TYPE.master);
  add('ppt/slideMasters/_rels/slideMaster1.xml.rels', relsXml([
    ...layouts.map((_, i): RelEntry => ({ id: layoutRelIds[i], type: 'slideLayout', target: `../slideLayouts/${baseName(layoutParts[i])}` })),
    { id: `rId${layouts.length + 1}`, type: 'theme', target: '../theme/theme1.xml' },
  ]));
  layouts.forEach((layout, i) => {
    const rels = new PartRels({ id: 'rId1', type: 'slideMaster', target: '../slideMasters/slideMaster1.xml' });
    add(layoutParts[i], layoutXml(layout, context(rels, slideTarget('../slides/'))), TYPE.layout);
    add(relsPart(layoutParts[i]), relsXml(rels.entries));
  });

  // a repeated layout part name (malformed deck) keeps its first layout
  const byPart = new Map<string, string>();
  presentation.layouts.forEach((l, i) => byPart.has(l.partName) || byPart.set(l.partName, layoutParts[i]));
  presentation.slides.forEach((slide, i) => {
    const layout = (slide.layoutPart !== undefined ? byPart.get(slide.layoutPart) : undefined) ?? layoutParts[0];
    const rels = new PartRels({ id: 'rId1', type: 'slideLayout', target: `../slideLayouts/${baseName(layout)}` });
    const ctx = context(rels, slideTarget(''));
    add(`ppt/slides/slide${i + 1}.xml`, slideXml(slide, ctx), TYPE.slide);
    if (slide.notes) {
      const part = `ppt/notesSlides/notesSlide${++notesCount}.xml`;
      const notesRels = new PartRels({ id: 'rId1', type: 'notesMaster', target: '../notesMasters/notesMaster1.xml' });
      notesRels.add('slide', `../slides/slide${i + 1}.xml`);
      add(part, notesSlideXml(slide.notes, context(notesRels, slideTarget('../slides/'))), TYPE.notesSlide);
      add(relsPart(part), relsXml(notesRels.entries));
      rels.add('notesSlide', `../notesSlides/${baseName(part)}`);
    }
    if (slide.threads?.length) {
      const part = `ppt/comments/comment${++commentCount}.xml`;
      for (const t of slide.threads) if (t.format === 'modern') ctx.lossy('modernComment');
      add(part, commentsXml(slide.threads, authors), TYPE.comments);
      rels.add('comments', `../comments/${baseName(part)}`);
    }
    add(`ppt/slides/_rels/slide${i + 1}.xml.rels`, relsXml(rels.entries));
  });
  if (hasNotes) {
    add('ppt/notesMasters/notesMaster1.xml', notesMasterXml(presentation.colorMap), TYPE.notesMaster);
    add('ppt/notesMasters/_rels/notesMaster1.xml.rels', relsXml([{ id: 'rId1', type: 'theme', target: '../theme/theme2.xml' }]));
    add('ppt/theme/theme2.xml', themeXml(presentation.theme), TYPE.theme);
  }
  if (authorsRel) add('ppt/commentAuthors.xml', authors.xml(), TYPE.commentAuthors);
  const defs = [...tableStyles].filter(([id]) => !BUILT_IN_TABLE_STYLES.has(id.toUpperCase())).map(([id, name]) => `<a:tblStyle${attrs({ styleId: id, styleName: name ?? '' })}/>`).join('');
  tableStylesEntry.data = `${XML_DECLARATION}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">${defs}</a:tblStyleLst>`;
  for (const image of media.values()) if (image) entries.push({ name: image.name, data: image.bytes });

  add('docProps/core.xml', corePropertiesXml(options.metadata), TYPE.core);
  add('docProps/app.xml', `${XML_DECLARATION}<Properties xmlns="${NS_APP}"><Application>Spindle</Application><Slides>${presentation.slides.length}</Slides></Properties>`, TYPE.app);

  const contentTypes =
    `${XML_DECLARATION}<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
    [...mediaTypes].map(([ext, type]) => `<Default${attrs({ Extension: ext, ContentType: type })}/>`).join('') +
    entries.filter((e) => e.type).map((e) => `<Override${attrs({ PartName: `/${e.name}`, ContentType: e.type })}/>`).join('') +
    '</Types>';

  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypes, { date: ZIP_DATE });
  for (const e of entries) zip.file(e.name, e.data, { date: ZIP_DATE });
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  for (const [key, n] of [...lossy].sort(([a], [b]) => a.localeCompare(b))) warnings.push(`Changed on export: ${LOSSY[key]} (${n}).`);
  return { bytes, warnings };
}
