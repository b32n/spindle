import type JSZip from 'jszip';
import type { PptxFonts, PptxSection, PptxSlide, PresentationImportResult } from '@b32nio/spindle-shared/export/pptx';
import { OoxmlPackage, type Rel } from '../common/ooxml-package';
import { ImportError, toImportError } from '../common/import-error';
import { openSafeZip, type ImportLimits } from '../common/safe-zip';
import { descendants, isOn, kid, kids, type XmlElement } from '../xlsx/xml';
import { DEFAULT_COLOR_MAP, parseColorMap } from './color';
import { isAvailableFont } from './fonts';
import { readAuthors } from './comments';
import { readLayouts } from './layouts';
import { readSlideContent } from './slides';
import { FALLBACK_THEME, parseTheme } from './theme';
import { emuToPx } from './units';

export interface ImportPptxOptions {
  /** Overrides of the safe-import limits (`DEFAULT_IMPORT_LIMITS`). */
  limits?: Partial<ImportLimits>;
}

const OFFICE_DOCUMENT = 'officeDocument';
const DEFAULT_SLIDE_SIZE = { w: 1280, h: 720 };
/** The part of a content type that says what the main part is. */
const PRESENTATION_TYPES = /(?:presentationml\.(?:presentation|template|slideshow)|ms-powerpoint\.(?:presentation|template|slideshow)\.macroEnabled)\.main\+xml$/;
const PRESENTATION_DEFAULT = 'ppt/presentation.xml';

/**
 * Read a .pptx (or .pptm, .potx, .ppsx). The only error this throws is an `ImportError` (`encrypted`, `legacy-binary`,
 * `not-a-package`, `size-limit` or `corrupt`): the package is opened through `openSafeZip` first, and whatever the reader
 * itself then fails on is reported as `corrupt`. A slide that cannot be read does not fail the import: it stays in the
 * list with an `error` and the result carries a warning.
 */
export async function importPptx(bytes: Uint8Array, options: ImportPptxOptions = {}): Promise<PresentationImportResult> {
  try {
    const zip = await openSafeZip(bytes, { limits: options.limits, target: 'a PowerPoint presentation (.pptx)' });
    return await readPptx(zip);
  } catch (e) {
    throw toImportError(e);
  }
}

/** The main part: the `officeDocument` relationship of the package, else `ppt/presentation.xml`; `undefined` when neither is a presentation. */
async function findPresentation(pkg: OoxmlPackage): Promise<string | undefined> {
  const types = await pkg.xml('[Content_Types].xml');
  const overrides = new Map(kids(types, 'Override').map((o) => [(o.attrs.PartName ?? '').replace(/^\//, ''), o.attrs.ContentType ?? '']));
  const root = (await pkg.rels('')).find((r) => r.type === OFFICE_DOCUMENT && !r.external && pkg.has(r.target));
  if (root) return root.target;
  const fallback = [...overrides].find(([, type]) => PRESENTATION_TYPES.test(type))?.[0];
  if (fallback && pkg.has(fallback)) return fallback;
  return pkg.has(PRESENTATION_DEFAULT) ? PRESENTATION_DEFAULT : undefined;
}

/** PowerPoint's largest slide dimension (51206400 EMU, 56 inches) in px. */
const MAX_SLIDE_PX = 5376;

/** `ppt/presentation.xml` -> its own `p:sldSz` in px, or the default 16:9 with a warning when it is missing, under 1 px or over what PowerPoint allows. */
function readSize(presentation: XmlElement, warn: (m: string) => void): { w: number; h: number } {
  const size = kid(presentation, 'sldSz');
  const w = emuToPx(size?.attrs.cx);
  const h = emuToPx(size?.attrs.cy);
  if (w !== undefined && h !== undefined && w >= 1 && h >= 1 && w <= MAX_SLIDE_PX && h <= MAX_SLIDE_PX) return { w: Math.round(w), h: Math.round(h) };
  warn('The presentation does not give a usable slide size; 16:9 (1280 x 720) is used.');
  return { ...DEFAULT_SLIDE_SIZE };
}

async function readPptx(zip: JSZip): Promise<PresentationImportResult> {
  const pkg = new OoxmlPackage(zip);
  const warnings: string[] = [];
  const warn = (m: string) => {
    if (!warnings.includes(m)) warnings.push(m);
  };

  const main = await findPresentation(pkg);
  const presentation = main === undefined ? undefined : await pkg.xml(main);
  if (main === undefined || !presentation || presentation.name !== 'presentation')
    throw new ImportError('not-a-package', 'This is not a PowerPoint presentation (the package has no presentation part).');

  const rels = new Map((await pkg.rels(main)).map((r) => [r.id, r]));
  const slides = await readSlides(pkg, presentation, rels, warn);
  const first = await firstMaster(pkg, rels, slides);
  const colorMap = first?.colorMap ?? DEFAULT_COLOR_MAP;
  const theme = first?.theme ?? FALLBACK_THEME;
  if (!first) warn('The presentation has no readable slide master and theme; the Office default theme is used.');
  for (const m of first?.warnings ?? []) warn(m);

  const fonts = new Set<string>();
  const defaultStyle = kid(presentation, 'defaultTextStyle');
  const layouts = await readLayouts(pkg, await masterParts(pkg, presentation, rels, slides), defaultStyle, warn, fonts);
  await readSlideContent(pkg, slides, layouts, defaultStyle, warn, fonts, await readAuthors(pkg, [...rels.values()]), [...rels.values()].find((r) => r.type === 'tableStyles' && !r.external)?.target);
  const sections = readSections(presentation, slides);
  for (const part of pkg.skipped) warn(`${part} was too large to read and was skipped.`);
  return {
    slideSize: readSize(presentation, warn),
    theme,
    colorMap,
    slides,
    layouts,
    fonts: reportFonts(fonts, theme, warn),
    ...(sections.length > 0 ? { sections } : {}),
    warnings,
  };
}

/**
 * The fonts the deck names (slides, layouts, masters and the theme's font scheme), one entry per family however it is
 * capitalised (the first spelling wins), and the ones outside the editor's short list (`pptx.fonts.unavailable`). Text in those
 * shows in a substitute and may reflow; nothing is fetched.
 */
function reportFonts(named: Set<string>, theme: PresentationImportResult['theme'], warn: (m: string) => void): PptxFonts {
  const seen = new Map<string, string>();
  for (const f of [...named, ...Object.values(theme.fonts)]) if (!seen.has(f.toLowerCase())) seen.set(f.toLowerCase(), f);
  const used = [...seen.values()].sort((a, b) => a.localeCompare(b));
  const unavailable = used.filter((f) => !isAvailableFont(f));
  if (unavailable.length > 0) warn(`${unavailable.length === 1 ? 'A font' : `${unavailable.length} fonts`} the editor may not have (${unavailable.join(', ')}) ${unavailable.length === 1 ? 'is' : 'are'} used; text in ${unavailable.length === 1 ? 'it' : 'them'} shows in a substitute and may reflow.`);
  return { used, unavailable };
}

/** The slides in `p:sldIdLst` order. A slide whose relationship or part is missing, or whose XML is damaged, stays in the list with an `error`. */
async function readSlides(pkg: OoxmlPackage, presentation: XmlElement, rels: Map<string, Rel>, warn: (m: string) => void): Promise<PptxSlide[]> {
  const slides: PptxSlide[] = [];
  const entries = kids(kid(presentation, 'sldIdLst'), 'sldId');
  // `attrs.id` is `r:id` when `id` is absent, so only a number counts
  const realId = (e: XmlElement): string | undefined => (/^\d+$/.test(e.attrs.id ?? '') ? e.attrs.id : undefined);
  const used = new Set(entries.map(realId));
  let synthetic = 255;
  // a slide without a usable id gets the next number no other slide uses, so ids (sections, K9c timing) stay unique
  const nextId = (): string => {
    do synthetic++;
    while (used.has(String(synthetic)));
    return String(synthetic);
  };
  for (const entry of entries) {
    const number = slides.length + 1;
    const rel = rels.get(entry.rattrs?.id ?? '');
    const pptxId = realId(entry) ?? nextId();
    const failed = (partName: string, error: string): PptxSlide => {
      warn(`Slide ${number} could not be read (${error}); it is kept as an empty placeholder.`);
      return { pptxId, partName, error };
    };
    if (!rel || rel.external) {
      slides.push(failed('', 'its relationship is missing'));
      continue;
    }
    // a slide with an unclosed element would be read as far as the lenient parser gets and lose the rest of its content, so it is a damaged part
    const { root, wellFormed } = await pkg.xmlChecked(rel.target);
    if (!root || root.name !== 'sld' || !wellFormed) {
      slides.push(failed(rel.target, pkg.skipped.includes(rel.target) ? 'the part is too large to read' : pkg.has(rel.target) ? 'the part is damaged' : 'the part is missing'));
      continue;
    }
    const layout = (await pkg.rels(rel.target)).find((r) => r.type === 'slideLayout' && !r.external && pkg.has(r.target));
    const name = (kid(root, 'cSld')?.attrs.name ?? '').trim();
    slides.push({
      pptxId,
      partName: rel.target,
      ...(name !== '' ? { name } : {}),
      ...(isOn(root.attrs.show) ? {} : { hidden: true }),
      ...(layout ? { layoutPart: layout.target } : {}),
    });
  }
  return slides;
}

/**
 * The master (and its theme) the deck takes its theme and colour map from: the master of the first slide's layout, else the
 * first master of `p:sldMasterIdLst`.
 */
async function firstMaster(
  pkg: OoxmlPackage,
  presentationRels: Map<string, Rel>,
  slides: PptxSlide[]
): Promise<{ theme: ReturnType<typeof parseTheme>; colorMap: ReturnType<typeof parseColorMap>; warnings: string[] } | undefined> {
  let masterPart: string | undefined;
  const layoutPart = slides.find((s) => s.layoutPart)?.layoutPart;
  if (layoutPart) masterPart = (await pkg.rels(layoutPart)).find((r) => r.type === 'slideMaster' && !r.external)?.target;
  if (!masterPart || !pkg.has(masterPart)) {
    masterPart = [...presentationRels.values()].find((r) => r.type === 'slideMaster' && !r.external && pkg.has(r.target))?.target;
  }
  if (!masterPart) return undefined;
  const master = await pkg.xml(masterPart);
  const warnings: string[] = [];
  const themePart = (await pkg.rels(masterPart)).find((r) => r.type === 'theme' && !r.external)?.target;
  const themeRoot = themePart === undefined ? undefined : await pkg.xml(themePart);
  if (!themeRoot) return { theme: FALLBACK_THEME, colorMap: parseColorMap(kid(master, 'clrMap')), warnings: ['The slide master has no readable theme; the Office default theme is used.'] };
  return { theme: parseTheme(themeRoot, (m) => warnings.push(m)), colorMap: parseColorMap(kid(master, 'clrMap')), warnings };
}

/** The master parts: those of `p:sldMasterIdLst` in order, then any master a slide's layout belongs to that the list leaves out. */
async function masterParts(pkg: OoxmlPackage, presentation: XmlElement, rels: Map<string, Rel>, slides: PptxSlide[]): Promise<string[]> {
  const parts: string[] = [];
  const add = (part: string | undefined): void => {
    if (part !== undefined && pkg.has(part) && !parts.includes(part)) parts.push(part);
  };
  for (const e of kids(kid(presentation, 'sldMasterIdLst'), 'sldMasterId')) {
    const rel = rels.get(e.rattrs?.id ?? '');
    if (rel && !rel.external && rel.type === 'slideMaster') add(rel.target);
  }
  for (const layout of new Set(slides.map((s) => s.layoutPart))) {
    if (layout !== undefined) add((await pkg.rels(layout)).find((r) => r.type === 'slideMaster' && !r.external)?.target);
  }
  return parts;
}

/** Sections (`p14:section` in the presentation's extension list): name and the slides in it. Slide ids that match no slide are dropped. */
function readSections(presentation: XmlElement, slides: PptxSlide[]): PptxSection[] {
  const known = new Set(slides.map((s) => s.pptxId));
  const sections: PptxSection[] = [];
  for (const s of descendants(presentation, 'section')) {
    const ids = descendants(s, 'sldId').map((e) => e.attrs.id ?? '').filter((id) => known.has(id));
    sections.push({ name: (s.attrs.name ?? '').trim(), slideIds: ids });
  }
  return sections;
}
