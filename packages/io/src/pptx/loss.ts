import JSZip from 'jszip';
import { createLossReport, type ImportLossItem, type ImportLossReport } from '@b32nio/spindle-shared';
import type { DetectPptxLossOptions } from '@b32nio/spindle-shared/export/pptx';
import { OoxmlPackage, PACKAGE_LIMITS } from '../common/ooxml-package';
import { isEncryptedOle, isOleFile } from '../xlsx/package-kind';
import { descendants, kid, kids, parseXml, type XmlElement } from '../xlsx/xml';
import { collectFonts, isAvailableFont } from './fonts';
import { COMMENTS, LAYOUT, MASTER, NOTES_SLIDE, SLIDE, baseName, isKnownPart } from './parts';
import { scanSlide } from './slide-scan';
import { Tally, newTallies, type Tallies } from './tally';

export type { DetectPptxLossOptions };
export { PACKAGE_LIMITS };

const MAX_WHERE = 5;

/** The default effect: the reader (`importPptx`) does not read the construct yet. Ids it does read say how, below. */
const NOT_IMPORTED = 'lost: not imported (the PPTX reader does not read this yet)';

/** Office's own default theme fonts: a font scheme with any other family is a theme the importer will have to carry. */
const DEFAULT_THEME_FONTS = new Set(['calibri', 'calibri light']);
/** The colour map every master starts from: background and text roles point at the light and dark theme colours. */
const DEFAULT_CLR_MAP: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
const THEME_EFFECTS = new Set(['outerShdw', 'innerShdw', 'prstShdw', 'glow', 'reflection', 'softEdge', 'blur', 'fillOverlay']);

interface SlideEntry {
  part: string;
  label: string;
}

const hasText = (e: XmlElement | undefined): boolean => !!e && e.text.trim() !== '';

export async function detectPptxLoss(bytes: Uint8Array, options: DetectPptxLossOptions = {}): Promise<ImportLossReport> {
  const items: ImportLossItem[] = [];
  try {
    await detect(bytes, options, items);
  } catch {
    // never throw: report whatever was detected before the failure
  }
  return createLossReport('pptx', items);
}

async function detect(bytes: Uint8Array, options: DetectPptxLossOptions, items: ImportLossItem[]): Promise<void> {
  const full = options.detail === 'full';
  const add = (id: string, label: string, count: number, effect: string, where?: string[], onlyFull = false) => {
    if (count > 0 && (full || !onlyFull))
      items.push({ id, label, count, effect, ...(where && where.length ? { where: where.slice(0, MAX_WHERE) } : {}) });
  };
  /** An id that only the `full` level reports. */
  const addFull = (id: string, label: string, count: number, effect: string, where?: string[]) =>
    add(id, label, count, effect, where, true);
  const addTally = (id: string, label: string, tally: Tally, effect = NOT_IMPORTED, onlyFull = false) =>
    add(id, label, tally.total, effect, tally.where, onlyFull);

  if (isOleFile(bytes)) {
    if (isEncryptedOle(bytes))
      add('pptx.encrypted', 'password-protected presentation (encrypted package)', 1, 'lost: cannot be opened; remove the password and save the file again');
    else add('pptx.legacyBinary', 'legacy binary presentation (.ppt)', 1, 'lost: not supported; save it as .pptx first');
    return;
  }
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    return;
  }
  const pkg = new OoxmlPackage(zip);
  const names = pkg.names;
  if (names.length > PACKAGE_LIMITS.entries) {
    add('pptx.sizeLimit', 'package(s) with too many parts to scan', 1, 'lost: not scanned, so its content is not reported', [`more than ${PACKAGE_LIMITS.entries} parts`]);
    return;
  }

  const presentationText = await pkg.text('ppt/presentation.xml');
  const presentation = (presentationText === undefined ? undefined : parseXml(presentationText)) ?? undefined;
  const t = newTallies();
  const slides = await listSlides(pkg, presentationText ?? '', names);
  const fonts = new Set<string>();
  const countParts = (re: RegExp) => names.filter((n) => re.test(n)).length;

  // ---- slides -------------------------------------------------------------
  const seenNotes = new Set<string>();
  const seenComments = new Set<string>();
  let externalLinks = 0;
  for (const slide of slides) {
    const root = await pkg.xml(slide.part);
    if (!root) continue;
    const rels = await pkg.rels(slide.part);
    scanSlide(root, { label: slide.label, rels: new Map(rels.map((r) => [r.id, r])), t });
    collectFonts(root, fonts);
    for (const rel of rels) {
      if (rel.external && rel.type !== 'image') {
        t.externalLinks.add(urlScheme(rel.raw));
        externalLinks++;
      } else if (rel.type === 'notesSlide') {
        seenNotes.add(rel.target);
        if (await notesHaveText(pkg, rel.target)) t.notes.add(slide.label);
      } else if (rel.type === 'comments') {
        seenComments.add(rel.target);
        t.comments.add(slide.label, await countComments(pkg, rel.target));
      }
    }
  }
  // notes pages and comment parts no slide points at still hold content
  for (const part of names.filter((n) => NOTES_SLIDE.test(n) && !seenNotes.has(n)))
    if (await notesHaveText(pkg, part)) t.notes.add(baseName(part));
  for (const part of names.filter((n) => COMMENTS.test(n) && !seenComments.has(n)))
    t.comments.add(baseName(part), await countComments(pkg, part));

  // ---- masters, layouts, themes -------------------------------------------
  const themes = new Set<string>();
  for (const part of names.filter((n) => MASTER.test(n)).sort()) {
    const root = await pkg.xml(part);
    collectFonts(root, fonts);
    const map = kid(root, 'clrMap')?.attrs ?? {};
    if (Object.entries(DEFAULT_CLR_MAP).some(([k, v]) => map[k] !== undefined && map[k] !== v)) t.colorMap.add(baseName(part));
    for (const rel of await pkg.rels(part)) if (rel.type === 'theme' && !rel.external) themes.add(rel.target);
  }
  for (const part of names.filter((n) => LAYOUT.test(n)).sort()) {
    const root = await pkg.xml(part);
    collectFonts(root, fonts);
    t.colorMap.add(baseName(part), descendants(root, 'overrideClrMapping').length);
  }
  for (const part of [...themes].sort()) {
    const root = await pkg.xml(part);
    collectFonts(root, fonts);
    scanTheme(root, baseName(part), t);
  }

  // ---- presentation-level facts -------------------------------------------
  const sections = descendants(presentation, 'section').length;
  const customShows = descendants(presentation, 'custShow').length;
  const size = kid(presentation, 'sldSz')?.attrs ?? {};
  const cx = Number(size.cx);
  const cy = Number(size.cy);
  const oddSize = cx > 0 && cy > 0 && Math.abs(cx / cy - 16 / 9) > 0.01;

  // ---- summary ids --------------------------------------------------------
  const unavailable = distinctFonts(fonts).filter((f) => !isAvailableFont(f)).sort();
  add('pptx.charts', 'chart(s)', t.charts.total > 0 ? t.charts.total : countParts(/^ppt\/charts\/chart\d*\.xml$/), NOT_IMPORTED, t.charts.where);
  add('pptx.smartArt', 'SmartArt diagram(s)', t.smartArt.total > 0 ? t.smartArt.total : countParts(/^ppt\/diagrams\/data\d*\.xml$/), NOT_IMPORTED, t.smartArt.where);
  addTally('pptx.media', 'audio or video clip(s)', t.media);
  addTally('pptx.embeddedObjects', 'embedded object(s) (OLE) and control(s)', t.embedded);
  addTally('pptx.transitions', 'slide(s) with a transition', t.transitions);
  addTally('pptx.animations', 'animation effect(s)', t.animations);
  addTally('pptx.comments', 'comment(s)', t.comments);
  add('pptx.macros', 'macro project', countParts(/^ppt\/vbaProject\.bin$/), 'lost: macros are never run or kept');
  addTally('pptx.hiddenSlides', 'hidden slide(s)', t.hidden, 'kept: imported as a hidden slide');
  addTally('pptx.notes', 'slide(s) with speaker notes', t.notes);
  add('pptx.fonts.unavailable', 'font(s) the editor may not have', unavailable.length, 'approximated: font names are kept as written, but text in these fonts will show in a substitute unless the host provides them', unavailable);
  addTally('pptx.shapes.unsupportedGeometry', 'shape(s) with a geometry the editor lacks (custom geometry, or a preset outside its library; those are kept under their name and drawn approximately)', t.geometry, 'approximated: a preset outside the library is kept under its name and drawn approximately; custom geometry keeps its frame, fill and text but not its path');
  addTally('pptx.images.unsupportedFormat', 'picture(s) in a format browsers cannot show (EMF, WMF)', t.unsupportedImages);
  add('pptx.sizeLimit', 'package part(s) over the size limit', pkg.skipped.length, 'lost: not scanned, so their content is not reported', pkg.skipped);

  if (!full) return;

  // ---- full detail: masters and theme -------------------------------------
  const masters = countParts(MASTER);
  addFull('pptx.masters.multiple', 'slide master(s) in a deck with more than one', masters > 1 ? masters : 0, NOT_IMPORTED);
  addFull('pptx.masters.handoutMaster', 'handout master(s)', countParts(/^ppt\/handoutMasters\/[^/]+\.xml$/), NOT_IMPORTED);
  addFull('pptx.masters.notesMaster', 'notes master(s)', countParts(/^ppt\/notesMasters\/[^/]+\.xml$/), NOT_IMPORTED);
  addTally('pptx.masters.layoutInheritance', 'placeholder(s) that take their position from the layout (flattened on import)', t.layoutInheritance, 'approximated: flattened, the placeholder keeps the position and size it inherits as its own', true);
  addTally('pptx.theme.colorMap', 'colour map(s) that are not the default', t.colorMap, 'approximated: colours are resolved through the map into theme colours; the map itself is not kept', true);
  addTally('pptx.theme.fontScheme', 'theme font scheme(s) other than Office default', t.fontScheme, 'approximated: font names are kept as written and theme font references are resolved to them', true);
  addTally('pptx.theme.effectScheme', 'theme effect scheme(s) with shadows, glow or reflection', t.effectScheme, NOT_IMPORTED, true);

  // ---- text ---------------------------------------------------------------
  addTally('pptx.text.autofit', 'text box(es) that shrink or resize to fit their text', t.autofit, 'approximated: the stored font scale and line spacing reduction are applied as written; line breaks may differ from PowerPoint', true);
  addTally('pptx.text.bullets.custom', 'paragraph(s) with custom bullets or numbering', t.bulletsCustom, 'approximated: bullet character, colour, size and numbering are kept; bullet fonts and numbering schemes beyond arabic, alpha and roman are not', true);
  addTally('pptx.text.bullets.picture', 'paragraph(s) with picture bullets', t.bulletsPicture, NOT_IMPORTED, true);
  addTally('pptx.text.tabStops', 'paragraph(s) with tab stops', t.tabStops, 'approximated: tab stops are kept, but the editor lays out default tabs', true);
  addTally('pptx.text.columns', 'text box(es) with several columns', t.columns, 'approximated: the column count and spacing are kept; shrink-to-fit measures the balanced height', true);
  addTally('pptx.text.vertical', 'text box(es) or cell(s) with vertical text', t.vertical, 'approximated: the text direction is kept and drawn by the browser', true);
  addTally('pptx.text.fields', 'field(s) (slide number, date)', t.fields, 'approximated: kept as a field with its stored text; the editor shows the stored text', true);
  addTally('pptx.text.wordArt', 'text box(es) with a WordArt warp', t.wordArt, NOT_IMPORTED, true);
  addTally('pptx.text.rtl', 'right-to-left paragraph(s)', t.rtl, 'approximated: the paragraph direction is kept and drawn by the browser', true);
  addTally('pptx.text.spacing', 'run(s) with character spacing or capitalisation', t.spacing, 'approximated: character spacing and capitalisation are kept; small caps are the browser\'s', true);
  addTally('pptx.text.baseline', 'run(s) raised or lowered (superscript, subscript)', t.baseline, 'approximated: kept, drawn at a reduced size shifted by the stored percent (PowerPoint\'s size change differs by font)', true);
  addTally('pptx.text.lang', 'run(s) tagged with a language other than en-US', t.lang, 'kept: the language tag is stored on the run; it does not drive spell checking', true);

  // ---- shapes -------------------------------------------------------------
  addTally('pptx.shapes.custGeom', 'shape(s) with custom geometry', t.custGeom, 'approximated: the shape keeps its frame, fill and text but not its path', true);
  addTally('pptx.shapes.effects', 'shape(s) or run(s) with shadow, glow, reflection or soft edges', t.effects, NOT_IMPORTED, true);
  addTally('pptx.shapes.gradientFill', 'gradient fill(s)', t.gradientFill, 'approximated: imported as a gradient; a gradient taken from the theme\'s style list (not the shape\'s own fill) is read as a solid colour', true);
  addTally('pptx.shapes.patternFill', 'pattern fill(s)', t.patternFill, 'approximated: kept as a pattern but painted as a solid blend of its two colours', true);
  addTally('pptx.shapes.connectors.elbowCurved', 'elbow or curved connector(s)', t.elbowCurved, NOT_IMPORTED, true);
  addTally('pptx.shapes.groups.nested', 'group(s) inside another group', t.nestedGroups, NOT_IMPORTED, true);
  addTally('pptx.shapes.scene3d', '3-D shape(s) or text', t.scene3d, NOT_IMPORTED, true);

  // ---- images and tables --------------------------------------------------
  addTally('pptx.images.crop', 'cropped picture(s)', t.crop, NOT_IMPORTED, true);
  addTally('pptx.images.recolor', 'recoloured picture(s)', t.recolor, NOT_IMPORTED, true);
  addTally('pptx.images.linked', 'linked (not embedded) picture(s)', t.linked, NOT_IMPORTED, true);
  addTally('pptx.images.svg', 'SVG picture(s)', t.svg, NOT_IMPORTED, true);
  addTally('pptx.tables.merge', 'merged table cell(s)', t.merge, NOT_IMPORTED, true);
  addTally('pptx.tables.styles', 'table(s) with a table style', t.tableStyles, NOT_IMPORTED, true);
  addTally('pptx.tables.banding', 'table(s) with banded rows or columns', t.banding, NOT_IMPORTED, true);

  // ---- presentation -------------------------------------------------------
  addFull('pptx.sections', 'section(s)', sections, NOT_IMPORTED);
  addFull('pptx.customShows', 'custom show(s)', customShows, NOT_IMPORTED);
  addFull('pptx.slideSize', 'slide size other than 16:9', oddSize ? 1 : 0, 'kept: the deck keeps its own slide size', oddSize ? [`${Math.round(cx / 12700)} x ${Math.round(cy / 12700)} pt`] : []);
  addTally('pptx.zoom', 'slide or section zoom object(s)', t.zoom, NOT_IMPORTED, true);
  add('pptx.ink', 'ink drawing(s)', t.ink.total > 0 ? t.ink.total : countParts(/^ppt\/ink\/[^/]+\.xml$/), NOT_IMPORTED, t.ink.where, true);
  addTally('pptx.equations', 'equation(s)', t.equations, NOT_IMPORTED, true);
  addFull('pptx.signatures', 'digital signature(s)', countParts(/^_xmlsignatures\/[^/]+\.xml$/), 'lost: signatures are not kept, and any edit would invalidate them anyway');
  addFull('pptx.externalLinks', 'external link target(s) (web pages, files, linked media)', externalLinks, 'lost: not imported; links are never fetched', t.externalLinks.where);
  addTally('pptx.actions', 'action hyperlink(s) (jump to slide, run macro, open file)', t.actions, 'lost: not imported; actions are never run, macros included', true);

  // ---- package ------------------------------------------------------------
  if (presentation) {
    const unaccounted = names.filter((n) => !isKnownPart(n));
    addFull('pptx.package.unaccountedParts', 'package part(s) the importer neither reads nor reports', unaccounted.length, 'lost: dropped without a more specific report', unaccounted);
  }
}

/** The font names, one per family however it is capitalised (the first spelling met wins). */
function distinctFonts(fonts: Set<string>): string[] {
  const seen = new Map<string, string>();
  for (const f of fonts) if (!seen.has(f.toLowerCase())) seen.set(f.toLowerCase(), f);
  return [...seen.values()];
}

/**
 * Slide parts in presentation order (labelled `Slide n`), then any slide part the list does not mention. The relationship
 * id of each `<p:sldId id="256" r:id="rId4"/>` is read from the text: the tree keeps only one of the two `id` attributes.
 */
async function listSlides(pkg: OoxmlPackage, presentationText: string, names: string[]): Promise<SlideEntry[]> {
  const rels = new Map((await pkg.rels('ppt/presentation.xml')).map((r) => [r.id, r]));
  const slides: SlideEntry[] = [];
  const listed = new Set<string>();
  for (const tag of presentationText.match(/<(?:[\w.-]+:)?sldId\b[^>]*>/g) ?? []) {
    const rel = rels.get(/\s[\w.-]+:id\s*=\s*["']([^"']*)["']/.exec(tag)?.[1] ?? '');
    if (rel && !rel.external && pkg.has(rel.target) && !listed.has(rel.target)) {
      listed.add(rel.target);
      slides.push({ part: rel.target, label: `Slide ${slides.length + 1}` });
    }
  }
  for (const part of names.filter((n) => SLIDE.test(n) && !listed.has(n)).sort()) slides.push({ part, label: baseName(part) });
  return slides;
}

const urlScheme = (target: string): string => /^([a-z][a-z0-9+.-]*):/i.exec(target)?.[1].toLowerCase() ?? 'relative';

/** True when a notes page holds text beyond the slide number, date, footer and header placeholders. */
async function notesHaveText(pkg: OoxmlPackage, part: string): Promise<boolean> {
  const root = await pkg.xml(part);
  const SKIP = new Set(['sldNum', 'dt', 'ftr', 'hdr', 'sldImg']);
  return descendants(root, 'sp').some((sp) => {
    const type = descendants(sp, 'ph')[0]?.attrs.type;
    return !(type !== undefined && SKIP.has(type)) && descendants(sp, 't').some(hasText);
  });
}

/** Comment threads in a comments part (legacy `p:cm` or modern `p188:cm`); a legacy reply (a `cm` with a `parentCm`) is not a thread. */
async function countComments(pkg: OoxmlPackage, part: string): Promise<number> {
  return descendants(await pkg.xml(part), 'cm').filter((cm) => descendants(cm, 'parentCm').length === 0).length;
}

function scanTheme(root: XmlElement | undefined, label: string, t: Tallies): void {
  const scheme = kid(root, 'themeElements');
  const font = kid(scheme, 'fontScheme');
  const custom = ['majorFont', 'minorFont']
    .map((k) => (kid(kid(font, k), 'latin')?.attrs.typeface ?? '').trim())
    .filter((f) => f !== '' && !DEFAULT_THEME_FONTS.has(f.toLowerCase()));
  if (custom.length > 0) t.fontScheme.add(custom.join(' / '));
  const styles = kids(kid(kid(scheme, 'fmtScheme'), 'effectStyleLst'), 'effectStyle');
  const real = styles.some((s) => s.children.some((g) => g.name === 'effectDag' || (g.name === 'effectLst' && g.children.some((e) => THEME_EFFECTS.has(e.name)))));
  if (real) t.effectScheme.add(label);
}
