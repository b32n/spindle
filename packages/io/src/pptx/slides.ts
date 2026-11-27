import type { PptxColorMap, PptxFill, PptxLayout, PptxSlide } from '@b32nio/spindle-shared/export/pptx';
import type { OoxmlPackage } from '../common/ooxml-package';
import { kid, type XmlElement } from '../xlsx/xml';
import { DEFAULT_COLOR_MAP, colorElement, overrideColorMap, parseColor, parseColorMap } from './color';
import { parseFill, type Rels } from './fill';
import { collectFonts } from './fonts';
import { parseLevels } from './levels';
import { elementIds, readAuthors, readThreads } from './comments';
import { readElements, type ElementContext } from './elements';
import { readNotes } from './notes';
import { readTableStyles } from './table-styles';

/** A slide's `p:bg`: the fill of `p:bgPr`, or the colour of a `p:bgRef` as a solid fill. */
function readBackground(root: XmlElement, map: PptxColorMap, rels: Rels): PptxFill | undefined {
  const bg = kid(kid(root, 'cSld'), 'bg');
  if (kid(bg, 'bgPr')) return parseFill(kid(bg, 'bgPr'), map, rels);
  const color = parseColor(colorElement(kid(bg, 'bgRef')), map);
  return color ? { kind: 'solid', color } : undefined;
}

const relsOf = async (pkg: OoxmlPackage, part: string) => new Map((await pkg.rels(part)).map((r) => [r.id, r]));

/**
 * Fill in the content of every readable slide (`elements`, `background`) from its part. A slide whose content cannot be
 * read is degraded to an error placeholder with a warning, like a slide whose part is missing; the others are unaffected.
 */
export async function readSlideContent(
  pkg: OoxmlPackage,
  slides: PptxSlide[],
  layouts: PptxLayout[],
  defaultStyle: XmlElement | undefined,
  warn: (m: string) => void,
  fonts: Set<string>,
  authors: Awaited<ReturnType<typeof readAuthors>>,
  tableStylesPart?: string
): Promise<void> {
  const layoutMaps = new Map<string, PptxColorMap>();
  const layoutMap = async (part: string | undefined): Promise<PptxColorMap> => {
    const key = part ?? '';
    const cached = layoutMaps.get(key);
    if (cached) return cached;
    let map = parseColorMap(undefined);
    if (part) {
      const masterPart = (await pkg.rels(part)).find((r) => r.type === 'slideMaster' && !r.external)?.target;
      map = parseColorMap(masterPart ? kid(await pkg.xml(masterPart), 'clrMap') : undefined);
      map = overrideColorMap(map, kid(kid(await pkg.xml(part), 'clrMapOvr'), 'overrideClrMapping'));
    }
    layoutMaps.set(key, map);
    return map;
  };
  const images: ElementContext['images'] = new Map();
  let styleRoot: XmlElement | undefined;
  if (tableStylesPart) styleRoot = await pkg.xml(tableStylesPart);
  // table styles depend on the style part and the default colour map only (style parts carry no blip fills)
  const styles = readTableStyles(styleRoot, DEFAULT_COLOR_MAP, new Map());
  for (const [i, slide] of slides.entries()) {
    if (slide.error) continue;
    try {
      // read (and cached) when the slide list was built, so the part is there
      const root = (await pkg.xml(slide.partName)) as XmlElement;
      collectFonts(root, fonts);
      const map = overrideColorMap(await layoutMap(slide.layoutPart), kid(kid(root, 'clrMapOvr'), 'overrideClrMapping'));
      const rels = await relsOf(pkg, slide.partName);
      const layout = layouts.find((l) => l.partName === slide.layoutPart);
      const ctx = { map, rels, defaultLevels: parseLevels(defaultStyle, map), ...(layout ? { layout } : {}), warn, pkg, images, styles };
      const background = readBackground(root, map, rels);
      if (background) slide.background = background;
      slide.elements = await readElements(kid(kid(root, 'cSld'), 'spTree'), ctx);
      // notes and comments are extras: one that cannot be read is dropped, with a warning, and the slide keeps its content
      try {
        const notes = await readNotes(pkg, [...rels.values()].find((r) => r.type === 'notesSlide' && !r.external)?.target, warn);
        if (notes) slide.notes = notes;
      } catch (e) {
        warn(`The speaker notes of slide ${i + 1} could not be read (${e instanceof Error ? e.message : 'unknown error'}); they are dropped.`);
      }
      try {
        const threads = await readThreads(pkg, slide.partName, slide.pptxId, authors, elementIds(slide.elements), warn);
        if (threads.length > 0) slide.threads = threads;
      } catch (e) {
        warn(`The comments of slide ${i + 1} could not be read (${e instanceof Error ? e.message : 'unknown error'}); they are dropped.`);
      }
    } catch (e) {
      delete slide.elements;
      delete slide.background;
      delete slide.notes;
      delete slide.threads;
      slide.error = 'the content could not be read';
      warn(`Slide ${i + 1} could not be read (${e instanceof Error ? e.message : 'unknown error'}); it is kept as an empty placeholder.`);
    }
  }
}
