import type { PptxColorMap, PptxLevelStyles, PptxTextBody } from '@b32nio/spindle-shared/export/pptx';
import type { OoxmlPackage } from '../common/ooxml-package';
import { kid, type XmlElement } from '../xlsx/xml';
import { DEFAULT_COLOR_MAP, overrideColorMap, parseColorMap } from './color';
import { mergeLevels, parseLevels } from './levels';
import { treeChildren } from './shapes';
import { readTextBody } from './text';

// Speaker notes (plan K10c): the body placeholder of a slide's notes page, as a text body. Its text styles come from the
// notes master (`p:notesStyle`, then the master's own body placeholder), the way a slide's placeholder takes the layout's.
// The rest of the notes page (slide image, header, footer, date, slide number, other shapes) is not kept.

const NO_LEVELS: PptxLevelStyles = Array.from({ length: 9 }, () => ({}));

/** The body placeholder of a notes tree: `p:ph type="body"`, in either its `p:sp` or alternate-content form. */
const isBody = (sp: XmlElement): boolean => kid(kid(kid(sp, 'nvSpPr'), 'nvPr'), 'ph')?.attrs.type === 'body';

/** What a paragraph of the notes body starts from: the notes master's text style under its body placeholder's own list style. */
export function notesBaseLevels(master: XmlElement | undefined, map: PptxColorMap): PptxLevelStyles {
  if (!master) return NO_LEVELS;
  const bodies = treeChildren(kid(kid(master, 'cSld'), 'spTree')).filter((e) => e.name === 'sp' && isBody(e));
  const own = parseLevels(kid(kid(bodies[0], 'txBody'), 'lstStyle'), map);
  return mergeLevels(parseLevels(kid(master, 'notesStyle'), map), own);
}

/** True when a text body holds a run with text (a notes page of empty paragraphs has no notes). */
const hasText = (body: PptxTextBody): boolean => body.paragraphs.some((p) => p.runs.some((r) => r.kind !== 'break' && r.text.trim() !== ''));

/**
 * The notes on the notes page `notesPart`: the paragraphs of every body placeholder of its notes page.
 * `undefined` when the slide has no notes page, the page is unreadable or it holds no text.
 */
export async function readNotes(pkg: OoxmlPackage, notesPart: string | undefined, warn: (m: string) => void): Promise<PptxTextBody | undefined> {
  const root = notesPart === undefined ? undefined : await pkg.xml(notesPart);
  if (!root || root.name !== 'notes') return undefined;
  const masterPart = (await pkg.rels(notesPart as string)).find((r) => r.type === 'notesMaster' && !r.external && pkg.has(r.target))?.target;
  const master = masterPart === undefined ? undefined : await pkg.xml(masterPart);
  const map = overrideColorMap(master ? parseColorMap(kid(master, 'clrMap')) : DEFAULT_COLOR_MAP, kid(kid(root, 'clrMapOvr'), 'overrideClrMapping'));
  const base = notesBaseLevels(master, map);
  const rels = new Map((await pkg.rels(notesPart as string)).map((r) => [r.id, r]));
  const bodies = treeChildren(kid(kid(root, 'cSld'), 'spTree'))
    .filter((e) => e.name === 'sp' && isBody(e))
    .map((sp) => kid(sp, 'txBody'))
    .filter((t): t is XmlElement => t !== undefined)
    .map((t) => readTextBody(t, base, undefined, map, undefined, { rels, warn }));
  const [first] = bodies;
  if (!first) return undefined;
  const notes: PptxTextBody = { props: first.props, paragraphs: bodies.flatMap((b) => b.paragraphs) };
  return hasText(notes) ? notes : undefined;
}
