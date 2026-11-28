// PPTX import and export for a deck (plan K15): `importFromPptx` maps what `@b32nio/spindle-io/pptx` reads onto the engine,
// `exportToPptx` maps the engine back onto what the writer takes, `inspectPptx` says what a file holds that the editor cannot take.
// `@b32nio/spindle-io` is loaded on first use (like the sheets and docs cores), so a host that never opens or saves a PPTX does not
// ship the reader and writer.

import type { Comment, ImportLossReport } from '@b32nio/spindle-shared';
import type { OpaquePart } from '@b32nio/spindle-shared/opaque';
import type { DetectPptxLossOptions, PptxCommentThread, PptxSlide, PresentationImportResult } from '@b32nio/spindle-shared/export/pptx';
import type { DeckImpl } from '../deck';
import { generateId } from '../utils/id';
import { indexesBetween } from '../scene/fractional-index';
import type { SlideElement } from '../scene/types';
import { richTextToPlainText } from '../text/model';
import { BUILTIN_LAYOUTS } from '../theme/builtin';
import type { LayoutData } from '../theme/types';
import type { DeckData, SlideData, SlidesCommentThread } from '../types';
import {
  elementsFromPptx,
  elementsToPptx,
  newElementTally,
  type ElementTally,
} from './pptx-elements';
import { DEFAULT_COLOR_MAP, layoutFromPptx, layoutToPptx, themeFromPptx, themeToPptx } from './pptx-layouts';
import { textFromPptx, textToPptx } from './pptx-text';
import { fillFromPptx, fillToPptx } from './pptx-values';

const loadIo = () => import('@b32nio/spindle-io/pptx');
type IoPptx = Awaited<ReturnType<typeof loadIo>>;

/** What `importFromPptx` takes: the reader's limits, and the time stamped on comments the file gives none. */
export interface ImportFromPptxOptions {
  /** Overrides of the safe-import limits (package size, entries, uncompressed size, compression ratio). */
  limits?: NonNullable<Parameters<IoPptx['importPptx']>[1]>['limits'];
  /** Overrides of the size caps of the parts kept as they are (charts, SmartArt, objects, media, ink). */
  opaqueLimits?: NonNullable<Parameters<IoPptx['importPptx']>[1]>['opaqueLimits'];
  /** The time of comments whose file states none (default: now). */
  now?: Date;
}

export interface ExportToPptxOptions {
  /** Written to `docProps/core.xml`; `title` defaults to the deck's. The package carries no clock unless `modified` is given. */
  metadata?: { title?: string; creator?: string; description?: string; modified?: Date };
}

/** The file and what the writer left out or changed, as plain sentences (empty when the whole deck was written). */
export interface ExportToPptxResult {
  bytes: Uint8Array;
  warnings: string[];
}

const plural = (n: number, one: string, many = `${one}s`): string => `${n} ${n === 1 ? one : many}`;

/** What the import changed or left out, as plain sentences. */
function elementWarnings(t: ElementTally): string[] {
  const out: string[] = [];
  if (t.hidden > 0) out.push(`${plural(t.hidden, 'hidden shape')} not imported.`);
  if (t.pictureFills > 0) out.push(`${plural(t.pictureFills, 'picture fill')} not imported (the shape keeps no fill).`);
  if (t.unreadablePictures > 0) out.push(`${plural(t.unreadablePictures, 'picture')} could not be shown and ${t.unreadablePictures === 1 ? 'is a labelled placeholder' : 'are labelled placeholders'}.`);
  if (t.customOutlines > 0) out.push(`${plural(t.customOutlines, 'custom outline')} could not be kept; the shape is a rectangle.`);
  if (t.emptyTables > 0) out.push(`${plural(t.emptyTables, 'empty table')} not imported.`);
  if (t.lineBreaks > 0) out.push(`${plural(t.lineBreaks, 'line break')} became paragraph breaks.`);
  if (t.verticalText > 0) out.push(`${plural(t.verticalText, 'text body', 'text bodies')} with a vertical text direction the editor does not have are drawn horizontally.`);
  return out;
}

// ── Import ───────────────────────────────────────────────────────────────────

function errorSlideElements(containerId: string, message: string, size: { w: number; h: number }): SlideElement[] {
  const text = `This slide could not be read: ${message}`;
  return [
    {
      id: generateId(),
      containerId,
      index: indexesBetween(null, null, 1)[0],
      type: 'text',
      x: size.w * 0.1,
      y: size.h * 0.4,
      w: size.w * 0.8,
      h: size.h * 0.2,
      rotation: 0,
      richText: { type: 'doc', content: [{ type: 'paragraph', attrs: { align: 'center' }, content: [{ type: 'text', text }] }] },
      bodyStyle: { vAlign: 'middle', wrap: true, padding: 8 },
    },
  ];
}

function threadsFromPptx(
  slideId: string,
  threads: PptxCommentThread[],
  byPptxId: Map<string, string>,
  now: string
): SlidesCommentThread[] {
  return threads.map((t) => {
    const comments: Comment[] = t.comments.map((c) => ({
      id: generateId(),
      authorId: `pptx:${c.author}`,
      authorName: c.author,
      ...(c.initials ? { authorInitials: c.initials } : {}),
      body: c.body,
      createdAt: c.createdAt ?? now,
    }));
    // A slide-level comment (and one whose shape is gone) has no element to hang on: the model's anchor is an element, so it keeps an
    // empty element id, which the comments panel lists as "no longer attached" and an export writes back as a slide comment.
    const elementId = t.anchor.kind === 'element' ? (byPptxId.get(t.anchor.pptxId) ?? '') : '';
    return {
      id: generateId(),
      anchor: { slideId, elementId },
      status: t.resolved ? 'resolved' : 'open',
      comments,
      createdAt: comments[0]?.createdAt ?? now,
    };
  });
}

function toDeckData(result: PresentationImportResult, deck: DeckImpl, options: ImportFromPptxOptions, warnings: string[]): DeckData {
  const tally = newElementTally();
  const now = (options.now ?? new Date()).toISOString();
  const layouts = result.layouts.map((l) => layoutFromPptx(l, tally));
  const layoutIds = new Set(layouts.map((l) => l.id));
  const threads: SlidesCommentThread[] = [];
  const slides: SlideData[] = result.slides.map((s, i): SlideData => {
    const id = generateId();
    const elements = s.error !== undefined ? errorSlideElements(id, s.error, result.slideSize) : elementsFromPptx(s.elements ?? [], id, tally);
    if (s.error !== undefined) warnings.push(`Slide ${i + 1} could not be read (${s.error}) and is a labelled placeholder.`);
    const byPptxId = new Map(elements.flatMap((e) => (e.pptxId !== undefined ? [[e.pptxId, e.id] as const] : [])));
    if (s.threads?.length) threads.push(...threadsFromPptx(id, s.threads, byPptxId, now));
    const background = fillFromPptx(s.background);
    if (s.background?.kind === 'picture') tally.pictureFills++;
    const notes = s.notes ? textFromPptx(s.notes, tally).richText : undefined;
    return {
      id,
      index: '',
      ...(s.layoutPart !== undefined && layoutIds.has(s.layoutPart) ? { layoutRef: s.layoutPart } : {}),
      ...(background ? { background } : {}),
      ...(notes && richTextToPlainText(notes) !== '' ? { notes } : {}),
      ...(s.hidden ? { hidden: true } : {}),
      ...(s.opaque && (s.opaque.transition || s.opaque.timing) ? { opaque: { ...s.opaque } } : {}),
      elements,
    };
  });
  const keys = indexesBetween(null, null, slides.length);
  slides.forEach((s, i) => {
    s.index = keys[i];
  });
  warnings.push(...elementWarnings(tally));
  if (result.sections?.length) warnings.push(`${plural(result.sections.length, 'slide section')} not kept (the editor has no sections).`);
  return {
    id: deck.id,
    title: deck.getTitle(),
    slideSize: { ...result.slideSize },
    theme: themeFromPptx(result.theme),
    layouts: layouts.length > 0 ? layouts : BUILTIN_LAYOUTS.map((l) => structuredClone(l)),
    slides,
    ...(threads.length > 0 ? { threads } : {}),
    ...(result.opaqueParts.length > 0 ? { opaqueParts: result.opaqueParts } : {}),
  };
}

/**
 * Replace `deck`'s entire contents with what is in the .pptx at `bytes`: the slide size, the theme, the layouts, every slide with its
 * shapes, text, pictures, tables, connectors, groups, notes and comments, and the charts, diagrams, objects and media that are kept
 * as they are (`opaqueParts`, shown as labelled placeholders). The deck's id and title are kept.
 *
 * Throws an `ImportError` (from `@b32nio/spindle-io/pptx`, with a `code`) for a file that cannot or must not be opened: encrypted,
 * legacy binary, not a package, over the limits, or corrupt; the deck is left untouched. A slide that cannot be read does not fail the
 * import: it becomes a labelled placeholder and `warnings` says so.
 *
 * `loss` is what the file holds that the editor cannot take, with the "kept, not editable" group for what is stored for saving back.
 * Throws a plain `Error` when the file has no slides. Not available while collaboration is attached (like `setData`).
 */
export async function importFromPptx(
  bytes: Uint8Array,
  deck: DeckImpl,
  options: ImportFromPptxOptions = {}
): Promise<{ warnings: string[]; loss: ImportLossReport }> {
  const io = await loadIo();
  const { now: _now, ...readerOptions } = options;
  const result = await io.importPptx(bytes, readerOptions);
  // the package is read twice (the reader, then the loss report): the report needs the file as it is, not as the editor took it
  const loss = await io.detectPptxLoss(bytes);
  if (result.slides.length === 0) throw new Error('The file contains no slides.');
  const warnings = [...result.warnings];
  deck.setData(toDeckData(result, deck, options, warnings));
  return { warnings, loss };
}

/**
 * What importing this .pptx would not bring across, and what it keeps without being able to edit (charts, SmartArt, objects, media,
 * transitions and animations). Read-only: nothing is imported or modified. `detail: 'full'` (default `'summary'`, what the import
 * dialog shows) adds the finer ids the fidelity harness records.
 */
export function inspectPptx(bytes: Uint8Array, options: Pick<DetectPptxLossOptions, 'detail'> = {}): Promise<ImportLossReport> {
  return loadIo().then((io) => io.detectPptxLoss(bytes, options));
}

// ── Export ───────────────────────────────────────────────────────────────────

/** The interchange form of a deck: what `importFromPptx` read it from, as far as the model keeps it. */
function toPresentation(deck: DeckImpl): PresentationImportResult {
  const data = deck.getData();
  const layouts = (data.layouts as LayoutData[]).map(layoutToPptx);
  const layoutIds = new Set(layouts.map((l) => l.partName));
  const slideNumberOf = new Map(data.slides.map((s, i) => [s.id, i]));
  const elementsById = new Map(data.slides.flatMap((s) => s.elements.map((e) => [e.id, e] as const)));
  const threadsBySlide = new Map<string, SlidesCommentThread[]>();
  for (const t of data.threads ?? []) {
    if (slideNumberOf.has(t.anchor.slideId)) threadsBySlide.set(t.anchor.slideId, [...(threadsBySlide.get(t.anchor.slideId) ?? []), t]);
  }
  const slides: PptxSlide[] = data.slides.map((s, i) => {
    const { elements, pptxIds } = elementsToPptx(s.elements);
    const threads = (threadsBySlide.get(s.id) ?? []).map((t): PptxCommentThread => {
      const target = elementsById.get(t.anchor.elementId);
      const pptxId = target ? pptxIds.get(target.id) : undefined;
      return {
        id: t.id,
        format: 'legacy',
        anchor: pptxId !== undefined ? { kind: 'element', pptxId } : { kind: 'slide' },
        ...(target ? { position: { x: target.x, y: target.y } } : {}),
        ...(t.status === 'resolved' ? { resolved: true } : {}),
        comments: t.comments.map((c) => ({ author: c.authorName, ...(c.authorInitials ? { initials: c.authorInitials } : {}), body: c.body, createdAt: c.createdAt })),
      };
    });
    const background = fillToPptx(s.background);
    return {
      pptxId: String(256 + i),
      partName: `ppt/slides/slide${i + 1}.xml`,
      ...(s.hidden ? { hidden: true } : {}),
      ...(s.layoutRef !== undefined && layoutIds.has(s.layoutRef) ? { layoutPart: s.layoutRef } : {}),
      ...(background ? { background } : {}),
      elements,
      ...(s.notes && richTextToPlainText(s.notes) !== '' ? { notes: textToPptx(s.notes, undefined) } : {}),
      ...(threads.length > 0 ? { threads } : {}),
      ...(s.opaque ? { opaque: { ...s.opaque } } : {}),
    };
  });
  const parts: OpaquePart[] = data.opaqueParts ?? [];
  return {
    slideSize: { ...data.slideSize },
    theme: themeToPptx(data.theme),
    colorMap: { ...DEFAULT_COLOR_MAP },
    slides,
    layouts,
    fonts: { used: [], unavailable: [] },
    opaqueParts: parts,
    warnings: [],
  };
}

/**
 * Write `deck` as a .pptx. Everything the model keeps is written: slides and their order, the slide size, the theme, the layouts, shapes,
 * text with its formatting, pictures, tables, connectors (glued to their shapes), groups, notes, comments, and the charts, diagrams,
 * objects and media that came in with a file, byte for byte. `warnings` says what the writer left out or changed.
 *
 * Elements of the deck that were never in a file get new shape ids; the ones that were keep theirs, so the transitions and animations
 * kept with a slide still point at their shapes.
 */
export async function exportToPptx(deck: DeckImpl, options: ExportToPptxOptions = {}): Promise<ExportToPptxResult> {
  const io = await loadIo();
  const presentation = toPresentation(deck);
  const metadata = { title: deck.getTitle(), ...options.metadata };
  const written = await io.exportPptx(presentation, { metadata });
  return { bytes: written.bytes, warnings: written.warnings };
}
