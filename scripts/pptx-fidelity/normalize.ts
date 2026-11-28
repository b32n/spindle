/**
 * Turns an imported presentation into a stable JSON model for the goldens. A stub until the importer exists (plan K9):
 * it already does what every golden needs (binary payloads become digests, maps and sets become arrays, dates become ISO
 * strings, nothing time- or machine-dependent is kept) and counts slides and elements, so the round-trip diff and the
 * report have something to say the day a model arrives. Slide-specific normalisation (element ids, pool keys) comes with it.
 */
import { createHash } from 'node:crypto';
import { isObj } from '../fidelity-core/diff';
import type { Json, JsonObject, ModelCounts } from '../fidelity-core/types';

const sha256 = (data: Uint8Array | string): string =>
  createHash('sha256').update(data).digest('hex');

/** An embedded image (a picture's `dataUrl`, plan K10): the goldens keep its type, size and digest, not the pixels. */
const DATA_URL = /^data:([^;,]*);base64,/;

/** A JSON value for anything: binaries and data URLs become `{ bytes, sha256 }`, undefined properties are dropped, non-finite numbers become text. */
export function normalizePresentation(value: unknown): Json {
  if (value instanceof Uint8Array) return { bytes: value.length, sha256: sha256(value) };
  if (typeof value === 'string') {
    const m = DATA_URL.exec(value);
    if (m) {
      const bytes = Buffer.from(value.slice(m[0].length), 'base64');
      return { dataUrl: m[1], bytes: bytes.length, sha256: sha256(bytes) };
    }
  }
  if (value instanceof ArrayBuffer) return normalizePresentation(new Uint8Array(value));
  if (value instanceof Date)
    return Number.isNaN(value.getTime()) ? 'Invalid Date' : value.toISOString();
  if (value instanceof Map) return normalizePresentation([...value.entries()]);
  if (value instanceof Set) return normalizePresentation([...value]);
  if (Array.isArray(value)) return value.map(normalizePresentation);
  if (value && typeof value === 'object') {
    const out: JsonObject = {};
    for (const [k, v] of Object.entries(value)) if (v !== undefined) out[k] = normalizePresentation(v);
    return out;
  }
  if (typeof value === 'number' && !Number.isFinite(value)) return `#number:${String(value)}`;
  return (value ?? null) as Json;
}

/** Size counters of a normalised model: slides and elements (every object under a slide's `elements`). */
export function countPresentation(model: Json): ModelCounts {
  const counts: ModelCounts = { slides: 0, elements: 0 };
  if (!isObj(model) || !Array.isArray(model.slides)) return counts;
  counts.slides = model.slides.length;
  for (const slide of model.slides)
    if (isObj(slide) && Array.isArray(slide.elements)) counts.elements += slide.elements.length;
  return counts;
}

/** Folds array indexes out of a diff path so `byField` stays readable. */
export const diffField = (path: string): string => path.replace(/\[\d+\]/g, '[]');

/** Replace absolute repository paths in an error message so goldens do not depend on the checkout location. */
export function scrubMessage(message: string, roots: string[]): string {
  return roots
    .filter(Boolean)
    .sort((a, b) => b.length - a.length)
    .reduce((m, root) => m.split(root).join('<repo>'), message);
}

/** Positions and sizes of a shape's frame may differ by this many px between the two imports (the writer rounds to whole EMU, 1 px = 9525 EMU). */
export const POSITION_TOLERANCE_PX = 0.5;

/** The tolerance of a diff path: half a pixel for the x, y, width and height of a frame, none elsewhere. */
export const positionTolerance = (path: string): number =>
  /\.frame\.(?:x|y|w|h)$/.test(path) ? POSITION_TOLERANCE_PX : 0;

const isElementList = (path: string): boolean => /(^|\.)(?:slides\[\d+\]\.elements|layouts\[\d+\]\.elements|children)$/.test(path);

/**
 * Aligns the elements of a slide, a layout or a group by kind and name, so an element the writer does not export yet is one
 * difference and not a shift of every element after it. Other arrays are compared by index.
 */
export function elementSignatures(
  path: string,
  before: Json[],
  after: Json[]
): { before: string[]; after: string[] } | null {
  if (!isElementList(path)) return null;
  const sig = (e: Json): string =>
    isObj(e) ? `${String(e.kind)}:${String(e.name)}` : '';
  return { before: before.map(sig), after: after.map(sig) };
}
