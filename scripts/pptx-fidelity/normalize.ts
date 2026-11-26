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

/** A JSON value for anything: binaries become `{ bytes, sha256 }`, undefined properties are dropped, non-finite numbers become text. */
export function normalizePresentation(value: unknown): Json {
  if (value instanceof Uint8Array) return { bytes: value.length, sha256: sha256(value) };
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
