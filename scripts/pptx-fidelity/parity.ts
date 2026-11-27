/**
 * The automated text parity check (plan K9): every text run of the source slides must appear in the imported model, as a
 * multiset compare. The source side is read independently of the reader, by tokenising the slide XML for `a:t` runs, so a
 * reader bug cannot hide itself. The check covers the shapes the reader reads so far (`p:sp` directly under the shape tree);
 * runs in shape kinds a later reader step owns (groups, pictures, tables and other graphic frames, connectors) and the
 * `mc:Fallback` copies of alternate content are counted as `outOfScope`, so the number shows what the check does not yet cover. An `a:t` outside every container the tokeniser
 * knows (outside any shape at all, or in an element no reader step owns) is in scope and so reported as `missing`: such text
 * would be silently dropped by the reader, which is what the check exists to show.
 */
import type { OoxmlParts } from '../fidelity-core/ooxml-package';
import type { Json } from '../fidelity-core/types';

export interface TextParity {
  /** Runs in the source slides that the model must hold. */
  sourceRuns: number;
  /** Runs the model holds (text and field runs). */
  modelRuns: number;
  /** Runs in shape kinds the reader does not read yet. */
  outOfScope: number;
  /** Runs in slides the model reports as an error placeholder (damaged part: reported with a warning, so not silent). */
  inErrorSlides: number;
  /** Source runs the model lacks, at most `MAX_LISTED`. */
  missing: string[];
  /** Model runs the source lacks (an invented or duplicated run), at most `MAX_LISTED`. */
  extra: string[];
}

const MAX_LISTED = 20;
const SLIDE = /^ppt\/slides\/slide\d+\.xml$/;
/** Containers whose text a later reader step owns. */
const LATER = new Set(['p:grpSp', 'p:pic', 'p:graphicFrame', 'p:cxnSp', 'mc:Fallback']);
const TAG = /<(\/?)([A-Za-z_][\w.-]*:[\w.-]+)(?:"[^"]*"|'[^']*'|[^>"'])*?(\/?)>|<!--[\s\S]*?-->/g;
const NAMED: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

/** Character data of an `a:t`: CDATA sections as they are, the predefined and numeric references decoded. */
const decode = (s: string): string =>
  s
    .split(/(<!\[CDATA\[[\s\S]*?\]\]>)/)
    .map((part) =>
      part.startsWith('<![CDATA[')
        ? part.slice(9, -3)
        : part.replace(/&(?:#(\d+)|#x([0-9a-fA-F]+)|(lt|gt|amp|quot|apos));/g, (_m, dec?: string, hex?: string, named?: string) =>
            named ? NAMED[named] : String.fromCodePoint(dec !== undefined ? Number(dec) : parseInt(hex as string, 16))
          )
    )
    .join('');

/** The `a:t` texts of one slide part, split by whether a shape kind the reader reads so far holds them. */
export function sourceRuns(xml: string): { inScope: string[]; outOfScope: string[] } {
  const inScope: string[] = [];
  const outOfScope: string[] = [];
  const stack: string[] = [];
  let later = 0;
  // where the character data of the `a:t` just opened starts
  let textFrom = -1;
  for (let m = TAG.exec(xml); m; m = TAG.exec(xml)) {
    if (textFrom >= 0) {
      if (m[1] === '/' && m[2] === 'a:t') (later > 0 ? outOfScope : inScope).push(decode(xml.slice(textFrom, m.index)));
      textFrom = -1;
    }
    const name = m[2];
    if (name === undefined) continue;
    if (m[1] === '/') {
      const at = stack.lastIndexOf(name);
      if (at >= 0) for (const closed of stack.splice(at)) if (LATER.has(closed)) later--;
    } else if (m[3] !== '/') {
      stack.push(name);
      if (LATER.has(name)) later++;
      if (name === 'a:t') textFrom = TAG.lastIndex;
    }
  }
  TAG.lastIndex = 0;
  return { inScope: inScope.filter((t) => t !== ''), outOfScope: outOfScope.filter((t) => t !== '') };
}

/** The text and field runs of every slide of a normalised model. */
export function modelRuns(model: Json): string[] {
  const out: string[] = [];
  const slides = (model as { slides?: unknown }).slides;
  if (!Array.isArray(slides)) return out;
  for (const slide of slides as { elements?: unknown }[]) {
    if (!Array.isArray(slide?.elements)) continue;
    for (const el of slide.elements as { text?: { paragraphs?: { runs?: { kind?: string; text?: string }[] }[] } }[])
      for (const p of el?.text?.paragraphs ?? [])
        for (const r of p.runs ?? []) if ((r.kind === 'text' || r.kind === 'field') && r.text) out.push(r.text);
  }
  return out;
}

function multisetDiff(from: string[], minus: string[]): string[] {
  const left = new Map<string, number>();
  for (const t of minus) left.set(t, (left.get(t) ?? 0) + 1);
  const out: string[] = [];
  for (const t of from) {
    const n = left.get(t) ?? 0;
    if (n > 0) left.set(t, n - 1);
    else out.push(t);
  }
  return out;
}

/** The part names of the slides the model keeps as error placeholders. */
function errorSlides(model: Json): Set<string> {
  const slides = (model as { slides?: unknown }).slides;
  const out = new Set<string>();
  if (Array.isArray(slides))
    for (const s of slides as { partName?: string; error?: string }[]) if (s?.error && s.partName) out.add(s.partName);
  return out;
}

export function textParity(parts: OoxmlParts, model: Json): TextParity {
  const source: string[] = [];
  const failed = errorSlides(model);
  let outOfScope = 0;
  let inErrorSlides = 0;
  for (const name of Object.keys(parts.xml).sort()) {
    if (!SLIDE.test(name)) continue;
    const runs = sourceRuns(parts.xml[name]);
    if (failed.has(name)) {
      inErrorSlides += runs.inScope.length + runs.outOfScope.length;
      continue;
    }
    source.push(...runs.inScope);
    outOfScope += runs.outOfScope.length;
  }
  const held = modelRuns(model);
  return {
    sourceRuns: source.length,
    modelRuns: held.length,
    outOfScope,
    inErrorSlides,
    missing: multisetDiff(source, held).slice(0, MAX_LISTED),
    extra: multisetDiff(held, source).slice(0, MAX_LISTED),
  };
}
