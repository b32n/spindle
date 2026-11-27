/**
 * The automated text parity check (plan K9): every text run of the source slides must appear in the imported model, as a
 * multiset compare. The source side is read independently of the reader, by tokenising the slide XML for `a:t` runs, so a
 * reader bug cannot hide itself. The check covers the shapes the reader reads so far (`p:sp`, `p:pic` and table `p:graphicFrame`s, at any depth of `p:grpSp` groups, table cells included);
 * the `mc:Fallback` copies of alternate content that also has a `mc:Choice` (the reader enters the Choice, else the Fallback; but the Fallback of a Choice that `Requires` the `a14` math extension, whose Choice carries only equations) are counted as `outOfScope`, so the number shows what the check does not yet cover. An `a:t` outside every container the tokeniser
 * knows (outside any shape at all, or in an element no reader step owns) is in scope and so reported as `missing`: such text
 * would be silently dropped by the reader, which is what the check exists to show. Speaker notes (K10c) count too: the runs of the
 * body placeholder of the notes page each (non-failed) slide points at must be the runs of `Slide.notes`; the other text of a notes page (header, footer, date,
 * slide number, other shapes) and a notes page no slide points at are `outOfScope`.
 */
import type { OoxmlParts } from '../fidelity-core/ooxml-package';
import type { Json } from '../fidelity-core/types';

export interface TextParity {
  /** Runs in the source slides that the model must hold. */
  sourceRuns: number;
  /** Runs the model holds (text and field runs). */
  modelRuns: number;
  /** Runs in the `mc:Fallback` copies of alternate content, which the reader does not enter when a `mc:Choice` is there. */
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
const NOTES_REL = /<Relationship\b(?:"[^"]*"|'[^']*'|[^>"'])*?>/g;
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
  interface Frame {
    name: string;
    later: boolean;
    choice: boolean;
    /** Alternate content with an `a14` Choice (equations): the text of that Choice is in scope only when no Fallback stands in for it. */
    math?: { texts: string[]; fallback: boolean };
    /** The Choice of such a pair: where its text is held until the alternate content closes. */
    hold?: string[];
  }
  const stack: Frame[] = [];
  let later = 0;
  // where the character data of the `a:t` just opened starts
  let textFrom = -1;
  for (let m = TAG.exec(xml); m; m = TAG.exec(xml)) {
    if (textFrom >= 0) {
      if (m[1] === '/' && m[2] === 'a:t') {
        const text = decode(xml.slice(textFrom, m.index));
        const hold = stack.find((f) => f.hold)?.hold;
        if (hold) hold.push(text);
        else (later > 0 ? outOfScope : inScope).push(text);
      }
      textFrom = -1;
    }
    const name = m[2];
    if (name === undefined) continue;
    if (m[1] === '/') {
      const at = stack.map((f) => f.name).lastIndexOf(name);
      if (at >= 0)
        for (const closed of stack.splice(at)) {
          if (closed.later) later--;
          if (closed.math) (closed.math.fallback || later > 0 ? outOfScope : inScope).push(...closed.math.texts);
        }
    } else if (m[3] !== '/') {
      const parent = stack[stack.length - 1];
      const inAlternate = parent?.name === 'mc:AlternateContent';
      const isChoice = name === 'mc:Choice' && inAlternate;
      // the reader passes over an `a14` Choice (math) for the Fallback, which carries the plain text
      const math = isChoice && /\bRequires\s*=\s*["'][^"']*\ba14\b/.test(m[0]);
      if (math) parent.math = { texts: [], fallback: false };
      else if (isChoice) parent.choice = true;
      if (name === 'mc:Fallback' && inAlternate && parent.math) parent.math.fallback = true;
      // a Fallback is the copy only when a Choice is there to be read instead
      const isLater = name === 'mc:Fallback' && inAlternate && parent.choice;
      stack.push({ name, later: isLater, choice: false, ...(math ? { hold: (parent.math as { texts: string[] }).texts } : {}) });
      if (isLater) later++;
      if (name === 'a:t') textFrom = TAG.lastIndex;
    }
  }
  // alternate content left open by damaged markup still decides its held text
  for (const f of stack) if (f.math) (f.math.fallback ? outOfScope : inScope).push(...f.math.texts);
  TAG.lastIndex = 0;
  return { inScope: inScope.filter((t) => t !== ''), outOfScope: outOfScope.filter((t) => t !== '') };
}

/** The `a:t` texts of one notes page: those of its body placeholders (what `Slide.notes` holds) and the rest. */
export function notesRuns(xml: string): { inScope: string[]; outOfScope: string[] } {
  const all = sourceRuns(xml);
  const body: string[] = [];
  for (const sp of xml.match(/<p:sp\b[\s\S]*?<\/p:sp>/g) ?? []) if (/<p:ph\b(?:"[^"]*"|'[^']*'|[^>"'])*?\btype\s*=\s*["']body["']/.test(sp)) body.push(...sourceRuns(sp).inScope);
  const left = multisetDiff(all.inScope, body);
  return { inScope: body, outOfScope: [...left, ...all.outOfScope] };
}

/** The notes part a slide's relationships name (`.../notesSlide`), resolved against `ppt/slides/`. */
function notesPartOf(parts: OoxmlParts, slidePart: string): string | undefined {
  const rels = parts.xml[slidePart.replace(/([^/]+)$/, '_rels/$1.rels')];
  for (const tag of rels?.match(NOTES_REL) ?? []) {
    const target = /\bTarget\s*=\s*"([^"]*)"/.exec(tag)?.[1];
    if (/\bType\s*=\s*"[^"]*\/notesSlide"/.test(tag) && target !== undefined && !/TargetMode\s*=\s*"External"/.test(tag))
      return target.startsWith('/') ? target.slice(1) : `ppt/${target.replace(/^(\.\.\/)+/, '')}`;
  }
  return undefined;
}

type TextBody = { paragraphs?: { runs?: { kind?: string; text?: string }[] }[] };

type ParityElement = { text?: TextBody; rows?: { cells?: { text?: TextBody }[] }[]; children?: unknown };

/** The runs of elements: a shape's text, the text of every cell of a table, and what a group holds at any depth. */
function collectRuns(elements: unknown[], out: string[]): void {
  for (const el of elements as ParityElement[]) {
    const bodies = [el?.text, ...(el?.rows ?? []).flatMap((row) => (row.cells ?? []).map((cell) => cell.text))];
    for (const body of bodies)
      for (const p of body?.paragraphs ?? [])
        for (const r of p.runs ?? []) if ((r.kind === 'text' || r.kind === 'field') && r.text) out.push(r.text);
    if (Array.isArray(el?.children)) collectRuns(el.children, out);
  }
}

/** The text and field runs of every slide of a normalised model. */
export function modelRuns(model: Json): string[] {
  const out: string[] = [];
  const slides = (model as { slides?: unknown }).slides;
  if (!Array.isArray(slides)) return out;
  for (const slide of slides as { elements?: unknown; notes?: TextBody }[]) {
    if (Array.isArray(slide?.elements)) collectRuns(slide.elements, out);
    if (slide?.notes) collectRuns([{ text: slide.notes }], out);
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
  /** Per notes part: the slides that point at it and are kept, and those the model keeps as error placeholders. */
  const notesOf = new Map<string, { live: number; failed: number }>();
  for (const name of Object.keys(parts.xml).sort()) {
    if (!SLIDE.test(name)) continue;
    const runs = sourceRuns(parts.xml[name]);
    const notes = notesPartOf(parts, name);
    if (notes !== undefined) {
      const n = notesOf.get(notes) ?? { live: 0, failed: 0 };
      if (failed.has(name)) n.failed++;
      else n.live++;
      notesOf.set(notes, n);
    }
    if (failed.has(name)) {
      inErrorSlides += runs.inScope.length + runs.outOfScope.length;
      continue;
    }
    source.push(...runs.inScope);
    outOfScope += runs.outOfScope.length;
  }
  for (const name of Object.keys(parts.xml).sort()) {
    if (!/^ppt\/notesSlides\/[^/]+\.xml$/.test(name)) continue;
    const runs = notesRuns(parts.xml[name]);
    const users = notesOf.get(name);
    if (!users || users.live === 0) {
      if (users) inErrorSlides += runs.inScope.length + runs.outOfScope.length;
      else outOfScope += runs.inScope.length + runs.outOfScope.length;
      continue;
    }
    // a notes page several slides share is held once by each of them
    for (let i = 0; i < users.live; i++) source.push(...runs.inScope);
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
