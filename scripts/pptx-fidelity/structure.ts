import { relationshipTargets } from '../fidelity-core/ooxml-package';
import type { OoxmlParts } from '../fidelity-core/ooxml-package';

/** Package-structure helpers for the PPTX probes: nesting depth, the relationship graph and the slide size. */

/** The rest of an opening tag that is not self-closing: attributes, then `>` (a `/>` never matches). */
export const OPEN = '(?:[^>/]|/(?!>))*>';

const attr = (tag: string, name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];

/** Deepest nesting of `<p:grpSp>` in one slide part. */
export function maxGroupDepth(xml: string): number {
  let depth = 0;
  let max = 0;
  for (const m of xml.matchAll(/<(\/?)p:grpSp\b[^>]*>/g)) {
    if (m[0].endsWith('/>')) continue;
    depth += m[1] ? -1 : 1;
    max = Math.max(max, depth);
  }
  return max;
}

/** Groups whose child coordinate space (`chExt`) differs from their size (`ext`), so that the children are scaled. */
export function scaledGroups(xml: string): number {
  let scaled = 0;
  for (const m of xml.matchAll(
    new RegExp(`<p:grpSpPr\\b${OPEN}\\s*<a:xfrm\\b${OPEN}([\\s\\S]*?)</a:xfrm>`, 'g')
  )) {
    const ext = /<a:ext\b[^>]*>/.exec(m[1])?.[0];
    const child = /<a:chExt\b[^>]*>/.exec(m[1])?.[0];
    if (ext === undefined || child === undefined) continue;
    if (attr(ext, 'cx') !== attr(child, 'cx') || attr(ext, 'cy') !== attr(child, 'cy')) scaled++;
  }
  return scaled;
}

export type SizeClass = 'a4' | 'portrait' | '4x3' | '16x9' | 'custom';

/** The declared slide size and its class (one class per deck); null when presentation.xml has none. */
export function slideSize(p: OoxmlParts): { cx: number; cy: number; kind: SizeClass } | null {
  const tag = /<p:sldSz\b[^>]*>/.exec(p.xml['ppt/presentation.xml'] ?? '')?.[0];
  if (!tag) return null;
  const cx = Number(attr(tag, 'cx'));
  const cy = Number(attr(tag, 'cy'));
  if (!(cx > 0) || !(cy > 0)) return null;
  const ratio = cx / cy;
  const near = (target: number) => Math.abs(ratio - target) < 0.01;
  let kind: SizeClass = 'custom';
  if (cy > cx) kind = 'portrait';
  // Only the declared type counts as A4: a deck with A4 dimensions and no type attribute is deliberately 'custom'.
  else if (attr(tag, 'type') === 'A4') kind = 'a4';
  else if (near(4 / 3)) kind = '4x3';
  else if (near(16 / 9)) kind = '16x9';
  return { cx, cy, kind };
}

/** Relationships as (source part, target part, type), internal ones only. */
export function relationshipEdges(
  p: OoxmlParts
): Array<{ from: string; to: string; type: string }> {
  const edges: Array<{ from: string; to: string; type: string }> = [];
  for (const [name, xml] of Object.entries(p.xml)) {
    if (!name.endsWith('.rels')) continue;
    const from = name.replace(/_rels\//, '').replace(/\.rels$/, '');
    // relationshipTargets matches the same tags in the same order, so the two lists line up.
    const types = [...xml.matchAll(/<Relationship\b[^>]*>/g)].map((m) => attr(m[0], 'Type') ?? '');
    relationshipTargets(name, xml).forEach((r, i) => {
      if (r.target !== null) edges.push({ from, to: r.target, type: types[i] });
    });
  }
  return edges;
}

/**
 * Cycles in the relationship graph, counted as DFS back edges. Layout to master and notes slide to slide are
 * how every deck points back at its owner, so those two relationship kinds are not counted.
 */
export function relationshipCycles(p: OoxmlParts): number {
  const graph = new Map<string, string[]>();
  for (const e of relationshipEdges(p)) {
    if (e.type.endsWith('/slideMaster')) continue;
    if (e.type.endsWith('/slide') && e.from.startsWith('ppt/notesSlides/')) continue;
    graph.set(e.from, [...(graph.get(e.from) ?? []), e.to]);
  }
  const state = new Map<string, 'open' | 'done'>();
  let back = 0;
  for (const start of graph.keys()) {
    if (state.has(start)) continue;
    state.set(start, 'open');
    const stack: Array<{ node: string; next: number }> = [{ node: start, next: 0 }];
    while (stack.length > 0) {
      const top = stack[stack.length - 1];
      const out = graph.get(top.node) ?? [];
      if (top.next >= out.length) {
        state.set(top.node, 'done');
        stack.pop();
        continue;
      }
      const to = out[top.next++];
      const seen = state.get(to);
      if (seen === 'open') back++;
      else if (seen === undefined) {
        state.set(to, 'open');
        stack.push({ node: to, next: 0 });
      }
    }
  }
  return back;
}

/** Slides that `presentation.xml` lists whose relationship is absent or points at a part that is not in the package. */
export function missingSlideParts(p: OoxmlParts): number {
  const rels = p.xml['ppt/_rels/presentation.xml.rels'];
  const pres = p.xml['ppt/presentation.xml'];
  if (rels === undefined || pres === undefined) return 0;
  const targets = new Map(
    relationshipTargets('ppt/_rels/presentation.xml.rels', rels).map((r) => [r.id, r.target])
  );
  let missing = 0;
  for (const m of pres.matchAll(/<p:sldId\b[^>]*>/g)) {
    const to = targets.get(attr(m[0], 'r:id') ?? '');
    if (to === undefined || to === null || !p.names.includes(to)) missing++;
  }
  return missing;
}
