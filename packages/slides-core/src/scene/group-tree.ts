// Nested groups (K6). A group has no entity of its own: an element names its innermost group in `groupId`, and every
// member of a group that sits inside another group also carries that outer group in `parentGroupId`. The tree is read
// back from those two fields, so it travels with the elements (per-key in the Y layout) and needs no extra record.

/** The grouping fields of an element. */
export interface GroupRef {
  id: string;
  groupId?: string;
  parentGroupId?: string;
}

/** Group id → its parent group id, read from the members. Members of one group agree; if concurrent edits made them
 *  differ, the member with the smallest element id wins so every peer reads the same tree. */
export function groupParents(els: Iterable<GroupRef>): Map<string, string> {
  const winner = new Map<string, string>();
  const parents = new Map<string, string>();
  for (const el of els) {
    const g = el.groupId;
    const p = el.parentGroupId;
    if (!g || !p || p === g) continue;
    const seen = winner.get(g);
    if (seen === undefined || el.id < seen) {
      winner.set(g, el.id);
      parents.set(g, p);
    }
  }
  return parents;
}

/** `groupId` and its ancestors, innermost first. A cycle (only possible from corrupt data) is cut where it closes. */
export function groupChain(groupId: string | undefined, parents: Map<string, string>): string[] {
  const chain: string[] = [];
  const seen = new Set<string>();
  let g = groupId;
  while (g && !seen.has(g)) {
    seen.add(g);
    chain.push(g);
    g = parents.get(g);
  }
  return chain;
}

/** The outermost group an element belongs to (undefined when ungrouped). */
export function rootGroupOf(el: GroupRef, parents: Map<string, string>): string | undefined {
  const chain = groupChain(el.groupId, parents);
  return chain[chain.length - 1];
}
