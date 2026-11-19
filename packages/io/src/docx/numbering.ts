import { attrsOf, childrenOf, find, findAll, parseXml } from './xml-tree';

// ---------------------------------------------------------------------------
// Numbering (word/numbering.xml): numId -> abstractNumId -> (per level) format and start.
// A list item only tells you its numId + ilvl; the bullet-vs-numbered distinction, the first number and the
// restart overrides live in this separate part.
// ---------------------------------------------------------------------------

export interface NumberingInfo {
  numIdToAbstract: Map<string, string>;
  abstractLevelFormat: Map<string, Map<number, string>>; // abstractNumId -> ilvl -> w:numFmt value
  /** abstractNumId -> ilvl -> `w:start` (default 1). */
  abstractLevelStart: Map<string, Map<number, number>>;
  /** numId -> ilvl -> `w:lvlOverride/w:startOverride`. */
  startOverrides: Map<string, Map<number, number>>;
}

export function parseNumbering(xml: string | undefined): NumberingInfo {
  const info: NumberingInfo = {
    numIdToAbstract: new Map(),
    abstractLevelFormat: new Map(),
    abstractLevelStart: new Map(),
    startOverrides: new Map(),
  };
  if (!xml) return info;
  const root = parseXml(xml);
  const numberingEl = find(root, 'w:numbering');
  if (!numberingEl) return info;

  for (const abstractNum of findAll(childrenOf(numberingEl), 'w:abstractNum')) {
    const abstractId = attrsOf(abstractNum)['@_w:abstractNumId'];
    if (!abstractId) continue;
    const levels = new Map<number, string>();
    const starts = new Map<number, number>();
    for (const lvl of findAll(childrenOf(abstractNum), 'w:lvl')) {
      const ilvl = Number(attrsOf(lvl)['@_w:ilvl'] ?? '0');
      const numFmtEl = find(childrenOf(lvl), 'w:numFmt');
      const fmt = numFmtEl ? attrsOf(numFmtEl)['@_w:val'] : undefined;
      if (fmt) levels.set(ilvl, fmt);
      const startEl = find(childrenOf(lvl), 'w:start');
      const start = startEl ? Number(attrsOf(startEl)['@_w:val']) : NaN;
      if (Number.isFinite(start)) starts.set(ilvl, start);
    }
    info.abstractLevelFormat.set(abstractId, levels);
    info.abstractLevelStart.set(abstractId, starts);
  }
  for (const num of findAll(childrenOf(numberingEl), 'w:num')) {
    const numId = attrsOf(num)['@_w:numId'];
    const abstractEl = find(childrenOf(num), 'w:abstractNumId');
    const abstractId = abstractEl ? attrsOf(abstractEl)['@_w:val'] : undefined;
    if (numId && abstractId) info.numIdToAbstract.set(numId, abstractId);
    const overrides = new Map<number, number>();
    for (const o of findAll(childrenOf(num), 'w:lvlOverride')) {
      const ilvl = Number(attrsOf(o)['@_w:ilvl'] ?? '0');
      const startEl = find(childrenOf(o), 'w:startOverride');
      const start = startEl ? Number(attrsOf(startEl)['@_w:val']) : NaN;
      if (Number.isFinite(start)) overrides.set(ilvl, start);
    }
    if (numId && overrides.size > 0) info.startOverrides.set(numId, overrides);
  }
  return info;
}

export function listTypeFor(numbering: NumberingInfo, numId: string, ilvl: number): 'bullet' | 'numbered' {
  const abstractId = numbering.numIdToAbstract.get(numId);
  const fmt = abstractId ? numbering.abstractLevelFormat.get(abstractId)?.get(ilvl) : undefined;
  return fmt === 'bullet' ? 'bullet' : 'numbered';
}

/**
 * Follows list numbering in document order, the way Word does: lists that share an abstract numbering share
 * their counters (so a list interrupted by a paragraph carries on counting), a deeper level restarts after a
 * shallower item, and a `startOverride` restarts a level the first time its numId is used there.
 */
export class ListCounter {
  private readonly counters = new Map<string, Array<number | undefined>>();
  private readonly used = new Set<string>();

  constructor(private readonly numbering: NumberingInfo) {}

  /** The number the item shows, or undefined when the numId is unknown. */
  next(numId: string, ilvl: number): number | undefined {
    const abstractId = this.numbering.numIdToAbstract.get(numId);
    if (!abstractId) return undefined;
    const counters = this.counters.get(abstractId) ?? [];
    this.counters.set(abstractId, counters);

    const key = `${numId}:${ilvl}`;
    if (!this.used.has(key)) {
      this.used.add(key);
      const override = this.numbering.startOverrides.get(numId)?.get(ilvl);
      if (override !== undefined) counters[ilvl] = override - 1;
    }
    const start = this.numbering.abstractLevelStart.get(abstractId)?.get(ilvl) ?? 1;
    const value = (counters[ilvl] ?? start - 1) + 1;
    counters[ilvl] = value;
    for (let l = ilvl + 1; l < counters.length; l++) counters[l] = undefined;
    return value;
  }
}
