import type { XlsxRelationship } from '@b32nio/spindle-shared/export/xlsx';
import { relsPart, resolveTarget } from '../../common/ooxml-package';
import { scanXml } from './sax';

export { relsPart };

/**
 * The part a relationship target names, resolved against the part that owns the relationships ('' for the package's own
 * `_rels/.rels`). Handles what producers write besides Excel's relative form: absolute (`/xl/worksheets/sheet1.xml`, openpyxl),
 * `./` and `..` segments, backslashes, and percent-encoded names. A target that climbs above the root stays at the root. A
 * fragment (`sheet1.xml#a`) is not part of a part name and is dropped.
 */
export function resolveRelationshipTarget(owner: string, target: string): string {
  let t = target.trim().replace(/\\/g, '/');
  const hash = t.indexOf('#');
  if (hash >= 0) t = t.slice(0, hash);
  try {
    t = decodeURIComponent(t);
  } catch {
    // not valid percent-encoding: the text is the name
  }
  return resolveTarget(owner, t);
}

/** A URI scheme of two or more characters (`file:`, `mailto:`, `http:`); one letter is a Windows drive. */
const SCHEME = /^[a-zA-Z][a-zA-Z0-9+.-]+:/;

const lastSegment = (type: string): string => type.slice(type.lastIndexOf('/') + 1);

/**
 * What went wrong in a part read as a list of entries: `damaged` (not well-formed: what precedes the damage was read), `over` (more
 * entries than the part may have: the first ones were kept) or `budget` (the package's relationship budget ran out in this part).
 */
export type ListProblem = 'damaged' | 'over' | 'budget';

export interface ListOptions {
  /** Most entries read from the part (`ImportLimits.maxPartElements`); the scan ends at the first one over it. */
  max?: number;
  /** Relationship entries that may still be read over the whole package (`ImportLimits.maxRelationships`), shared by every part: each entry read draws on it, and the scan ends when it is spent. */
  budget?: { left: number };
  /** Which relationships are kept, by `kind` (the last segment of the type) and id; the others are read (they count) but not kept. All when absent. */
  keep?: (kind: string, id: string) => boolean;
  problem?: (what: ListProblem) => void;
}

/**
 * Relationships of `owner` from the text of its `.rels`, read as events: memory is the relationships kept (those `options.keep`
 * names), not a tree of the part, and the ids read (at most `options.max`). Entries without an `Id` or a `Target` are skipped; a
 * repeated id keeps the first, whatever it is. Damaged markup ends the read there (`problem('damaged')`); an entry over the part's
 * limit ends it too (`problem('over')`), as does the package's budget running out (`problem('budget')`).
 */
export function parseRelationships(xml: string | undefined, owner: string, options: ListOptions = {}): XlsxRelationship[] {
  const out: XlsxRelationship[] = [];
  if (xml === undefined) return out;
  const max = options.max ?? Infinity;
  const seen = new Set<string>();
  // the same few type URIs on every entry: one string each, not one per entry
  const types = new Map<string, string>();
  let depth = 0;
  let over = false;
  let spent = false;
  const wellFormed = scanXml(xml, {
    open(name, attrs) {
      depth++;
      const id = attrs.Id;
      const raw = attrs.Target;
      if (depth !== 2 || name !== 'Relationship' || id === undefined || raw === undefined || seen.has(id)) return;
      if (seen.size >= max) {
        over = true;
        return;
      }
      if (options.budget !== undefined && options.budget.left-- <= 0) {
        spent = true;
        return;
      }
      seen.add(id);
      const written = attrs.Type ?? '';
      const kind = lastSegment(written);
      if (options.keep !== undefined && !options.keep(kind, id)) return;
      const type = types.get(written) ?? (types.set(written, written), written);
      // a target with a URI scheme names nothing in the package even without `TargetMode`, so it is kept as written and counts as external
      const external = (attrs.TargetMode ?? '').toLowerCase() === 'external' || SCHEME.test(raw.trim());
      out.push({ id, type, kind, target: external ? raw : resolveRelationshipTarget(owner, raw), external });
    },
    close() {
      depth--;
    },
    halted: () => over || spent,
  });
  if (spent) options.problem?.('budget');
  else if (over) options.problem?.('over');
  else if (!wellFormed) options.problem?.('damaged');
  return out;
}

/** Relationships of one kind (case-insensitive: LibreOffice writes `officedocument` in some type URIs). */
export function relsOfKind(rels: readonly XlsxRelationship[], kind: string): XlsxRelationship[] {
  const k = kind.toLowerCase();
  return rels.filter((r) => r.kind.toLowerCase() === k);
}
