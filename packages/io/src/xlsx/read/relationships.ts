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

/** What went wrong in a part read as a list of entries: `damaged` (not well-formed: what precedes the damage was read) or `over` (more entries than the limit: the first ones were kept). */
export type ListProblem = 'damaged' | 'over';

export interface ListOptions {
  /** Most entries kept (`ImportLimits.maxPartElements`); the scan ends at the first one over it. */
  max?: number;
  problem?: (what: ListProblem) => void;
}

/**
 * Relationships of `owner` from the text of its `.rels`, read as events: memory is the relationships kept (at most `options.max`),
 * not a tree of the part. Entries without an `Id` or a `Target` are skipped; a repeated id keeps the first. Damaged markup ends
 * the read there (`problem('damaged')`); an entry over the limit ends it too (`problem('over')`).
 */
export function parseRelationships(xml: string | undefined, owner: string, options: ListOptions = {}): XlsxRelationship[] {
  const out: XlsxRelationship[] = [];
  if (xml === undefined) return out;
  const max = options.max ?? Infinity;
  const seen = new Set<string>();
  let depth = 0;
  let over = false;
  const wellFormed = scanXml(xml, {
    open(name, attrs) {
      depth++;
      const id = attrs.Id;
      const raw = attrs.Target;
      if (depth !== 2 || name !== 'Relationship' || id === undefined || raw === undefined || seen.has(id)) return;
      if (out.length >= max) {
        over = true;
        return;
      }
      seen.add(id);
      const type = attrs.Type ?? '';
      // a target with a URI scheme names nothing in the package even without `TargetMode`, so it is kept as written and counts as external
      const external = (attrs.TargetMode ?? '').toLowerCase() === 'external' || SCHEME.test(raw.trim());
      out.push({ id, type, kind: lastSegment(type), target: external ? raw : resolveRelationshipTarget(owner, raw), external });
    },
    close() {
      depth--;
    },
    halted: () => over,
  });
  if (over) options.problem?.('over');
  else if (!wellFormed) options.problem?.('damaged');
  return out;
}

/** Relationships of one kind (case-insensitive: LibreOffice writes `officedocument` in some type URIs). */
export function relsOfKind(rels: readonly XlsxRelationship[], kind: string): XlsxRelationship[] {
  const k = kind.toLowerCase();
  return rels.filter((r) => r.kind.toLowerCase() === k);
}
