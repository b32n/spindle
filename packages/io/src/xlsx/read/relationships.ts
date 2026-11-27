import type { XlsxRelationship } from '@b32nio/spindle-shared/export/xlsx';
import { relsPart, resolveTarget } from '../../common/ooxml-package';
import { kids, type XmlElement } from '../xml';

export { relsPart };

/**
 * The part a relationship target names, resolved against the part that owns the relationships ('' for the package's own
 * `_rels/.rels`). Handles what producers write besides Excel's relative form: absolute (`/xl/worksheets/sheet1.xml`, openpyxl),
 * `./` and `..` segments, backslashes, and percent-encoded names. A target that climbs above the root stays at the root.
 */
export function resolveRelationshipTarget(owner: string, target: string): string {
  let t = target.trim().replace(/\\/g, '/');
  try {
    t = decodeURIComponent(t);
  } catch {
    // not valid percent-encoding: the text is the name
  }
  return resolveTarget(owner, t);
}

const lastSegment = (type: string): string => type.slice(type.lastIndexOf('/') + 1);

/** Relationships of `owner` from its parsed `.rels` root. Entries without an `Id` or a `Target` are skipped; a repeated id keeps the first. */
export function parseRelationships(root: XmlElement | undefined, owner: string): XlsxRelationship[] {
  const out: XlsxRelationship[] = [];
  const seen = new Set<string>();
  for (const r of kids(root, 'Relationship')) {
    const id = r.attrs.Id;
    const raw = r.attrs.Target;
    if (id === undefined || raw === undefined || seen.has(id)) continue;
    seen.add(id);
    const type = r.attrs.Type ?? '';
    const external = (r.attrs.TargetMode ?? '').toLowerCase() === 'external';
    out.push({ id, type, kind: lastSegment(type), target: external ? raw : resolveRelationshipTarget(owner, raw), external });
  }
  return out;
}

/** Relationships of one kind (case-insensitive: LibreOffice writes `officedocument` in some type URIs). */
export function relsOfKind(rels: readonly XlsxRelationship[], kind: string): XlsxRelationship[] {
  const k = kind.toLowerCase();
  return rels.filter((r) => r.kind.toLowerCase() === k);
}
