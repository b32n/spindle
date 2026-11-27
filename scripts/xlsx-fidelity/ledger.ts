/**
 * The package ledger: every part of an input package gets exactly one status.
 *
 *   mapped       the part's content flows into the sheets-core model (or it is package plumbing that an export
 *                rebuilds). A mapped part can still carry features the model lacks; those are reported through
 *                loss ids, not through the ledger.
 *   opaque       the part is kept byte-for-byte and written back on export. Nothing is opaque until the opaque
 *                part store (plan X5) exists: `opaque` is the hook that will decide it.
 *   dropped      the part is not imported and the loss report says so: a loss id covers it (`loss` lists the ids).
 *   unaccounted  nothing accounts for the part: it is not read and no reported loss covers it. The count is
 *                recorded as is (plan row L8 wants 0). A part no rule knows is covered by the detector's
 *                catch-all `xlsx.package.unaccountedParts`, so it is `dropped` rather than silent.
 *
 * Rules are the first-match table below, keyed by part name. A `dropped` rule only holds when the file's loss
 * report really contains one of the rule's loss ids; otherwise the part is `unaccounted`. If the import threw,
 * nothing reached the model, so `mapped` parts become `unaccounted` as well (dropped parts keep their status:
 * the loss report is produced independently of the import).
 */
import type { OoxmlParts } from '../fidelity-core/ooxml-package';
import type { LedgerTotals } from '../fidelity-core/types';

export type PartStatus = 'mapped' | 'opaque' | 'dropped' | 'unaccounted';

export interface LedgerRule {
  /** Stable name of the rule, recorded next to each part so a golden diff says why a status changed. */
  id: string;
  match: RegExp;
  /** `mapped`, or the loss ids that would cover the part if reported (the first one present wins). */
  outcome: 'mapped' | { dropped: string[] };
}

export const LEDGER_RULES: readonly LedgerRule[] = [
  { id: 'package.contentTypes', match: /^\[Content_Types\]\.xml$/i, outcome: 'mapped' },
  { id: 'package.relationships', match: /\.rels$/i, outcome: 'mapped' },
  // The writer rebuilds document properties and the calculation chain is a cache: nothing to carry over.
  { id: 'docProps', match: /^docProps\//, outcome: 'mapped' },
  { id: 'calcChain', match: /^xl\/calcChain\.xml$/, outcome: 'mapped' },
  { id: 'workbook', match: /^xl\/workbook\.xml$/, outcome: 'mapped' },
  { id: 'worksheet', match: /^xl\/worksheets\/[^/]+\.xml$/, outcome: 'mapped' },
  { id: 'sharedStrings', match: /^xl\/sharedStrings\.xml$/, outcome: 'mapped' },
  { id: 'styles', match: /^xl\/styles\.xml$/, outcome: 'mapped' },
  // theme colours are resolved into the cell styles on import (S2b)
  { id: 'theme', match: /^xl\/theme\//, outcome: 'mapped' },
  {
    id: 'printerSettings',
    match: /^xl\/printerSettings\//,
    outcome: { dropped: ['xlsx.print.setup'] },
  },
  {
    id: 'metadata',
    match: /^xl\/metadata\.xml$/,
    outcome: { dropped: ['xlsx.formulas.dynamicArray'] },
  },
  { id: 'chartsheet', match: /^xl\/chartsheets\//, outcome: { dropped: ['xlsx.charts'] } },
  {
    id: 'vmlDrawing',
    match: /^xl\/drawings\/[^/]*\.vml$/,
    outcome: { dropped: ['xlsx.comments', 'xlsx.objects.formControls', 'xlsx.images'] },
  },
  {
    id: 'drawing',
    match: /^xl\/drawings\//,
    outcome: {
      dropped: [
        'xlsx.images',
        'xlsx.charts',
        'xlsx.objects.shapes',
        'xlsx.objects.textBoxes',
        'xlsx.objects.formControls',
        'xlsx.objects.embeddedObjects',
      ],
    },
  },
  {
    id: 'media',
    match: /^xl\/media\//,
    outcome: { dropped: ['xlsx.images', 'xlsx.objects.cellImages'] },
  },
  { id: 'chart', match: /^xl\/charts\//, outcome: { dropped: ['xlsx.charts'] } },
  {
    id: 'comments',
    match: /^xl\/(comments[^/]*\.xml|comments\/|threadedComments\/|persons\/)/,
    outcome: { dropped: ['xlsx.comments', 'xlsx.comments.threaded'] },
  },
  { id: 'table', match: /^xl\/tables\//, outcome: { dropped: ['xlsx.tables'] } },
  {
    id: 'pivot',
    match: /^xl\/(pivotTables|pivotCache)\//,
    outcome: { dropped: ['xlsx.pivotTables'] },
  },
  { id: 'slicer', match: /^xl\/(slicers|slicerCaches)\//, outcome: { dropped: ['xlsx.slicers'] } },
  {
    id: 'externalLink',
    match: /^xl\/externalLinks\//,
    outcome: { dropped: ['xlsx.externalLinks'] },
  },
  {
    id: 'controls',
    match: /^xl\/(ctrlProps|activeX)\//,
    outcome: { dropped: ['xlsx.objects.formControls'] },
  },
  {
    id: 'embedding',
    match: /^xl\/embeddings\//,
    outcome: { dropped: ['xlsx.objects.embeddedObjects'] },
  },
  {
    id: 'richData',
    match: /^xl\/(richData\/|cellimages\.xml$)/,
    outcome: {
      dropped: ['xlsx.workbook.richData', 'xlsx.objects.cellImages', 'xlsx.images'],
    },
  },
  {
    id: 'externalData',
    match: /^xl\/(connections\.xml$|queryTables\/|model\/)/,
    outcome: { dropped: ['xlsx.workbook.externalData'] },
  },
  {
    id: 'macro',
    match: /^xl\/vbaProject\.bin$|^xl\/vbaProjectSignature\.bin$/,
    outcome: { dropped: ['xlsx.macros'] },
  },
  // Any other part: covered only by the detector's catch-all report (`full` detail), which names it.
  {
    id: 'unknownPart',
    match: /./,
    outcome: { dropped: ['xlsx.package.unaccountedParts'] },
  },
];

export interface LedgerPart {
  name: string;
  status: PartStatus;
  /** The rule that decided, or `none` when no rule matched. */
  rule: string;
  /** For `dropped`: the loss ids that cover the part. */
  loss?: string[];
}

export interface LedgerSnapshot {
  /** `zip` = readable package, `ole` = OLE compound file (encrypted or legacy binary), `unreadable` = neither. */
  package: OoxmlParts['kind'];
  importError: boolean;
  totals: LedgerTotals;
  parts: LedgerPart[];
}

export interface LedgerInput {
  parts: Pick<OoxmlParts, 'kind' | 'names'>;
  /** Ids of the loss items reported for the file. */
  lossIds: string[];
  importFailed: boolean;
  /** Part names stored in the opaque part store (none today). */
  isOpaque?: (name: string) => boolean;
}

export function classifyPart(
  name: string,
  input: Pick<LedgerInput, 'lossIds' | 'importFailed' | 'isOpaque'>
): LedgerPart {
  if (input.isOpaque?.(name)) return { name, status: 'opaque', rule: 'opaque' };
  const rule = LEDGER_RULES.find((r) => r.match.test(name));
  if (!rule) return { name, status: 'unaccounted', rule: 'none' };
  if (rule.outcome === 'mapped')
    return { name, status: input.importFailed ? 'unaccounted' : 'mapped', rule: rule.id };
  const covering = rule.outcome.dropped.filter((id) => input.lossIds.includes(id));
  if (covering.length === 0) return { name, status: 'unaccounted', rule: rule.id };
  return { name, status: 'dropped', rule: rule.id, loss: [covering[0]] };
}

export function buildLedger(input: LedgerInput): LedgerSnapshot {
  const parts = [...input.parts.names].sort().map((n) => classifyPart(n, input));
  const count = (s: PartStatus) => parts.filter((p) => p.status === s).length;
  return {
    package: input.parts.kind,
    importError: input.importFailed,
    totals: {
      mapped: count('mapped'),
      opaque: count('opaque'),
      dropped: count('dropped'),
      unaccounted: count('unaccounted'),
      total: parts.length,
    },
    parts,
  };
}
