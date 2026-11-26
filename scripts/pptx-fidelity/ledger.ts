/**
 * The package ledger: every part of an input package gets exactly one status.
 *
 *   mapped       the part's content will flow into the slides-core model (or it is package plumbing that an export
 *                rebuilds). A mapped part can still carry features the model lacks; those are reported through
 *                loss ids, not through the ledger. Nothing is mapped today: there is no importer, so a `mapped` rule
 *                turns into `unaccounted` for every file (the import "failed"). The first importer PR (K9) flips this.
 *   opaque       the part is kept byte-for-byte and written back on export. Nothing is opaque until the opaque
 *                part store (plan X5) exists: `opaque` is the hook that will decide it.
 *   dropped      the part is not imported and the loss report says so: a loss id covers it (`loss` lists the ids).
 *   unaccounted  nothing accounts for the part: it is not read and no reported loss covers it. The count is
 *                recorded as is (plan row L13 wants 0). A part no rule knows is covered by the detector's
 *                catch-all `pptx.package.unaccountedParts`, so it is `dropped` rather than silent.
 *
 * Rules are the first-match table below, keyed by part name. A `dropped` rule only holds when the file's loss
 * report really contains one of the rule's loss ids; otherwise the part is `unaccounted`. If the import threw (it
 * always does today), `mapped` parts become `unaccounted` as well; dropped parts keep their status, because the loss
 * report is produced independently of the import.
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

const AUDIO_VIDEO = /\.(?:mp3|mp4|m4a|m4v|mov|wav|wma|wmv|avi|mpg|mpeg|mid|midi|ogg|oga|ogv|webm|flac|aac)$/i;
const VECTOR = /\.(?:emf|wmf|emz|wmz)$/i;

export const LEDGER_RULES: readonly LedgerRule[] = [
  { id: 'package.contentTypes', match: /^\[Content_Types\]\.xml$/i, outcome: 'mapped' },
  { id: 'package.relationships', match: /\.rels$/i, outcome: 'mapped' },
  // The writer rebuilds document properties; presentation and view properties and the table style list are settings
  // the exporter will write afresh; printer settings are a binary printer configuration with no content.
  { id: 'docProps', match: /^docProps\//, outcome: 'mapped' },
  { id: 'presentation', match: /^ppt\/presentation\.xml$/, outcome: 'mapped' },
  { id: 'presProps', match: /^ppt\/(?:presProps|viewProps)\.xml$/, outcome: 'mapped' },
  { id: 'tableStyles', match: /^ppt\/tableStyles\.xml$/, outcome: 'mapped' },
  { id: 'printerSettings', match: /^ppt\/printerSettings\//, outcome: 'mapped' },
  { id: 'slide', match: /^ppt\/slides\/[^/]+\.xml$/, outcome: 'mapped' },
  { id: 'layout', match: /^ppt\/slideLayouts\/[^/]+\.xml$/, outcome: 'mapped' },
  { id: 'master', match: /^ppt\/slideMasters\/[^/]+\.xml$/, outcome: 'mapped' },
  { id: 'theme', match: /^ppt\/theme\//, outcome: 'mapped' },
  { id: 'notesSlide', match: /^ppt\/notesSlides\//, outcome: { dropped: ['pptx.notes'] } },
  {
    id: 'notesMaster',
    match: /^ppt\/notesMasters\//,
    outcome: { dropped: ['pptx.masters.notesMaster'] },
  },
  {
    id: 'handoutMaster',
    match: /^ppt\/handoutMasters\//,
    outcome: { dropped: ['pptx.masters.handoutMaster'] },
  },
  { id: 'chart', match: /^ppt\/charts\//, outcome: { dropped: ['pptx.charts'] } },
  { id: 'diagram', match: /^ppt\/diagrams\//, outcome: { dropped: ['pptx.smartArt'] } },
  { id: 'audioVideo', match: new RegExp(`^ppt/media/.*${AUDIO_VIDEO.source}`, 'i'), outcome: { dropped: ['pptx.media'] } },
  {
    id: 'vectorPicture',
    match: new RegExp(`^ppt/media/.*${VECTOR.source}`, 'i'),
    outcome: { dropped: ['pptx.images.unsupportedFormat'] },
  },
  // every other media part is a picture the importer will read
  { id: 'picture', match: /^ppt\/media\//, outcome: 'mapped' },
  {
    id: 'embedding',
    match: /^ppt\/embeddings\//,
    outcome: { dropped: ['pptx.embeddedObjects', 'pptx.charts'] },
  },
  {
    id: 'controls',
    match: /^ppt\/activeX\//,
    outcome: { dropped: ['pptx.embeddedObjects'] },
  },
  {
    id: 'comments',
    match: /^ppt\/(?:comments\/|commentAuthors\.xml$|authors\.xml$)/,
    outcome: { dropped: ['pptx.comments'] },
  },
  { id: 'ink', match: /^ppt\/ink\//, outcome: { dropped: ['pptx.ink'] } },
  {
    id: 'macro',
    match: /^ppt\/vbaProject(?:Signature)?\.bin$/,
    outcome: { dropped: ['pptx.macros'] },
  },
  { id: 'signature', match: /^_xmlsignatures\//, outcome: { dropped: ['pptx.signatures'] } },
  // Any other part: covered only by the detector's catch-all report (`full` detail), which names it.
  {
    id: 'unknownPart',
    match: /./,
    outcome: { dropped: ['pptx.package.unaccountedParts'] },
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
