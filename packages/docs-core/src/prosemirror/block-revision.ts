// The tracked changes of a paragraph mark or a table row, as node attributes.
//
// A change is kept per kind in its own attribute, so two people suggesting different things on the same paragraph mark
// or row at once (one splits it, the other marks it deleted) each set a different Y attribute and both survive; one
// attribute holding all three would be replaced as a whole (last writer wins).

import type { BlockRevision, LineSpacing, PreviousParagraphFormat, RevisionMeta } from '../types';

/** Node attributes that hold a paragraph's or row's tracked changes: an insertion, a deletion, a formatting change. */
export const BLOCK_REVISION_ATTRS = {
  revisionInsertion: { default: null },
  revisionDeletion: { default: null },
  revisionFormat: { default: null },
};

export type BlockRevisionKind = keyof BlockRevision;

export interface BlockRevisionAttrs {
  revisionInsertion: BlockRevision['insertion'] | null;
  revisionDeletion: BlockRevision['deletion'] | null;
  revisionFormat: BlockRevision['formatChange'] | null;
}

/** The node attribute that holds a kind of change. */
export const BLOCK_REVISION_ATTR_OF: Record<BlockRevisionKind, keyof BlockRevisionAttrs> = {
  insertion: 'revisionInsertion',
  deletion: 'revisionDeletion',
  formatChange: 'revisionFormat',
};

/** A plain-data copy (revisions are stored in node attributes, which must not share objects with the model). */
export function cloneJson<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

const isRecord = (value: unknown): value is Record<string, unknown> => !!value && typeof value === 'object' && !Array.isArray(value);

const finiteNumber = (v: unknown): v is number => typeof v === 'number' && Number.isFinite(v) && Math.abs(v) < 100000;
const ALIGNMENTS = ['left', 'center', 'right', 'justify'];
const LINE_SPACINGS = ['single', 'onePointFive', 'double', 'atLeast', 'exactly', 'multiple'];

/**
 * The paragraph formatting of a change record, from a source that is not trusted (another client's Y document): known
 * fields of the right type only, so a malformed value cannot reach the DOCX writer.
 */
export function sanitizePreviousParagraphFormat(raw: unknown): PreviousParagraphFormat {
  if (!isRecord(raw)) return {};
  const out: Record<string, unknown> = {};
  if (typeof raw.alignment === 'string' && ALIGNMENTS.includes(raw.alignment)) out.alignment = raw.alignment;
  if (isRecord(raw.lineSpacing) && typeof raw.lineSpacing.type === 'string' && LINE_SPACINGS.includes(raw.lineSpacing.type)) {
    const spacing: LineSpacing = { type: raw.lineSpacing.type as LineSpacing['type'] };
    if (finiteNumber(raw.lineSpacing.value)) spacing.value = raw.lineSpacing.value;
    out.lineSpacing = spacing;
  }
  for (const key of ['spaceBefore', 'spaceAfter', 'leftIndent', 'rightIndent', 'firstLineIndent'] as const) if (finiteNumber(raw[key])) out[key] = raw[key];
  for (const key of ['keepWithNext', 'keepTogether', 'pageBreakBefore'] as const) if (typeof raw[key] === 'boolean') out[key] = raw[key];
  if (typeof raw.namedStyleId === 'string' && raw.namedStyleId.length > 0 && raw.namedStyleId.length <= 100) out.namedStyleId = raw.namedStyleId;
  if (typeof raw.headingLevel === 'number' && Number.isInteger(raw.headingLevel) && raw.headingLevel >= 1 && raw.headingLevel <= 6) out.headingLevel = raw.headingLevel;
  return out;
}

/** The three attributes for a model revision (null for a kind it has no change of). */
export function blockRevisionAttrs(revision: BlockRevision | null | undefined): BlockRevisionAttrs {
  return {
    revisionInsertion: revision?.insertion ? cloneJson(revision.insertion) : null,
    revisionDeletion: revision?.deletion ? cloneJson(revision.deletion) : null,
    revisionFormat: revision?.formatChange ? cloneJson(revision.formatChange) : null,
  };
}

const MAX_ID = 200;

/** A change record from a source that is not trusted: an id (a string or a number) and strings for the rest, or nothing when it has no usable id. */
function sanitizeMeta(raw: unknown): RevisionMeta | undefined {
  if (!isRecord(raw)) return undefined;
  const id = typeof raw.id === 'number' && Number.isFinite(raw.id) ? String(raw.id) : raw.id;
  if (typeof id !== 'string' || id.length === 0 || id.length > MAX_ID) return undefined;
  const meta: RevisionMeta = { id, author: typeof raw.author === 'string' && raw.author.length <= MAX_ID ? raw.author : 'Unknown' };
  if (typeof raw.date === 'string' && raw.date.length <= 64) meta.date = raw.date;
  if (typeof raw.moveId === 'string' && raw.moveId.length > 0 && raw.moveId.length <= MAX_ID) meta.moveId = raw.moveId;
  return meta;
}

/**
 * The model revision of a node's attributes, or undefined when it has none. The attributes may come from another client's
 * Y document, so each record is rebuilt from the fields it should have (an id, an author, a date, a move name, the
 * formatting before); a record without a usable id is ignored, and one without `previous` has an empty one. Always a copy:
 * `copy` is kept for callers that pass false and means nothing now that the records are rebuilt.
 */
export function readBlockRevision(attrs: Record<string, unknown>, _copy = true): BlockRevision | undefined {
  const revision: BlockRevision = {};
  const insertion = sanitizeMeta(attrs.revisionInsertion);
  const deletion = sanitizeMeta(attrs.revisionDeletion);
  if (insertion) revision.insertion = insertion;
  if (deletion) revision.deletion = deletion;
  const format = sanitizeMeta(attrs.revisionFormat);
  if (format) {
    const { moveId: _moveId, ...meta } = format;
    revision.formatChange = { ...meta, previous: sanitizePreviousParagraphFormat(isRecord(attrs.revisionFormat) ? cloneJson(attrs.revisionFormat.previous ?? {}) : {}) };
  }
  return revision.insertion || revision.deletion || revision.formatChange ? revision : undefined;
}

/** The attributes with one kind of change set (or cleared with null); the other kinds stay. */
export function withBlockRevisionKind<A extends Record<string, unknown>>(attrs: A, kind: BlockRevisionKind, value: unknown): A {
  return { ...attrs, [BLOCK_REVISION_ATTR_OF[kind]]: value ?? null };
}

/** The attributes with all three kinds replaced by those of `revision`. */
export function withBlockRevision<A extends Record<string, unknown>>(attrs: A, revision: BlockRevision | null | undefined): A {
  return { ...attrs, ...blockRevisionAttrs(revision) };
}
