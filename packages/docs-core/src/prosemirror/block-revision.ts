// The tracked changes of a paragraph mark or a table row, as node attributes.
//
// A change is kept per kind in its own attribute, so two people suggesting different things on the same paragraph mark
// or row at once (one splits it, the other marks it deleted) each set a different Y attribute and both survive; one
// attribute holding all three would be replaced as a whole (last writer wins).

import type { BlockRevision } from '../types';

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

/** The three attributes for a model revision (null for a kind it has no change of). */
export function blockRevisionAttrs(revision: BlockRevision | null | undefined): BlockRevisionAttrs {
  return {
    revisionInsertion: revision?.insertion ? cloneJson(revision.insertion) : null,
    revisionDeletion: revision?.deletion ? cloneJson(revision.deletion) : null,
    revisionFormat: revision?.formatChange ? cloneJson(revision.formatChange) : null,
  };
}

/**
 * The model revision of a node's attributes, or undefined when it has none. Values that are not objects are ignored.
 * The records are copies unless `copy` is false (for reading only: the records are the node's own objects).
 */
export function readBlockRevision(attrs: Record<string, unknown>, copy = true): BlockRevision | undefined {
  const revision: BlockRevision = {};
  const take = <T,>(value: unknown): T => (copy ? cloneJson(value as T) : (value as T));
  if (isRecord(attrs.revisionInsertion)) revision.insertion = take<NonNullable<BlockRevision['insertion']>>(attrs.revisionInsertion);
  if (isRecord(attrs.revisionDeletion)) revision.deletion = take<NonNullable<BlockRevision['deletion']>>(attrs.revisionDeletion);
  if (isRecord(attrs.revisionFormat)) revision.formatChange = take<NonNullable<BlockRevision['formatChange']>>(attrs.revisionFormat);
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
