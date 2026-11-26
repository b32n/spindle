// The attributes that hold the tracked changes of a paragraph mark or table row, one per kind.

import {
  BLOCK_REVISION_ATTR_OF,
  blockRevisionAttrs,
  readBlockRevision,
  withBlockRevision,
  withBlockRevisionKind,
} from './prosemirror/block-revision';
import type { BlockRevision } from './types';

const ADA = { id: '1', author: 'Ada', date: '2024-01-01T00:00:00Z' };
const BO = { id: '2', author: 'Bo' };
const CY = { id: '3', author: 'Cy', previous: { alignment: 'right' as const } };

describe('block revision attributes', () => {
  it('names one attribute per kind of change', () => {
    expect(BLOCK_REVISION_ATTR_OF).toEqual({ insertion: 'revisionInsertion', deletion: 'revisionDeletion', formatChange: 'revisionFormat' });
  });

  it('sets one kind and leaves the others and other attributes alone, or clears it with null or undefined', () => {
    const attrs = { alignment: 'left', revisionInsertion: ADA, revisionDeletion: null, revisionFormat: null };
    const withDeletion = withBlockRevisionKind(attrs, 'deletion', BO);
    expect(withDeletion).toEqual({ alignment: 'left', revisionInsertion: ADA, revisionDeletion: BO, revisionFormat: null });
    expect(attrs.revisionDeletion).toBeNull(); // the input is not changed
    expect(withBlockRevisionKind(withDeletion, 'insertion', null).revisionInsertion).toBeNull();
    expect(withBlockRevisionKind(withDeletion, 'deletion', undefined).revisionDeletion).toBeNull();
    expect(withBlockRevisionKind(withDeletion, 'deletion', undefined).revisionInsertion).toEqual(ADA);
  });

  it('replaces all three kinds at once, and clears them for no revision', () => {
    const attrs = { alignment: 'left', revisionInsertion: ADA, revisionDeletion: BO, revisionFormat: null };
    const revision: BlockRevision = { formatChange: CY };
    expect(withBlockRevision(attrs, revision)).toEqual({ alignment: 'left', revisionInsertion: null, revisionDeletion: null, revisionFormat: CY });
    expect(withBlockRevision(attrs, undefined)).toEqual({ alignment: 'left', revisionInsertion: null, revisionDeletion: null, revisionFormat: null });
    expect(withBlockRevision(attrs, null)).toEqual(withBlockRevision(attrs, undefined));
  });

  it('reads back what blockRevisionAttrs wrote', () => {
    const revision: BlockRevision = { insertion: ADA, deletion: BO, formatChange: CY };
    expect(readBlockRevision({ ...blockRevisionAttrs(revision) })).toEqual(revision);
    expect(readBlockRevision(withBlockRevision({}, { deletion: BO }))).toEqual({ deletion: BO });
    expect(readBlockRevision({ ...blockRevisionAttrs(undefined) })).toBeUndefined();
  });
});
