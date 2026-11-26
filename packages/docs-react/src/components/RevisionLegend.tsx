/**
 * RevisionLegend - the authors of the tracked changes in the document, each in the colour their changes are drawn in.
 */

import { memo } from 'react';
import type { RevisionAuthor } from '@b32nio/spindle-docs-core';

export interface RevisionLegendProps {
  authors: readonly RevisionAuthor[];
}

export const RevisionLegend = memo(function RevisionLegend({ authors }: RevisionLegendProps) {
  if (authors.length === 0) return null;
  return (
    <div
      data-testid="revision-legend"
      role="group"
      aria-label="Authors of tracked changes"
      style={{
        display: 'flex',
        flexWrap: 'wrap',
        alignItems: 'center',
        gap: 12,
        padding: '4px 0',
        font: '12px/16px Arial, sans-serif',
        color: '#3c4043',
      }}
    >
      <span style={{ fontWeight: 600 }}>Tracked changes</span>
      <div role="list" style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 12 }}>
        {authors.map((author) => (
          <span key={author.name} role="listitem" data-revision-author={author.name} style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
            <span aria-hidden="true" style={{ width: 10, height: 10, borderRadius: 2, background: author.color, display: 'inline-block' }} />
            <span>{author.name}</span>
            <span style={{ color: '#80868b' }}>({author.count})</span>
          </span>
        ))}
      </div>
    </div>
  );
});
