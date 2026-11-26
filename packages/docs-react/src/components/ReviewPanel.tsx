/**
 * ReviewPanel - the tracked changes of the document, each with Accept and Reject, a choice of how the changes show
 * (all markup, final, original), going to the next and previous change, and accepting or rejecting all of them or all of
 * one author's.
 */

import React, { memo, useEffect, useId, useMemo, useState } from 'react';
import { X, Check, ChevronUp, ChevronDown } from 'lucide-react';
import { assignAuthorColors, type ReviewChange } from '@b32nio/spindle-docs-core';
import type { RevisionView } from '../core';

export interface ReviewPanelProps {
  changes: readonly ReviewChange[];
  view: RevisionView;
  onViewChange: (view: RevisionView) => void;
  /** Accept or reject one change, or all of them (all of one author's with `author`). Absent in a read-only editor. */
  onResolve?: (action: 'accept' | 'reject', target: { change: ReviewChange } | { author?: string }) => void;
  /** Go to a change in the document. */
  onGoTo: (change: ReviewChange) => void;
  onClose: () => void;
  /** Every author of the document, whose colours are shared out (default: the authors of `changes`). */
  palette?: readonly string[];
  /** Colours some authors are shown in (as the markup is); the others share out the palette. */
  authorColors?: ReadonlyMap<string, string>;
}

/** The value of the author choice that stands for every author (an author's own name can be any text, also empty). */
const ALL_AUTHORS = '\u0000all';

const FONT = '"Inter", "SF Pro Display", -apple-system, BlinkMacSystemFont, sans-serif';

const KIND_LABEL: Record<ReviewChange['kind'], string> = {
  insertion: 'Inserted',
  deletion: 'Deleted',
  move: 'Moved',
  format: 'Formatted',
  paragraphInserted: 'Inserted paragraph break',
  paragraphDeleted: 'Deleted paragraph break',
  paragraphFormat: 'Changed paragraph formatting',
  rowInserted: 'Inserted table row',
  rowDeleted: 'Deleted table row',
};

const VIEWS: Array<{ id: RevisionView; label: string }> = [
  { id: 'markup', label: 'All markup' },
  { id: 'final', label: 'Final' },
  { id: 'original', label: 'Original' },
];

const button: React.CSSProperties = {
  border: '1px solid #dadce0',
  borderRadius: 6,
  background: '#fff',
  color: '#3c4043',
  font: `12px/16px ${FONT}`,
  padding: '3px 8px',
  cursor: 'pointer',
};

const EXCERPT = 80;
const excerpt = (text: string) => {
  const flat = text.replace(new RegExp('[\\u200B\\s]+', 'g'), ' ').trim();
  return flat.length > EXCERPT ? `${flat.slice(0, EXCERPT)}…` : flat;
};

function when(iso: string | undefined): string {
  if (!iso) return '';
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? '' : d.toLocaleDateString();
}

/**
 * The change after (or before) the current one, wrapping around; the first (last) one when there is no current change yet.
 * `changes` are in document order. The current change is found by its id when it is still listed (so changes that start at
 * one position are visited one after the other); otherwise, after an edit changed the ids, by position: the first change
 * after (before) `from`.
 */
export function neighbourChange(
  changes: readonly ReviewChange[],
  from: number | null,
  direction: 'next' | 'previous',
  currentId: string | null = null
): ReviewChange | undefined {
  if (changes.length === 0) return undefined;
  const at = currentId === null ? -1 : changes.findIndex((c) => c.id === currentId);
  if (at >= 0) return changes[(at + (direction === 'next' ? 1 : changes.length - 1)) % changes.length];
  if (direction === 'next') return (from === null ? undefined : changes.find((c) => c.from > from)) ?? changes[0];
  return (from === null ? undefined : [...changes].reverse().find((c) => c.from < from)) ?? changes[changes.length - 1];
}

export const ReviewPanel = memo(function ReviewPanel({ changes, view, onViewChange, onResolve, onGoTo, onClose, palette, authorColors }: ReviewPanelProps) {
  // The author whose changes are shown; null: all of them (an empty author name is an author like another).
  const [author, setAuthor] = useState<string | null>(null);
  const authorSelectId = useId();
  const [currentId, setCurrentId] = useState<string | null>(null);
  // Where the last change gone to was: next and previous go on from here even after the document changed.
  const [currentFrom, setCurrentFrom] = useState<number | null>(null);
  const authors = useMemo(() => [...new Set(changes.map((c) => c.author))], [changes]);
  // The colours are shared out among all authors of the document (a move's other end can have another author), as on the page.
  const colors = useMemo(() => assignAuthorColors(palette ?? authors, authorColors), [palette, authors, authorColors]);
  // A filter on an author who has no changes left shows everybody again.
  useEffect(() => {
    if (author !== null && !authors.includes(author)) setAuthor(null);
  }, [author, authors]);
  const shown = useMemo(() => (author !== null ? changes.filter((c) => c.author === author) : changes), [changes, author]);

  const go = (change: ReviewChange | undefined) => {
    if (!change) return;
    setCurrentId(change.id);
    setCurrentFrom(change.from);
    onGoTo(change);
  };
  const next = () => go(neighbourChange(shown, currentFrom, 'next', currentId));
  const previous = () => go(neighbourChange(shown, currentFrom, 'previous', currentId));

  return (
    <aside
      data-testid="review-panel"
      aria-label="Review changes"
      style={{ display: 'flex', flexDirection: 'column', height: '100%', background: '#fff', font: `13px/18px ${FONT}`, color: '#202124' }}
    >
      <header style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '10px 12px', borderBottom: '1px solid #e8eaed' }}>
        <strong style={{ flex: 1, fontSize: 14 }}>Review changes</strong>
        <button type="button" aria-label="Close review panel" onClick={onClose} style={{ ...button, border: 'none', padding: 4, display: 'inline-flex' }}>
          <X size={16} />
        </button>
      </header>

      <div style={{ padding: '8px 12px', display: 'flex', flexDirection: 'column', gap: 8, borderBottom: '1px solid #e8eaed' }}>
        <div role="group" aria-label="Show changes as" style={{ display: 'flex', gap: 4 }}>
          {VIEWS.map((v) => (
            <button
              key={v.id}
              type="button"
              aria-pressed={view === v.id}
              onClick={() => onViewChange(v.id)}
              style={{ ...button, ...(view === v.id ? { background: '#e8f0fe', border: '1px solid #aecbfa', color: '#1967d2' } : {}) }}
            >
              {v.label}
            </button>
          ))}
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <label htmlFor={authorSelectId} style={{ color: '#5f6368' }}>
            Author
          </label>
          <select id={authorSelectId} value={author ?? ALL_AUTHORS} onChange={(e) => setAuthor(e.target.value === ALL_AUTHORS ? null : e.target.value)} style={{ ...button, flex: 1 }}>
            <option value={ALL_AUTHORS}>All authors</option>
            {authors.map((a) => (
              <option key={a} value={a}>
                {a || 'Unknown'}
              </option>
            ))}
          </select>
          <button type="button" aria-label="Previous change" disabled={shown.length === 0} onClick={previous} style={{ ...button, padding: 3, display: 'inline-flex' }}>
            <ChevronUp size={16} />
          </button>
          <button type="button" aria-label="Next change" disabled={shown.length === 0} onClick={next} style={{ ...button, padding: 3, display: 'inline-flex' }}>
            <ChevronDown size={16} />
          </button>
        </div>
        {onResolve && (
          <div style={{ display: 'flex', gap: 6 }}>
            <button type="button" disabled={shown.length === 0} onClick={() => onResolve('accept', { author: author ?? undefined })} style={button}>
              {author !== null ? `Accept all by ${author || 'Unknown'}` : 'Accept all'}
            </button>
            <button type="button" disabled={shown.length === 0} onClick={() => onResolve('reject', { author: author ?? undefined })} style={button}>
              {author !== null ? `Reject all by ${author || 'Unknown'}` : 'Reject all'}
            </button>
          </div>
        )}
      </div>

      <div role="list" aria-label="Changes" style={{ flex: 1, overflowY: 'auto' }}>
        {shown.length === 0 && <p style={{ padding: 16, color: '#5f6368', margin: 0 }}>No tracked changes.</p>}
        {shown.map((change) => {
          const text = excerpt(change.text);
          return (
            <div
              key={change.id}
              role="listitem"
              data-testid="review-change"
              data-kind={change.kind}
              data-author={change.author}
              aria-current={change.id === currentId ? 'true' : undefined}
              tabIndex={0}
              onClick={() => go(change)}
              onKeyDown={(e) => {
                // Only the row itself: Enter or Space on its Accept and Reject buttons are theirs.
                if (e.target === e.currentTarget && (e.key === 'Enter' || e.key === ' ')) {
                  e.preventDefault();
                  go(change);
                }
              }}
              style={{
                padding: '8px 12px',
                borderBottom: '1px solid #f1f3f4',
                borderLeft: `3px solid ${colors.get(change.author) ?? '#c5221f'}`,
                background: change.id === currentId ? '#f1f6ff' : undefined,
                cursor: 'pointer',
              }}
            >
              <div style={{ display: 'flex', gap: 6, alignItems: 'baseline' }}>
                <span style={{ fontWeight: 600 }}>{change.author || 'Unknown'}</span>
                <span style={{ color: '#80868b', fontSize: 12 }}>{when(change.date)}</span>
              </div>
              <div style={{ color: '#5f6368' }}>{KIND_LABEL[change.kind]}</div>
              {text && <div style={{ marginTop: 2, wordBreak: 'break-word' }}>“{text}”</div>}
              {onResolve && (
                <div style={{ display: 'flex', gap: 6, marginTop: 6 }}>
                  <button
                    type="button"
                    aria-label="Accept change"
                    onClick={(e) => {
                      e.stopPropagation();
                      onResolve('accept', { change });
                    }}
                    style={{ ...button, display: 'inline-flex', alignItems: 'center', gap: 4 }}
                  >
                    <Check size={14} /> Accept
                  </button>
                  <button
                    type="button"
                    aria-label="Reject change"
                    onClick={(e) => {
                      e.stopPropagation();
                      onResolve('reject', { change });
                    }}
                    style={{ ...button, display: 'inline-flex', alignItems: 'center', gap: 4 }}
                  >
                    <X size={14} /> Reject
                  </button>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </aside>
  );
});
