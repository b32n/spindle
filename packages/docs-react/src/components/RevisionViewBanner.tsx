/**
 * RevisionViewBanner - shown while the document is displayed as its final or its original text. Those views hide text
 * (as zero-width placeholders), so typing in them would write text nobody can see: the editor is read-only until the
 * reader switches back to All markup, which the banner offers in one click.
 */

import { memo } from 'react';
import type { RevisionView } from '../core';

export interface RevisionViewBannerProps {
  view: RevisionView;
  /** Called to go back to All markup. */
  onSwitchToMarkup: () => void;
}

/** True when the editor cannot be edited: the host asked for it, or the document shows only its final or original text. */
export function isEditLocked(readOnly: boolean, view: RevisionView): boolean {
  return readOnly || view !== 'markup';
}

/** What the banner says for a view (null for All markup, where there is no banner). */
export function revisionViewMessage(view: RevisionView): string | null {
  if (view === 'final') return 'Showing final text: switch to All markup to edit';
  if (view === 'original') return 'Showing original text: switch to All markup to edit';
  return null;
}

export const RevisionViewBanner = memo(function RevisionViewBanner({ view, onSwitchToMarkup }: RevisionViewBannerProps) {
  const message = revisionViewMessage(view);
  if (!message) return null;
  return (
    <div
      data-testid="revision-view-banner"
      data-view={view}
      role="status"
      style={{
        display: 'flex',
        alignItems: 'center',
        gap: 12,
        padding: '6px 16px',
        background: '#fef7e0',
        borderBottom: '1px solid #f0d58c',
        font: '13px/18px Arial, sans-serif',
        color: '#3c4043',
      }}
    >
      <span>{message}</span>
      <button
        type="button"
        onClick={onSwitchToMarkup}
        style={{
          font: 'inherit',
          padding: '2px 10px',
          border: '1px solid #dadce0',
          borderRadius: 4,
          background: '#ffffff',
          color: '#1a73e8',
          cursor: 'pointer',
        }}
      >
        Switch to All markup
      </button>
    </div>
  );
});
