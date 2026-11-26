import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { RevisionViewBanner, isEditLocked, revisionViewMessage } from './RevisionViewBanner';

describe('RevisionViewBanner', () => {
  it('says which text is shown and offers the switch back to All markup', () => {
    const final = renderToStaticMarkup(<RevisionViewBanner view="final" onSwitchToMarkup={() => {}} />);
    expect(final).toContain('data-testid="revision-view-banner"');
    expect(final).toContain('Showing final text: switch to All markup to edit');
    expect(final).toContain('Switch to All markup');
    const original = renderToStaticMarkup(<RevisionViewBanner view="original" onSwitchToMarkup={() => {}} />);
    expect(original).toContain('Showing original text: switch to All markup to edit');
    expect(original).toContain('data-view="original"');
  });

  it('is not shown in All markup', () => {
    expect(renderToStaticMarkup(<RevisionViewBanner view="markup" onSwitchToMarkup={() => {}} />)).toBe('');
    expect(revisionViewMessage('markup')).toBeNull();
  });

  it('locks editing in the Final and Original views and when the host asked for read-only', () => {
    expect(isEditLocked(false, 'markup')).toBe(false);
    expect(isEditLocked(false, 'final')).toBe(true);
    expect(isEditLocked(false, 'original')).toBe(true);
    expect(isEditLocked(true, 'markup')).toBe(true);
  });
});
