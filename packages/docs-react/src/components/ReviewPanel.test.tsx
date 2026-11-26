import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { assignAuthorColors, type ReviewChange } from '@b32nio/spindle-docs-core';
import { ReviewPanel, neighbourChange, type ReviewPanelProps } from './ReviewPanel';

const unit = { key: 'ins:1', kind: 'insertion', id: '1', author: 'Ada', from: 1, to: 4, text: 'abc' } as const;
const change = (over: Partial<ReviewChange>): ReviewChange => ({ id: 'c', kind: 'insertion', author: 'Ada', text: 'abc', from: 1, to: 4, units: [unit], ...over });

function render(props: Partial<ReviewPanelProps> = {}): string {
  return renderToStaticMarkup(<ReviewPanel changes={[]} view="markup" onViewChange={() => {}} onGoTo={() => {}} onClose={() => {}} onResolve={() => {}} {...props} />);
}

describe('ReviewPanel', () => {
  it('says so when there are no changes, and still offers the views', () => {
    const html = render();
    expect(html).toContain('No tracked changes.');
    expect(html).toContain('All markup');
    expect(html).toContain('Final');
    expect(html).toContain('Original');
    expect(html).not.toContain('data-testid="review-change"');
  });

  it('lists each change with its author, what it did and the text, in the colour of its author', () => {
    const changes = [
      change({ id: 'a', author: 'Ada', date: '2024-05-06T07:08:09Z' }),
      change({ id: 'b', kind: 'deletion', author: 'Bo', text: 'old\u200B text' }),
      change({ id: 'c', kind: 'move', author: 'Cy' }),
      change({ id: 'd', kind: 'rowDeleted', author: 'Bo', text: '' }),
    ];
    const html = render({ changes });
    expect((html.match(/data-testid="review-change"/g) ?? []).length).toBe(4);
    expect(html).toContain('data-kind="deletion"');
    expect(html).toContain('Inserted');
    expect(html).toContain('Deleted');
    expect(html).toContain('Moved');
    expect(html).toContain('Deleted table row');
    expect(html).toContain('“abc”');
    expect(html).toContain('“old text”');
    const colors = assignAuthorColors(['Ada', 'Bo', 'Cy']);
    expect(html.toLowerCase()).toContain(`border-left:3px solid ${colors.get('Ada')!.toLowerCase()}`);
    // Authors are the options of the author filter.
    expect(html).toContain('<option value="Ada">Ada</option>');
    expect(html).toContain('<option value="Bo">Bo</option>');
  });

  it('shortens a long text and names every kind of change', () => {
    const kinds: ReviewChange['kind'][] = ['insertion', 'deletion', 'move', 'format', 'paragraphInserted', 'paragraphDeleted', 'paragraphFormat', 'rowInserted', 'rowDeleted'];
    const html = render({ changes: kinds.map((kind, i) => change({ id: String(i), kind, text: 'x'.repeat(200) })) });
    expect(html).toContain(`“${'x'.repeat(80)}…”`);
    for (const label of ['Formatted', 'Inserted paragraph break', 'Deleted paragraph break', 'Changed paragraph formatting', 'Inserted table row']) expect(html).toContain(label);
  });

  it('offers Accept and Reject for each change and for all, and neither when it cannot change the document', () => {
    const changes = [change({ id: 'a' })];
    const editable = render({ changes });
    expect(editable).toContain('aria-label="Accept change"');
    expect(editable).toContain('aria-label="Reject change"');
    expect(editable).toContain('Accept all');
    expect(editable).toContain('Reject all');
    const readOnly = render({ changes, onResolve: undefined });
    expect(readOnly).not.toContain('Accept');
    expect(readOnly).not.toContain('Reject');
    expect(readOnly).toContain('data-testid="review-change"');
  });

  it('marks the view that is showing, and disables going through the changes when there are none', () => {
    const html = render({ view: 'final', changes: [] });
    expect(html).toMatch(/aria-pressed="true"[^>]*>Final</);
    expect(html).toMatch(/aria-pressed="false"[^>]*>All markup</);
    expect(html).toMatch(/aria-label="Next change" disabled/);
    expect(render({ changes: [change({})] })).not.toMatch(/aria-label="Next change" disabled/);
  });
});

describe('neighbourChange', () => {
  const list = [change({ id: 'a', from: 5 }), change({ id: 'b', from: 20 }), change({ id: 'c', from: 40 })];
  const id = (c: ReviewChange | undefined) => c?.id;

  it('goes to the first change from nowhere, forwards, and to the last, backwards', () => {
    expect(id(neighbourChange(list, null, 'next'))).toBe('a');
    expect(id(neighbourChange(list, null, 'previous'))).toBe('c');
  });

  it('goes on from a position, wrapping around at the ends', () => {
    expect(id(neighbourChange(list, 5, 'next'))).toBe('b');
    expect(id(neighbourChange(list, 40, 'next'))).toBe('a');
    expect(id(neighbourChange(list, 20, 'previous'))).toBe('a');
    expect(id(neighbourChange(list, 5, 'previous'))).toBe('c');
  });

  it('goes on from where it was even when the document changed since (ids moved, positions did not)', () => {
    expect(id(neighbourChange([change({ id: 'x', from: 6 }), change({ id: 'y', from: 22 })], 5, 'next'))).toBe('x');
    expect(id(neighbourChange([change({ id: 'x', from: 6 }), change({ id: 'y', from: 22 })], 10, 'next'))).toBe('y');
  });

  it('is nothing without changes', () => {
    expect(neighbourChange([], 3, 'next')).toBeUndefined();
    expect(neighbourChange([], null, 'previous')).toBeUndefined();
  });
});

describe('ReviewPanel rows', () => {
  it('can be reached by keyboard', () => {
    expect(render({ changes: [change({ id: 'a' })] })).toContain('tabindex="0"');
  });
});
