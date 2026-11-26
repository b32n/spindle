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

  it('offers a change without an author as "Unknown", a choice of its own, and labels the author choice for each panel', () => {
    const html = render({ changes: [change({ id: 'a', author: '' }), change({ id: 'b', author: 'Ada' })] });
    expect(html).toContain('<option value="">Unknown</option>');
    expect(html).toContain('All authors');
    // The label points at the select of this panel: two panels on one page do not share an id.
    const ids = (markup: string) => [markup.match(/<label for="([^"]+)"/)?.[1], markup.match(/<select id="([^"]+)"/)?.[1]];
    const [labelFor, selectId] = ids(html);
    expect(labelFor).toBeTruthy();
    expect(labelFor).toBe(selectId);
    const both = renderToStaticMarkup(
      <>
        <ReviewPanel changes={[]} view="markup" onViewChange={() => {}} onGoTo={() => {}} onClose={() => {}} />
        <ReviewPanel changes={[]} view="markup" onViewChange={() => {}} onGoTo={() => {}} onClose={() => {}} />
      </>
    );
    const selects = [...both.matchAll(/<select id="([^"]+)"/g)].map((m) => m[1]);
    expect(selects).toHaveLength(2);
    expect(new Set(selects).size).toBe(2);
  });

  it('draws an author in the colour the page gives them when the palette lists every author of the document', () => {
    // Bo and Di share a first colour; Bo is only in the other end of a move, so no change of Bo's is in the list.
    const changes = [change({ id: 'a', author: 'Ada' }), change({ id: 'b', author: 'Di' })];
    const page = assignAuthorColors(['Ada', 'Bo', 'Di']);
    const html = render({ changes, palette: ['Ada', 'Bo', 'Di'] }).toLowerCase();
    expect(html).toContain(`border-left:3px solid ${page.get('Di')!.toLowerCase()}`);
    expect(page.get('Di')).not.toBe(assignAuthorColors(['Ada', 'Di']).get('Di'));
    expect(render({ changes }).toLowerCase()).toContain(`border-left:3px solid ${assignAuthorColors(['Ada', 'Di']).get('Di')!.toLowerCase()}`);
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

  it('visits changes that start at one position one after the other, in both directions', () => {
    const same = [change({ id: 'a', from: 5 }), change({ id: 'b', from: 5 }), change({ id: 'c', from: 5 }), change({ id: 'd', from: 30 })];
    expect(id(neighbourChange(same, 5, 'next', 'a'))).toBe('b');
    expect(id(neighbourChange(same, 5, 'next', 'b'))).toBe('c');
    expect(id(neighbourChange(same, 5, 'next', 'c'))).toBe('d');
    expect(id(neighbourChange(same, 30, 'next', 'd'))).toBe('a');
    expect(id(neighbourChange(same, 5, 'previous', 'c'))).toBe('b');
    expect(id(neighbourChange(same, 5, 'previous', 'a'))).toBe('d');
  });

  it('falls back to the position when the current id is not listed any more', () => {
    expect(id(neighbourChange(list, 20, 'next', 'gone'))).toBe('c');
    expect(id(neighbourChange(list, 20, 'previous', 'gone'))).toBe('a');
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
