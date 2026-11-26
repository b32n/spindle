import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { RevisionLegend } from './RevisionLegend';

describe('RevisionLegend', () => {
  it('shows nothing without authors', () => {
    expect(renderToStaticMarkup(<RevisionLegend authors={[]} />)).toBe('');
  });

  it('lists each author with their colour and count, as a group holding a list', () => {
    const html = renderToStaticMarkup(
      <RevisionLegend
        authors={[
          { name: 'Ada', color: '#1a73e8', count: 3 },
          { name: 'Bo', color: '#188038', count: 1 },
        ]}
      />
    );
    expect(html).toContain('role="group"');
    expect(html).toContain('Tracked changes');
    expect(html.match(/role="listitem"/g)).toHaveLength(2);
    expect(html).toContain('data-revision-author="Ada"');
    expect(html).toContain('(3)');
    expect(html).toContain('background:#188038');
    // The label sits outside the list, which holds list items only.
    expect(html).toMatch(/<span[^>]*>Tracked changes<\/span><div role="list"/);
  });

  it('has a band of its own (padding, background, border) unless it is drawn bare', () => {
    const authors = [{ name: 'Ada', color: '#1a73e8', count: 1 }];
    const band = renderToStaticMarkup(<RevisionLegend authors={authors} />);
    expect(band).toContain('padding:4px 16px');
    expect(band).toContain('border-bottom:1px solid #e8eaed');
    const bare = renderToStaticMarkup(<RevisionLegend authors={authors} bare />);
    expect(bare).toContain('padding:4px 0');
    expect(bare).not.toContain('border-bottom');
    expect(bare).not.toContain('background:#ffffff');
  });
});
