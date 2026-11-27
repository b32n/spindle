import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { createOpaqueElement, type OpaqueElement } from '@b32nio/spindle-slides-core';
import { OpaqueView } from './elements/OpaqueView';

const el = (extra: Partial<OpaqueElement> = {}): OpaqueElement => ({ ...createOpaqueElement({ containerId: 's', index: 'a', kind: 'chart', partRefs: [] }), ...extra });
const html = (e: OpaqueElement): string => renderToStaticMarkup(<OpaqueView el={e} />);

describe('OpaqueView', () => {
  it('draws a labelled placeholder with the chart title and says what the element is for assistive technology', () => {
    const h = html(el({ title: 'Sales 2024' }));
    expect(h).toContain('Chart: Sales 2024 (kept, not editable)');
    expect(h).toContain('role="img"');
    expect(h).toContain('aria-label="Chart: Sales 2024 (kept, not editable)"');
    expect(h).toContain('data-opaque-kind="chart"');
    expect(h).not.toContain('<img');
  });

  it('adds the source\'s alternative text to the accessible name, and shows a poster behind the label', () => {
    const h = html(el({ kind: 'media', props: { clip: 'videoFile' }, alt: 'Intro clip', poster: 'data:image/png;base64,AA==' }));
    expect(h).toContain('aria-label="Video (kept, not editable). Intro clip"');
    expect(h).toContain('<img src="data:image/png;base64,AA==" alt=""');
    expect(h).toContain('>Video (kept, not editable)<');
  });
});
