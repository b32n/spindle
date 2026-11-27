import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { createLineElement, getBuiltinTheme, type LineElement } from '@b32nio/spindle-slides-core';
import { LineView } from './elements/LineView';

let selected: string[] = [];
let elements: Record<string, unknown> = {};
const updateElement = jest.fn();

jest.mock('../hooks', () => ({
  ...jest.requireActual('../hooks'),
  useDeck: () => ({ getElement: (id: string) => elements[id], updateElement, getActiveSlideId: () => 's1' }),
  useSelection: () => ({ slideId: 's1', elementIds: selected }),
  useClipboard: () => ({ hasContent: () => false }),
}));

import { ContextMenu } from './ContextMenu';

const line = (over: Partial<LineElement> = {}): LineElement => ({ ...createLineElement({ containerId: 's1', index: 'a0', x: 0, y: 0, w: 100, h: 50, ...over }), id: 'l' });
const theme = getBuiltinTheme('default');

describe('LineView connector routing', () => {
  it('a plain line stays a <line>', () => {
    const html = renderToStaticMarkup(<LineView el={line()} theme={theme} />);
    expect(html).toContain('<line');
    expect(html).not.toContain('<path d=');
  });

  it('an elbow draws right-angle legs through the endpoints (padded by the SVG offset)', () => {
    const html = renderToStaticMarkup(<LineView el={line({ connector: 'elbow', endArrow: 'triangle' })} theme={theme} endpoints={{ x1: 0, y1: 0, x2: 100, y2: 50 }} dirs={{ start: { x: 1, y: 0 }, end: { x: -1, y: 0 } }} />);
    expect(html).toContain('d="M0 12 L50 12 L50 62 L100 62"');
    expect(html).toContain('marker-end="url(#arw-e-l)"');
    expect(html).not.toContain('<line');
  });

  it('a curved connector draws a cubic, and a free elbow without dirs is still routed', () => {
    expect(renderToStaticMarkup(<LineView el={line({ connector: 'curved' })} theme={theme} dirs={{ start: { x: 1, y: 0 }, end: { x: -1, y: 0 } }} />)).toContain('d="M0 12 C50 12 50 62 100 62"');
    expect(renderToStaticMarkup(<LineView el={line({ connector: 'elbow', adjustments: [0.25] })} theme={theme} />)).toContain('d="M0 12 L25 12 L25 62 L100 62"');
  });
});

describe('ContextMenu connector routing', () => {
  beforeEach(() => { updateElement.mockClear(); });

  it('offers straight, elbow and curved for a single selected line, marking the current one', () => {
    elements = { l: line({ connector: 'elbow' }) };
    selected = ['l'];
    const html = renderToStaticMarkup(<ContextMenu x={0} y={0} onClose={() => undefined} />);
    expect(html).toContain('Straight connector<');
    expect(html).toContain('Elbow connector ✓<');
    expect(html).toContain('Curved connector<');
  });

  it('marks straight for a plain line and offers nothing for other selections', () => {
    elements = { l: line(), t: { id: 't', type: 'text' } };
    selected = ['l'];
    expect(renderToStaticMarkup(<ContextMenu x={0} y={0} onClose={() => undefined} />)).toContain('Straight connector ✓<');
    selected = ['t'];
    expect(renderToStaticMarkup(<ContextMenu x={0} y={0} onClose={() => undefined} />)).not.toContain('connector');
    selected = ['l', 't'];
    expect(renderToStaticMarkup(<ContextMenu x={0} y={0} onClose={() => undefined} />)).not.toContain('connector');
  });

  it('choosing a routing updates the line and resets the elbow position', () => {
    elements = { l: line() };
    selected = ['l'];
    const tree = ContextMenu({ x: 0, y: 0, onClose: () => undefined }) as React.ReactElement<{ items: Array<{ label: string; run: () => void } | 'sep'> }>;
    const items = tree.props.items.filter((i): i is { label: string; run: () => void } => i !== 'sep');
    items.find((i) => i.label.startsWith('Elbow'))!.run();
    expect(updateElement).toHaveBeenLastCalledWith('l', { connector: 'elbow', adjustments: undefined });
    items.find((i) => i.label.startsWith('Straight'))!.run();
    expect(updateElement).toHaveBeenLastCalledWith('l', { connector: undefined, adjustments: undefined });
  });
});
