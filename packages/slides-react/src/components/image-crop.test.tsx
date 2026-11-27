import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import type { ImageElement } from '@b32nio/spindle-slides-core';
import { ImageView } from './elements/ImageView';
import { ImageCropFields } from './ImageCropFields';

const PNG = 'data:image/png;base64,iVBORw0KGgo=';
const el = (extra: Partial<ImageElement> = {}): ImageElement =>
  ({ id: 'i', type: 'image', containerId: 's', index: 'a', x: 0, y: 0, w: 200, h: 100, rotation: 0, src: PNG, naturalW: 100, naturalH: 100, ...extra }) as ImageElement;
const html = (e: ImageElement) => renderToStaticMarkup(<ImageView el={e} />);

describe('ImageView', () => {
  it('an uncropped image is a single img with its fit (unchanged markup)', () => {
    const h = html(el({ fit: 'cover' }));
    expect(h.startsWith('<img')).toBe(true);
    expect(h).toContain('object-fit:cover');
    expect(h).not.toContain('data-image-crop');
    expect(html(el({ crop: { l: 0, t: 0, r: 0, b: 0 } }))).toContain('object-fit:fill');
    expect(html(el())).toContain('object-fit:fill');
  });
  it('an uncropped image keeps flips on the img', () => {
    expect(html(el({ flipH: true }))).toContain('scale(-1, 1)');
    expect(html(el({ flipV: true }))).toContain('scale(1, -1)');
  });
  it('a cropped image is clipped by the box and the region; the crop is a DOM attribute', () => {
    const h = html(el({ crop: { l: 10, t: 20, r: 30, b: 0 } }));
    expect(h).toContain('data-image-crop="10,20,30,0"');
    expect(h).toContain('overflow:hidden');
    expect(h).toContain('data-image-crop-region');
    expect(h).toContain('max-width:none');
    expect(h).toContain('width:166.666');
    expect(h).toContain('height:125%');
    expect(h).toContain('left:-16.666');
    expect(h).toContain('top:-25%');
  });
  it('a cropped image honours fit and flips', () => {
    const h = html(el({ crop: { l: 0, t: 0, r: 50, b: 0 }, fit: 'contain', flipH: true, flipV: true }));
    // visible 50x100 (aspect 0.5) contained in 200x100: region 25% wide, centred
    expect(h).toContain('left:37.5%');
    expect(h).toContain('scale(-1, -1)');
  });
  it('an unusable crop is ignored', () => {
    expect(html(el({ crop: { l: NaN } as never }))).not.toContain('data-image-crop');
  });
});

describe('ImageCropFields', () => {
  type P = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
  const find = (node: React.ReactNode, type: string, out: P[] = []): P[] => {
    React.Children.forEach(node, (c) => {
      if (!React.isValidElement(c)) return;
      const p = c.props as P;
      if (c.type === type) out.push(p);
      find(p.children, type, out);
    });
    return out;
  };

  it('shows the four edges (zero when uncropped) with labels, and a reset', () => {
    const tree = ImageCropFields({ crop: undefined, onChange: jest.fn() });
    const inputs = find(tree, 'input');
    expect(inputs.map((i) => i['data-testid'])).toEqual(['image-crop-l', 'image-crop-t', 'image-crop-r', 'image-crop-b']);
    expect(inputs.map((i) => i.value)).toEqual([0, 0, 0, 0]);
    // each input is named by exactly one visible label (no aria-label overriding it)
    expect(inputs.every((i) => i['aria-label'] === undefined)).toBe(true);
    expect(find(tree, 'label').map((l) => [l.htmlFor, l.children])).toEqual([
      ['crop-l', ['Left', ' (%)']], ['crop-t', ['Top', ' (%)']], ['crop-r', ['Right', ' (%)']], ['crop-b', ['Bottom', ' (%)']],
    ]);
    expect(inputs.map((i) => i.id)).toEqual(['crop-l', 'crop-t', 'crop-r', 'crop-b']);
    expect(find(tree, 'button')[0]['aria-label']).toBe('Reset crop');
    expect(renderToStaticMarkup(tree)).toContain('role="group"');
  });
  it('idPrefix namespaces the input ids and label targets', () => {
    const tree = ImageCropFields({ crop: undefined, onChange: jest.fn(), idPrefix: 'crop-x' });
    expect(find(tree, 'input').map((i) => i.id)).toEqual(['crop-x-l', 'crop-x-t', 'crop-x-r', 'crop-x-b']);
    expect(find(tree, 'label').map((l) => l.htmlFor)).toEqual(['crop-x-l', 'crop-x-t', 'crop-x-r', 'crop-x-b']);
  });
  it('an edit changes one edge and keeps the rest; limits apply; reset zeroes', () => {
    const onChange = jest.fn();
    const crop = { l: 10, t: 0, r: 20, b: 0 };
    const tree = ImageCropFields({ crop, onChange });
    const [l, , r, b] = find(tree, 'input');
    expect(l.value).toBe(10);
    r.onChange({ target: { valueAsNumber: 95 } });
    expect(onChange).toHaveBeenLastCalledWith({ l: 10, t: 0, r: 89, b: 0 });
    b.onChange({ target: { valueAsNumber: NaN } });
    expect(onChange).toHaveBeenLastCalledWith({ l: 10, t: 0, r: 20, b: 0 });
    find(tree, 'button')[0].onClick();
    expect(onChange).toHaveBeenLastCalledWith({ l: 0, t: 0, r: 0, b: 0 });
  });
});
