import * as pptx from './index';
import * as io from '../index';

describe('@b32nio/spindle-io/pptx', () => {
  it('exports the loss detector, and the package root re-exports it', () => {
    expect(Object.keys(pptx)).toEqual(['detectPptxLoss']);
    expect(io.detectPptxLoss).toBe(pptx.detectPptxLoss);
  });

  it('reports the pptx format', async () => {
    expect(await pptx.detectPptxLoss(new Uint8Array(0))).toEqual({ format: 'pptx', items: [] });
  });
});
