import * as pptx from './index';
import * as io from '../index';

describe('@b32nio/spindle-io/pptx', () => {
  it('exports the loss detector and the reader, and the package root re-exports them', () => {
    expect(Object.keys(pptx).sort()).toEqual(['ImportError', 'detectPptxLoss', 'importPptx']);
    expect(io.detectPptxLoss).toBe(pptx.detectPptxLoss);
    expect(io.importPptx).toBe(pptx.importPptx);
  });

  it('reports the pptx format', async () => {
    expect(await pptx.detectPptxLoss(new Uint8Array(0))).toEqual({ format: 'pptx', items: [] });
  });

  it('rejects what is not a presentation with an ImportError', async () => {
    await expect(pptx.importPptx(new Uint8Array(0))).rejects.toBeInstanceOf(pptx.ImportError);
  });
});
