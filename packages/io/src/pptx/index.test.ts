import * as pptx from './index';
import * as io from '../index';

describe('@b32nio/spindle-io/pptx', () => {
  it('exports the loss detector, the reader and the writer, and the package root re-exports them', () => {
    expect(Object.keys(pptx).sort()).toEqual(['ImportError', 'detectPptxLoss', 'exportPptx', 'importPptx', 'lintPptxPackage']);
    expect(io.detectPptxLoss).toBe(pptx.detectPptxLoss);
    expect(io.importPptx).toBe(pptx.importPptx);
    expect(io.exportPptx).toBe(pptx.exportPptx);
  });

  it('reports the pptx format', async () => {
    expect(await pptx.detectPptxLoss(new Uint8Array(0))).toEqual({ format: 'pptx', items: [] });
  });

  it('rejects what is not a presentation with an ImportError', async () => {
    await expect(pptx.importPptx(new Uint8Array(0))).rejects.toBeInstanceOf(pptx.ImportError);
  });
});
