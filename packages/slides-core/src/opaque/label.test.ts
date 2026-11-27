import { opaqueLabel } from './label';

describe('opaqueLabel', () => {
  it('names the kind, adds the title when there is one, and says it is kept but not editable', () => {
    expect(opaqueLabel({ kind: 'chart', title: 'Sales 2024' })).toBe('Chart: Sales 2024 (kept, not editable)');
    expect(opaqueLabel({ kind: 'chart' })).toBe('Chart (kept, not editable)');
    expect(opaqueLabel({ kind: 'smartArt' })).toBe('SmartArt diagram (kept, not editable)');
    expect(opaqueLabel({ kind: 'ole' })).toBe('Embedded object (kept, not editable)');
    expect(opaqueLabel({ kind: 'ink' })).toBe('Ink drawing (kept, not editable)');
  });

  it('tells a video from an audio clip by the clip element, and falls back to a media clip', () => {
    expect(opaqueLabel({ kind: 'media', props: { clip: 'videoFile' } })).toBe('Video (kept, not editable)');
    expect(opaqueLabel({ kind: 'media', props: { clip: 'quickTimeFile' } })).toBe('Video (kept, not editable)');
    expect(opaqueLabel({ kind: 'media', props: { clip: 'audioFile' } })).toBe('Audio (kept, not editable)');
    expect(opaqueLabel({ kind: 'media', props: { clip: 'other' } })).toBe('Media clip (kept, not editable)');
    expect(opaqueLabel({ kind: 'media' })).toBe('Media clip (kept, not editable)');
  });
});
