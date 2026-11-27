import {
  countPresentation,
  diffField,
  normalizePresentation,
  scrubMessage,
} from '../../../../scripts/pptx-fidelity/normalize';

describe('normalizePresentation', () => {
  it('digests binaries, converts dates, maps and sets, and drops undefined properties', () => {
    const out = normalizePresentation({
      bytes: new Uint8Array([1, 2, 3]),
      buffer: new Uint8Array([1, 2, 3]).buffer,
      when: new Date(Date.UTC(2020, 0, 2)),
      bad: new Date(NaN),
      map: new Map([['a', 1]]),
      set: new Set([1, 2]),
      list: [undefined, null, 1],
      skipped: undefined,
      inf: Infinity,
      nan: NaN,
      neg: -Infinity,
      text: 'x',
      flag: false,
    });
    const digest = { bytes: 3, sha256: '039058c6f2c0cb492c533b0a4d14ef77cc0f78abccced5287d84a1a2011cfb81' };
    expect(out).toEqual({
      bytes: digest,
      buffer: digest,
      when: '2020-01-02T00:00:00.000Z',
      bad: 'Invalid Date',
      map: [['a', 1]],
      set: [1, 2],
      list: [null, null, 1],
      inf: '#number:Infinity',
      nan: '#number:NaN',
      neg: '#number:-Infinity',
      text: 'x',
      flag: false,
    });
  });

  it('digests an embedded image (a base64 data URL) and leaves other strings alone', () => {
    const bytes = Buffer.from([1, 2, 3]);
    expect(normalizePresentation({ dataUrl: `data:image/png;base64,${bytes.toString('base64')}`, plain: 'data:text/plain,hi', text: 'data:' })).toEqual({
      dataUrl: { dataUrl: 'image/png', bytes: 3, sha256: '039058c6f2c0cb492c533b0a4d14ef77cc0f78abccced5287d84a1a2011cfb81' },
      plain: 'data:text/plain,hi',
      text: 'data:',
    });
  });

  it('turns null and undefined into null', () => {
    expect(normalizePresentation(undefined)).toBeNull();
    expect(normalizePresentation(null)).toBeNull();
  });
});

describe('countPresentation', () => {
  it('counts slides and their elements, and is zero for anything else', () => {
    expect(countPresentation({ slides: [{ elements: [1, 2] }, { elements: [] }, {}, 'x'] })).toEqual({ slides: 4, elements: 2 });
    expect(countPresentation({ slides: 'no' })).toEqual({ slides: 0, elements: 0 });
    expect(countPresentation(null)).toEqual({ slides: 0, elements: 0 });
    expect(countPresentation([])).toEqual({ slides: 0, elements: 0 });
  });
});

describe('diffField and scrubMessage', () => {
  it('folds indexes out of a path', () => {
    expect(diffField('slides[3].elements[12].text')).toBe('slides[].elements[].text');
  });

  it('replaces the longest root first and ignores empty roots', () => {
    expect(scrubMessage('at /a/b/c and /a/x', ['/a', '/a/b', ''])).toBe('at <repo>/c and <repo>/x');
  });
});
