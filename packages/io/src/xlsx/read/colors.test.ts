import { applyTint, DEFAULT_INDEXED_COLORS, DEFAULT_THEME_COLORS, parseRgb, readColor, resolveColor } from './colors';
import { el } from './test-support';

const ctx = { themeColors: DEFAULT_THEME_COLORS };

describe('parseRgb', () => {
  it.each([
    ['FF336699', '#336699'],
    ['00336699', '#336699'],
    ['AbCdEf', '#abcdef'],
  ])('%s -> %s', (input, expected) => expect(parseRgb(input)).toBe(expected));
  it.each([undefined, '', 'red', '12345', '1234567', 'GG0000', '#336699'])('%s is not a colour', (input) => expect(parseRgb(input)).toBeUndefined());
});

describe('applyTint (oracle: the colours Excel shows for the Office 2013 theme; Excel computes in integer HLS, this in floating point, so channels may differ by 1 or 2)', () => {
  // [base, tint, expected] from Excel's colour picker; the tints are the values Excel writes for "Lighter 80/60/40%" and "Darker 25/50%"
  const oracle: Array<[string, number, string]> = [
    ['#4472c4', 0.7999816888943144, '#d9e1f2'],
    ['#4472c4', 0.5999938962981048, '#b4c6e7'],
    ['#4472c4', 0.3999755859375, '#8eaadb'],
    ['#4472c4', -0.249977111117893, '#2f5597'],
    ['#4472c4', -0.499984740745262, '#203864'],
    ['#ed7d31', 0.7999816888943144, '#fbe5d6'],
    ['#ed7d31', 0.5999938962981048, '#f8cbad'],
    ['#ed7d31', 0.3999755859375, '#f4b183'],
    ['#ed7d31', -0.249977111117893, '#c55a11'],
    ['#ed7d31', -0.499984740745262, '#843c0c'],
    ['#a5a5a5', 0.7999816888943144, '#ededed'],
    ['#a5a5a5', 0.5999938962981048, '#dbdbdb'],
    ['#a5a5a5', 0.3999755859375, '#c9c9c9'],
    ['#a5a5a5', -0.249977111117893, '#7c7c7c'],
    ['#a5a5a5', -0.499984740745262, '#525252'],
    ['#ffffff', -0.05, '#f2f2f2'],
    ['#ffffff', -0.15, '#d9d9d9'],
    ['#000000', 0.5, '#808080'],
  ];
  it.each(oracle)('%s with tint %s is %s (within 2 per channel)', (base, tint, expected) => {
    const got = applyTint(base, tint);
    for (const at of [1, 3, 5]) expect(Math.abs(parseInt(got.slice(at, at + 2), 16) - parseInt(expected.slice(at, at + 2), 16))).toBeLessThanOrEqual(2);
  });
  it('leaves the colour alone at tint 0, whites it at +1 and blacks it at -1, and clamps beyond', () => {
    expect(applyTint('#4472c4', 0)).toBe('#4472c4');
    expect(applyTint('#4472c4', 1)).toBe('#ffffff');
    expect(applyTint('#4472c4', -1)).toBe('#000000');
    expect(applyTint('#4472c4', 7)).toBe('#ffffff');
    expect(applyTint('#4472c4', -7)).toBe('#000000');
  });
  it('keeps greys grey and covers every hue sector', () => {
    expect(applyTint('#808080', 0.5)).toBe('#c0c0c0');
    // primaries and secondaries exercise each branch of the hue conversion
    for (const hex of ['#ff0000', '#ffff00', '#00ff00', '#00ffff', '#0000ff', '#ff00ff', '#ff0080', '#10a030', '#3010a0'])
      expect(applyTint(hex, 0.25)).toMatch(/^#[0-9a-f]{6}$/);
    expect(applyTint('#ff0000', -0.5)).toBe('#800000');
    expect(applyTint('#00ff00', -0.5)).toBe('#008000');
    expect(applyTint('#0000ff', -0.5)).toBe('#000080');
    // a bright, saturated colour lightened: the second lightness branch of hls -> rgb
    expect(applyTint('#ff0000', 0.5)).toBe('#ff8080');
  });
});

describe('resolveColor', () => {
  it('reads rgb (alpha ignored), theme (Excel order), indexed, the system indexes, and applies tint', () => {
    expect(resolveColor({ rgb: 'FF112233' }, ctx)).toBe('#112233');
    expect(resolveColor({ theme: 0 }, ctx)).toBe('#ffffff'); // lt1
    expect(resolveColor({ theme: 1 }, ctx)).toBe('#000000'); // dk1
    expect(resolveColor({ theme: 4 }, ctx)).toBe('#4472c4');
    expect(resolveColor({ theme: 11 }, ctx)).toBe('#954f72');
    expect(resolveColor({ indexed: 2 }, ctx)).toBe('#ff0000');
    expect(resolveColor({ indexed: 64 }, ctx)).toBe('#000000');
    expect(resolveColor({ indexed: 65 }, ctx)).toBe('#ffffff');
    expect(resolveColor({ theme: 4, tint: 0.5 }, ctx)).toBe(applyTint('#4472c4', 0.5));
  });
  it('uses the workbook palette over the default one', () => {
    expect(resolveColor({ indexed: 2 }, { ...ctx, indexedColors: ['#000000', '#ffffff', '#123456'] })).toBe('#123456');
    expect(resolveColor({ indexed: 3 }, { ...ctx, indexedColors: ['#000000'] })).toBeUndefined();
  });
  it('gives nothing for auto, an unreadable rgb, an unknown index, or an empty colour', () => {
    expect(resolveColor({ auto: true }, ctx)).toBeUndefined();
    expect(resolveColor({ rgb: 'zz' }, ctx)).toBeUndefined();
    expect(resolveColor({ theme: 99 }, ctx)).toBeUndefined();
    expect(resolveColor({ indexed: 200 }, ctx)).toBeUndefined();
    expect(resolveColor({}, ctx)).toBeUndefined();
  });
});

describe('default tables', () => {
  it('have 64 palette entries and 12 theme colours', () => {
    expect(DEFAULT_INDEXED_COLORS).toHaveLength(64);
    expect(DEFAULT_THEME_COLORS).toHaveLength(12);
    expect(DEFAULT_INDEXED_COLORS[10]).toBe('#ff0000');
  });
});

describe('readColor', () => {
  it('keeps what the file wrote and adds the resolved value', () => {
    expect(readColor(el('<color theme="4" tint="-0.25"/>'), ctx)).toEqual({ theme: 4, tint: -0.25, resolved: applyTint('#4472c4', -0.25) });
    expect(readColor(el('<fgColor rgb="FFC6EFCE"/>'), ctx)).toEqual({ rgb: 'FFC6EFCE', resolved: '#c6efce' });
    expect(readColor(el('<color indexed="13"/>'), ctx)).toEqual({ indexed: 13, resolved: '#ffff00' });
    expect(readColor(el('<color auto="1"/>'), ctx)).toEqual({ auto: true });
    expect(readColor(el('<color auto="true"/>'), ctx)).toEqual({ auto: true });
    expect(readColor(el('<color auto="0"/>'), ctx)).toEqual({});
  });
  it('ignores attributes that are not numbers and gives undefined for no element', () => {
    expect(readColor(el('<color theme="x" tint="" indexed="1.5e400"/>'), ctx)).toEqual({});
    expect(readColor(undefined, ctx)).toBeUndefined();
  });
});
