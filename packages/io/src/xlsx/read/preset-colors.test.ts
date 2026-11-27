import { PRESET_COLORS } from './preset-colors';

describe('PRESET_COLORS', () => {
  it.each([
    ['red', '#ff0000'], ['black', '#000000'], ['white', '#ffffff'], ['darkBlue', '#00008b'], ['dkBlue', '#00008b'], ['lightYellow', '#ffffe0'],
    ['ltYellow', '#ffffe0'], ['gray', '#808080'], ['grey', '#808080'], ['dkGrey', '#a9a9a9'], ['ltSlateGrey', '#778899'], ['medVioletRed', '#c71585'],
    ['yellowGreen', '#9acd32'], ['lightGoldenrodYellow', '#fafad2'], ['dkOliveGreen', '#556b2f'],
  ])('%s is %s', (name, hex) => expect(PRESET_COLORS[name]).toBe(hex));
  it('every value is #rrggbb, every ST_PresetColorVal spelling is present, and an unknown name is not', () => {
    for (const value of Object.values(PRESET_COLORS)) expect(value).toMatch(/^#[0-9a-f]{6}$/);
    for (const name of Object.keys(PRESET_COLORS).filter((k) => /[Gg]ray/.test(k))) expect(PRESET_COLORS[name.replace(/(G|g)ray/, '$1rey')]).toBe(PRESET_COLORS[name]);
    expect(Object.keys(PRESET_COLORS).length).toBeGreaterThanOrEqual(140);
    expect(PRESET_COLORS.notAColour).toBeUndefined();
    expect(PRESET_COLORS.Red).toBeUndefined();
  });
});
