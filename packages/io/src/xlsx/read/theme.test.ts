import { DEFAULT_THEME_COLORS } from './colors';
import { el } from './test-support';
import { parseTheme } from './theme';

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

describe('parseTheme', () => {
  it('lists the colours in a cell colour\'s theme-index order (lt1, dk1, lt2, dk2, accents, hlink, folHlink), whatever the order in the part', () => {
    const theme = parseTheme(
      el(`<a:theme ${A}><a:themeElements><a:clrScheme name="x">
        <a:dk1><a:sysClr val="windowText" lastClr="101010"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FEFEFE"/></a:lt1>
        <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
        <a:accent1><a:srgbClr val="111111"/></a:accent1><a:accent2><a:srgbClr val="222222"/></a:accent2><a:accent3><a:srgbClr val="333333"/></a:accent3>
        <a:accent4><a:srgbClr val="444444"/></a:accent4><a:accent5><a:srgbClr val="555555"/></a:accent5><a:accent6><a:srgbClr val="666666"/></a:accent6>
        <a:hlink><a:srgbClr val="0000EE"/></a:hlink><a:folHlink><a:srgbClr val="551A8B"/></a:folHlink>
      </a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Cambria"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`),
    );
    expect(theme.colors).toEqual(['#fefefe', '#101010', '#e7e6e6', '#44546a', '#111111', '#222222', '#333333', '#444444', '#555555', '#666666', '#0000ee', '#551a8b']);
    expect(theme.majorFont).toBe('Cambria');
    expect(theme.minorFont).toBe('Calibri');
  });
  it('falls back per slot: system colour without lastClr, unknown colour element, bad value, missing slot', () => {
    const theme = parseTheme(
      el(`<a:theme ${A}><a:themeElements><a:clrScheme name="x">
        <a:dk1><a:sysClr val="windowText"/></a:dk1><a:lt1><a:sysClr val="window"/></a:lt1>
        <a:dk2><a:sysClr val="other"/></a:dk2><a:lt2><a:prstClr val="red"/></a:lt2>
        <a:accent1><a:srgbClr val="nope"/></a:accent1><a:accent2/><a:accent3><a:sysClr/></a:accent3>
      </a:clrScheme></a:themeElements></a:theme>`),
    );
    expect(theme.colors).toEqual([
      '#ffffff', '#000000', DEFAULT_THEME_COLORS[2], DEFAULT_THEME_COLORS[3], DEFAULT_THEME_COLORS[4], DEFAULT_THEME_COLORS[5],
      ...DEFAULT_THEME_COLORS.slice(6),
    ]);
    expect(theme.majorFont).toBeUndefined();
    expect(theme.minorFont).toBeUndefined();
  });
  it('is the Office theme for a missing part', () => expect(parseTheme(undefined).colors).toEqual([...DEFAULT_THEME_COLORS]));
});
