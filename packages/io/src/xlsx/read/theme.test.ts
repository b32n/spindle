import { DEFAULT_THEME_COLORS, hlsToRgb } from './colors';
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
  const warnings: string[] = [];
  it('falls back per slot: system colour without lastClr, unknown colour element, bad value, missing slot', () => {
    const theme = parseTheme(
      el(`<a:theme ${A}><a:themeElements><a:clrScheme name="x">
        <a:dk1><a:sysClr val="windowText"/></a:dk1><a:lt1><a:sysClr val="window"/></a:lt1>
        <a:dk2><a:sysClr val="other"/></a:dk2><a:lt2><a:prstClr val="nothing"/></a:lt2>
        <a:accent1><a:srgbClr val="nope"/></a:accent1><a:accent2/><a:accent3><a:sysClr/></a:accent3>
        <a:accent4><a:schemeClr val="accent1"/></a:accent4><a:accent5><a:scrgbClr r="1" g="2"/></a:accent5><a:accent6><a:hslClr hue="x" sat="1" lum="1"/></a:accent6>
      </a:clrScheme></a:themeElements></a:theme>`),
      warnings,
    );
    expect(theme.colors).toEqual([
      '#ffffff', '#000000', DEFAULT_THEME_COLORS[2], DEFAULT_THEME_COLORS[3], DEFAULT_THEME_COLORS[4], DEFAULT_THEME_COLORS[5],
      ...DEFAULT_THEME_COLORS.slice(6),
    ]);
    expect(warnings).toEqual([
      'The theme colour lt2 (prstClr) could not be read; the Office default is used.',
      'The theme colour dk2 (sysClr) could not be read; the Office default is used.',
      'The theme colour accent1 (srgbClr) could not be read; the Office default is used.',
      'The theme colour accent3 (sysClr) could not be read; the Office default is used.',
      'The theme colour accent4 (schemeClr) could not be read; the Office default is used.',
      'The theme colour accent5 (scrgbClr) could not be read; the Office default is used.',
      'The theme colour accent6 (hslClr) could not be read; the Office default is used.',
    ]);
    expect(theme.majorFont).toBeUndefined();
    expect(theme.minorFont).toBeUndefined();
  });
  it('is the Office theme for a missing part', () => expect(parseTheme(undefined).colors).toEqual([...DEFAULT_THEME_COLORS]));
  it('resolves prstClr (incl. dk/lt/grey spellings), scrgbClr (linear light) and hslClr slots', () => {
    const slot = (name: string, inner: string) => `<a:${name}>${inner}</a:${name}>`;
    const warnings: string[] = [];
    const theme = parseTheme(
      el(`<a:theme ${A}><a:themeElements><a:clrScheme name="x">${[
        slot('dk1', '<a:prstClr val="darkBlue"/>'),
        slot('lt1', '<a:prstClr val="ltGrey"/>'),
        slot('dk2', '<a:scrgbClr r="100000" g="0" b="0"/>'),
        slot('lt2', '<a:scrgbClr r="0" g="100000" b="0.5"/>'),
        slot('accent1', '<a:hslClr hue="0" sat="100000" lum="50000"/>'),
        slot('accent2', '<a:hslClr hue="7200000" sat="100000" lum="50000"/>'),
        slot('accent3', '<a:hslClr hue="-14400000" sat="100000" lum="50000"/>'),
        slot('accent4', '<a:hslClr hue="0" sat="0" lum="50000"/>'),
        slot('accent5', '<a:scrgbClr r="50000" g="50000" b="50000"/>'),
        slot('accent6', '<a:scrgbClr r="0.1" g="0.1" b="0.1"/>'),
      ].join('')}</a:clrScheme></a:themeElements></a:theme>`),
      warnings,
    );
    // hue 7200000 is 120 degrees (green); a negative hue wraps (-240 degrees is 120 as well)
    expect(theme.colors.slice(0, 10)).toEqual(['#d3d3d3', '#00008b', '#00ff00', '#ff0000', '#ff0000', '#00ff00', '#00ff00', '#808080', '#bcbcbc', '#000000']);
    expect(warnings).toEqual([]);
  });
  it('hlsToRgb is the HLS to RGB conversion the tint uses', () => expect(hlsToRgb(0, 0.5, 1)).toBe('#ff0000'));
  it.each(['constructor', '__proto__', 'toString', 'hasOwnProperty', 'valueOf'])('inherited object key %s as a prstClr or sysClr name never leaks into the colours', (key) => {
    const warnings: string[] = [];
    const theme = parseTheme(
      el(`<a:theme ${A}><a:themeElements><a:clrScheme name="x"><a:lt1><a:prstClr val="${key}"/></a:lt1><a:dk1><a:sysClr val="${key}"/></a:dk1></a:clrScheme></a:themeElements></a:theme>`),
      warnings,
    );
    expect(theme.colors).toEqual([...DEFAULT_THEME_COLORS]);
    expect(theme.colors.every((c) => typeof c === 'string')).toBe(true);
    expect(warnings).toEqual([
      'The theme colour lt1 (prstClr) could not be read; the Office default is used.',
      'The theme colour dk1 (sysClr) could not be read; the Office default is used.',
    ]);
    expect(structuredClone(theme)).toEqual(theme);
  });
  it('a prstClr without a name is unreadable', () => {
    const warnings: string[] = [];
    parseTheme(el(`<a:theme ${A}><a:themeElements><a:clrScheme name="x"><a:dk1><a:prstClr/></a:dk1></a:clrScheme></a:themeElements></a:theme>`), warnings);
    expect(warnings).toEqual(['The theme colour dk1 (prstClr) could not be read; the Office default is used.']);
  });
});
