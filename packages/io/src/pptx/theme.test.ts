import { parseXml, type XmlElement } from '../xlsx/xml';
import { FALLBACK_THEME, parseTheme } from './theme';

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
const scheme = (over: Record<string, string> = {}): string =>
  `<a:clrScheme name="c">${SLOTS.flatMap((s, i) =>
    over[s] === '' ? [] : [`<a:${s}>${over[s] ?? `<a:srgbClr val="0000${i.toString(16).padStart(2, '0')}"/>`}</a:${s}>`]
  ).join('')}</a:clrScheme>`;
const fonts = (major: string, minor: string): string =>
  `<a:fontScheme name="f"><a:majorFont>${major}</a:majorFont><a:minorFont>${minor}</a:minorFont></a:fontScheme>`;
const theme = (inner: string, attrs = 'name="Mine"'): XmlElement => parseXml(`<a:theme ${A} ${attrs}><a:themeElements>${inner}</a:themeElements></a:theme>`) as XmlElement;

describe('parseTheme', () => {
  it('reads the twelve colours, a system colour by its last value, and the font families', () => {
    const warnings: string[] = [];
    const t = parseTheme(
      theme(
        scheme({ dk1: '<a:sysClr val="windowText" lastClr="010203"/>', lt1: '<a:sysClr val="window" lastClr="FFFFFF"/>' }) +
          fonts('<a:latin typeface="Georgia"/><a:ea typeface="MS Mincho"/><a:cs typeface="Arial"/>', '<a:latin typeface=" Verdana "/><a:ea typeface=""/><a:cs typeface="Tahoma"/>')
      ),
      (m) => warnings.push(m)
    );
    expect(warnings).toEqual([]);
    expect(t.name).toBe('Mine');
    expect(t.colors.dk1).toBe('#010203');
    expect(t.colors.lt1).toBe('#FFFFFF');
    expect(t.colors.accent1).toBe('#000004');
    expect(t.fonts).toEqual({ major: 'Georgia', minor: 'Verdana', majorEastAsian: 'MS Mincho', majorComplex: 'Arial', minorComplex: 'Tahoma' });
  });

  it('keeps the Office default for a slot or a font the theme does not define, and says so once', () => {
    const warnings: string[] = [];
    const t = parseTheme(theme(scheme({ accent2: '', hlink: '<a:scrgbClr r="0" g="0" b="0"/>' }) + fonts('', '<a:latin typeface="+mn-lt"/>'), ''), (m) => warnings.push(m));
    expect(warnings).toEqual(['The theme does not define accent2, hlink; the Office default colours are used for them.']);
    expect(t.colors.accent2).toBe(FALLBACK_THEME.colors.accent2);
    expect(t.colors.hlink).toBe(FALLBACK_THEME.colors.hlink);
    expect(t.name).toBe(FALLBACK_THEME.name);
    expect(t.fonts.major).toBe('Calibri Light');
    expect(t.fonts.minor).toBe('+mn-lt');
  });

  it('is the Office theme for a part that is not there, and does not need a warning callback', () => {
    expect(parseTheme(undefined)).toEqual(FALLBACK_THEME);
    expect(parseTheme(undefined, () => undefined).colors).toEqual(FALLBACK_THEME.colors);
  });
});
