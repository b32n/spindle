import { DEFAULT_THEME_COLORS } from './colors';
import { readFont } from './font';
import { el } from './test-support';

const ctx = { themeColors: DEFAULT_THEME_COLORS };

describe('readFont', () => {
  it('reads a styles font', () => {
    expect(
      readFont(el('<font><b/><i val="1"/><strike val="true"/><u/><vertAlign val="superscript"/><sz val="10.5"/><color theme="1"/><name val="Arial"/><family val="2"/><charset val="1"/><scheme val="minor"/></font>'), ctx),
    ).toEqual({
      bold: true,
      italic: true,
      strike: true,
      underline: 'single',
      verticalAlign: 'superscript',
      size: 10.5,
      color: { theme: 1, resolved: '#000000' },
      name: 'Arial',
      family: 2,
      charset: 1,
      scheme: 'minor',
    });
  });
  it('reads a run property (the face is rFont) and says "off" when a flag is switched off', () => {
    expect(readFont(el('<rPr><rFont val="Calibri"/><b val="0"/><i val="false"/><strike val="0"/><u val="double"/><sz val="9"/></rPr>'), ctx)).toEqual({
      name: 'Calibri',
      bold: false,
      italic: false,
      strike: false,
      underline: 'double',
      size: 9,
    });
  });
  it.each(['singleAccounting', 'doubleAccounting'])('keeps the %s underline', (u) => expect(readFont(el(`<font><u val="${u}"/></font>`), ctx)?.underline).toBe(u));
  it('drops an underline of none or an unknown kind, an unknown vertical alignment and unreadable numbers', () => {
    expect(readFont(el('<font><u val="none"/><vertAlign val="middle"/><sz val="big"/><family val="x"/></font>'), ctx)).toEqual({});
  });
  it('is an empty font for an empty element and undefined for none', () => {
    expect(readFont(el('<font/>'), ctx)).toEqual({});
    expect(readFont(undefined, ctx)).toBeUndefined();
  });
});
