import { DEFAULT_THEME_COLORS } from './colors';
import { isDateFormat, lookOf, LookCache, unescapeLiterals } from './cell-style';
import { parseStyles } from './styles';
import { el, MAIN_NS, numberFormatCases } from './test-support';

const styles = (inner: string) => parseStyles(el(`<styleSheet ${MAIN_NS}>${inner}</styleSheet>`), DEFAULT_THEME_COLORS);

const look = (xf: string, tables = '', isDefault = false) => {
  const s = styles(`${tables}<cellXfs>${xf}</cellXfs>`);
  return lookOf(s.cellXfs[0], s, isDefault);
};

describe('lookOf', () => {
  it('maps fonts, with the colour resolved', () => {
    const l = look('<xf fontId="1"/>', '<fonts><font/><font><b/><i/><u val="double"/><strike/><sz val="12"/><name val="Arial"/><color theme="4"/></font></fonts>');
    expect(l.style).toEqual({ bold: true, italic: true, underline: true, strikethrough: true, fontFamily: 'Arial', fontSize: 12, fontColor: '#4472c4', underlineStyle: 'double' });
  });

  it('applies the tint of a theme colour and resolves indexed and rgb colours', () => {
    const fonts = (color: string) => `<fonts><font><color ${color}/></font></fonts>`;
    expect(look('<xf/>', fonts('theme="4" tint="0.8"')).style?.fontColor).toBe('#dae3f3');
    expect(look('<xf/>', fonts('indexed="10"')).style?.fontColor).toBe('#ff0000');
    expect(look('<xf/>', fonts('rgb="FF00FF00"')).style?.fontColor).toBe('#00ff00');
    expect(look('<xf/>', fonts('auto="1"')).style).toBeUndefined();
  });

  it('leaves plain black text to the editor: theme 1 untinted, indexed 8 and 64', () => {
    const fonts = (color: string) => `<fonts><font><color ${color}/></font></fonts>`;
    for (const c of ['theme="1"', 'indexed="8"', 'indexed="64"']) expect(look('<xf/>', fonts(c)).style).toBeUndefined();
    expect(look('<xf/>', fonts('theme="1" tint="0.5"')).style?.fontColor).toBe('#808080');
    expect(look('<xf/>', fonts('theme="1" tint="0"')).style).toBeUndefined();
    expect(look('<xf/>', fonts('indexed="65"')).style?.fontColor).toBe('#ffffff');
  });

  it('maps solid fills, pattern fills (pattern, its colour, the colour behind it) and gradient fills (first stop), with the colours resolved', () => {
    const fill = (inner: string) => `<fills><fill>${inner}</fill></fills>`;
    expect(look('<xf/>', fill('<patternFill patternType="solid"><fgColor theme="5" tint="0.4"/></patternFill>')).style?.backgroundColor).toBe('#f4b183');
    expect(look('<xf/>', fill('<patternFill patternType="solid"><fgColor rgb="FF112233"/></patternFill>')).style?.backgroundColor).toBe('#112233');
    expect(look('<xf/>', fill('<patternFill patternType="solid"/>')).style).toBeUndefined();
    expect(look('<xf/>', fill('<patternFill patternType="lightGray"><fgColor rgb="FF112233"/></patternFill>')).style).toEqual({ fillPattern: 'lightGray', patternColor: '#112233' });
    expect(look('<xf/>', fill('<patternFill patternType="darkGrid"><fgColor rgb="FF112233"/><bgColor rgb="FFEEDDCC"/></patternFill>')).style).toEqual({
      fillPattern: 'darkGrid', patternColor: '#112233', backgroundColor: '#eeddcc',
    });
    // a pattern without a colour has nothing to draw, and "none" is no fill
    expect(look('<xf/>', fill('<patternFill patternType="lightGray"/>')).style).toBeUndefined();
    expect(look('<xf/>', fill('<patternFill patternType="none"><fgColor rgb="FF112233"/></patternFill>')).style).toBeUndefined();
    // a gradient is its stop with the lowest position, whatever the order they are written in
    expect(look('<xf/>', fill('<gradientFill><stop position="0"><color rgb="FF112233"/></stop></gradientFill>')).style).toEqual({ backgroundColor: '#112233' });
    expect(look('<xf/>', fill('<gradientFill><stop position="1"><color rgb="FF000001"/></stop><stop position="0.25"><color rgb="FF445566"/></stop><stop position="0.5"><color rgb="FF000002"/></stop></gradientFill>')).style).toEqual({ backgroundColor: '#445566' });
    expect(look('<xf/>', fill('<gradientFill/>')).style).toBeUndefined();
  });

  it('maps alignment: known horizontal and vertical values (center is middle), wrap, rotation', () => {
    const a = (attrs: string) => look(`<xf><alignment ${attrs}/></xf>`).style;
    expect(a('horizontal="center" vertical="center" wrapText="1"')).toEqual({ horizontalAlign: 'center', verticalAlign: 'middle', wrapText: true });
    expect(a('horizontal="right" vertical="top"')).toEqual({ horizontalAlign: 'right', verticalAlign: 'top' });
    expect(a('horizontal="left" vertical="bottom"')).toEqual({ horizontalAlign: 'left', verticalAlign: 'bottom' });
    expect(a('horizontal="general" vertical="distributed"')).toBeUndefined();
    expect(a('horizontal="fill"')).toBeUndefined();
  });

  it('maps indent (levels, whole) and shrink-to-fit', () => {
    const a = (attrs: string) => look(`<xf><alignment ${attrs}/></xf>`).style;
    expect(a('horizontal="left" indent="2"')).toEqual({ horizontalAlign: 'left', indent: 2 });
    expect(a('indent="0"')).toBeUndefined();
    expect(a('indent="1.9"')).toBeUndefined(); // (the reader takes whole numbers only)
    expect(a('shrinkToFit="1"')).toEqual({ shrinkToFit: true });
    expect(a('shrinkToFit="0"')).toBeUndefined();
  });

  it('maps underline variants next to the plain flag, and strike-through', () => {
    const u = (value: string) => look('<xf fontId="1"/>', `<fonts><font/><font><u${value}/></font></fonts>`).style;
    expect(u('')).toEqual({ underline: true });
    expect(u(' val="single"')).toEqual({ underline: true });
    for (const kind of ['double', 'singleAccounting', 'doubleAccounting']) expect(u(` val="${kind}"`)).toEqual({ underline: true, underlineStyle: kind });
    expect(look('<xf fontId="1"/>', '<fonts><font/><font><strike/></font></fonts>').style).toEqual({ strikethrough: true });
  });

  it('maps rotation to -90..90: 91 to 180 turn clockwise, 255 (stacked) is not modelled', () => {
    const rot = (v: string) => look(`<xf><alignment textRotation="${v}"/></xf>`).style?.textRotation;
    expect([rot('0'), rot('45'), rot('90')]).toEqual([0, 45, 90]);
    expect([rot('91'), rot('135'), rot('180')]).toEqual([-1, -45, -90]);
    expect(rot('255')).toBeUndefined();
    expect(rot('181')).toBeUndefined();
    expect(look('<xf><alignment/></xf>').style).toBeUndefined();
  });

  it('maps borders: the thirteen styles, an unknown one thin, the colour when it resolves', () => {
    const b = (sides: string) => look('<xf borderId="0"/>', `<borders><border>${sides}</border></borders>`).style?.border;
    expect(b('<left style="thin"><color rgb="FFFF0000"/></left><right style="double"/><top style="dashed"/><bottom style="hair"><color auto="1"/></bottom>')).toEqual({
      left: { style: 'thin', color: '#ff0000' },
      right: { style: 'double' },
      top: { style: 'dashed' },
      bottom: { style: 'hair' },
    });
    expect(b('<left style="mediumDashDot"/><right style="bogus"/>')).toEqual({ left: { style: 'mediumDashDot' }, right: { style: 'thin' } });
    // all thirteen styles of the file format are kept as written, each side with its own colour
    const all = ['thin', 'hair', 'dotted', 'dashed', 'dashDot', 'dashDotDot', 'double', 'medium', 'mediumDashed', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot', 'thick'];
    for (const style of all) expect(b(`<left style="${style}"><color rgb="FF0000FF"/></left><top style="${style}"><color rgb="FF00FF00"/></top>`)).toEqual({
      left: { style, color: '#0000ff' }, top: { style, color: '#00ff00' },
    });
    expect(b('<left/><right><color rgb="FF000000"/></right>')).toBeUndefined();
  });

  it('gives the number format code, built-in or custom, and none for General', () => {
    const f = (id: number) => look(`<xf numFmtId="${id}"/>`, '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0&quot; kg&quot;"/></numFmts>').numberFormat;
    expect(f(0)).toBeUndefined();
    expect(f(2)).toEqual({ code: '0.00' });
    expect(f(14)).toEqual({ code: 'm/d/yyyy' });
    expect(f(164)).toEqual({ code: '0.0" kg"' });
    expect(f(27)).toBeUndefined();
  });

  it('keeps a built-in code as the table has it, and gives the code of a repeated id once (the first wins)', () => {
    const tables = '<numFmts count="2"><numFmt numFmtId="165" formatCode="0.0"/><numFmt numFmtId="165" formatCode="0.00"/></numFmts>';
    const f = (id: number) => look(`<xf numFmtId="${id}"/>`, tables).numberFormat;
    expect(f(165)).toEqual({ code: '0.0' });
    expect(f(5)).toEqual({ code: '"$"#,##0_);\\("$"#,##0\\)' });
    // a defined code for id 0 is not General's
    expect(look('<xf numFmtId="0"/>', '<numFmts><numFmt numFmtId="0" formatCode="0.0"/></numFmts>').numberFormat).toBeUndefined();
  });

  it('is empty for a format that names tables the file does not have', () => {
    expect(look('<xf fontId="9" fillId="9" borderId="9"/>')).toEqual({ isDate: false });
  });

  it('does not repeat the workbook default font, bottom alignment or zero rotation on the default format', () => {
    const tables = '<fonts><font><sz val="11"/><name val="Calibri"/><b/></font></fonts>';
    const xf = '<xf><alignment vertical="bottom" textRotation="0"/></xf>';
    expect(look(xf, tables, true).style).toEqual({ bold: true });
    expect(look(xf, tables, false).style).toEqual({ bold: true, fontFamily: 'Calibri', fontSize: 11, verticalAlign: 'bottom', textRotation: 0 });
    expect(look('<xf><alignment vertical="top" textRotation="90"/></xf>', '', true).style).toEqual({ verticalAlign: 'top', textRotation: 90 });
  });
});

describe('isDateFormat in either case', () => {
  it.each(['YYYY', 'DD', 'DDDD', 'H', 'S', 'MM', 'Y', 'B', 'hh:MM AM/PM', 'D-MMM'])('%s is a date format', (code) => expect(isDateFormat(code)).toBe(true));
  it.each(['General', 'GENERAL', '0.00E+00', '0.00e+00', '@', '0.00 "DD"', '[H]0', '\\D 0'])('%s is not', (code) => expect(isDateFormat(code)).toBe(false));
});

describe('isDateFormat', () => {
  it.each(['yyyy-mm-dd', 'm/d/yy', 'h:mm', 'mm-dd-yy', '[$-409]mmmm d, yyyy', 'd-mmm', 'dd/mm/yyyy hh:mm:ss', 'mm:ss.0', '[h]:mm:ss', 'bb'])('%s is a date format', (code) => {
    expect(isDateFormat(code)).toBe(true);
  });

  it.each([undefined, '', 'General', '0.00', '#,##0', '0%', '0.00E+00', '"day" 0', '[Red]0.0', '@', '_(* #,##0_)'])('%s is not', (code) => {
    expect(isDateFormat(code)).toBe(false);
  });
});

describe('isDateFormat on hostile codes', () => {
  it('skips bracket groups and quoted runs, and keeps an unclosed one as it is', () => {
    expect(isDateFormat('[Red]0.00')).toBe(false);
    expect(isDateFormat('[h]:00')).toBe(false);
    expect(isDateFormat('"d" 0')).toBe(false);
    expect(isDateFormat('"d" 0 y')).toBe(true);
    expect(isDateFormat('[$-409]d')).toBe(true);
    expect(isDateFormat('0 "xxx')).toBe(false);
    expect(isDateFormat('0 "xxx d')).toBe(true);
    expect(isDateFormat('[unclosed y')).toBe(true);
    expect(isDateFormat('"a" [x] "b" m')).toBe(true);
  });

  it('skips escapes and the characters `_` and `*` take, which are not tokens', () => {
    expect(isDateFormat('0\\ \\d')).toBe(false);
    expect(isDateFormat('\\y\\m\\d\\h\\s')).toBe(false);
    expect(isDateFormat('\\\\d')).toBe(true);
    expect(isDateFormat('0_d')).toBe(false);
    expect(isDateFormat('*m0')).toBe(false);
    expect(isDateFormat('0\\-d')).toBe(true);
    expect(isDateFormat('0\\')).toBe(false);
    expect(isDateFormat('_')).toBe(false);
  });

  it('is linear in the length of a code that is a long run of unclosed brackets or quotes (the regex form was quadratic)', () => {
    // (a quadratic form never ends on these: the test's own timeout is the failure)
    for (const unit of ['[', '"', '["', '[]"', '"[']) expect(isDateFormat(unit.repeat(1_000_000) + 'y')).toBe(true);
  }, 120_000);
});


const CODES = numberFormatCases();

describe('number format codes: what the editor gets and what is a date', () => {
  const defined = (raw: string) => {
    const attr = raw.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
    const s = styles(`<numFmts count="1"><numFmt numFmtId="170" formatCode="${attr}"/></numFmts><cellXfs><xf numFmtId="170"/></cellXfs>`);
    return { s, look: lookOf(s.cellXfs[0], s, false) };
  };

  it.each(CODES)('%s is %s (date: %s): %s', (raw, editor, isDate) => {
    const { s, look: l } = defined(raw);
    expect(l.numberFormat).toEqual({ code: editor });
    expect(l.isDate).toBe(isDate);
    expect(unescapeLiterals(raw)).toBe(editor);
    // the workbook's styles keep the code as written
    expect(s.numFmts[0].code).toBe(raw);
  });

  it('is idempotent where the result has no escape of a plain character left', () => {
    for (const [, editor] of CODES) expect(unescapeLiterals(unescapeLiterals(editor))).toBe(unescapeLiterals(editor));
  });

  it('judges a date on the code as written: a built-in id by its table code, an id the table lacks as no date', () => {
    expect(look('<xf numFmtId="14"/>').isDate).toBe(true);
    expect(look('<xf numFmtId="2"/>').isDate).toBe(false);
    expect(look('<xf numFmtId="27"/>').isDate).toBe(false);
    expect(look('<xf numFmtId="0"/>').isDate).toBe(false);
  });

  it('does not scan a table of formats once per format: 20,000 formats, 20,000 formats that name built-in ids', () => {
    const fmts = Array.from({ length: 20_000 }, (_, i) => `<numFmt numFmtId="${1000 + i}" formatCode="0.${i}"/>`).join('');
    const xfs = Array.from({ length: 20_000 }, (_, i) => `<xf numFmtId="${i % 2 === 0 ? 14 : 1000 + i}" fontId="${i}"/>`).join('');
    const s = styles(`<numFmts>${fmts}</numFmts><cellXfs>${xfs}</cellXfs>`);
    const cache = new LookCache(s);
    for (let i = 0; i < s.cellXfs.length; i++) cache.get(i);
    for (const xf of s.cellXfs) lookOf(xf, s, false);
  }, 120_000);
});

describe('unescapeLiterals and isDateFormat are linear on long codes', () => {
  const units = ['\\', '\\-', '"', '[', ']', '[$\\$', '"\\', '_', '*', '\\"', '[]"', '"[', '\\\\', '0\\ \\d'];
  it.each(units)('a million of %j', (unit) => {
    const code = unit.repeat(1_000_000) + 'y';
    unescapeLiterals(code);
    isDateFormat(code);
  }, 120_000);
});

describe('LookCache: a format is read once however many cells and sheets use it', () => {
  const tables = (code: string, xfs = 3) => `<numFmts><numFmt numFmtId="164" formatCode="${code.replace(/"/g, '&quot;')}"/></numFmts>`.concat(`<cellXfs>${Array.from({ length: xfs }, () => '<xf numFmtId="164"/>').join('')}</cellXfs>`);

  it('gives the same look object for the same format index, and none for an index the file does not have (also remembered)', () => {
    const cache = new LookCache(styles('<cellXfs><xf/><xf numFmtId="14"/></cellXfs>'));
    expect(cache.get(1)).toBe(cache.get(1));
    expect(cache.get(1)?.isDate).toBe(true);
    expect(cache.get(0)?.numberFormat).toBeUndefined();
    expect(cache.get(9)).toBeUndefined();
    expect(cache.get(9)).toBeUndefined();
  });

  it('reads the code of a number format once for all the formats that name it', () => {
    const cache = new LookCache(styles(tables('yyyy\\-mm\\-dd', 3)));
    const [a, b, c] = [cache.get(0), cache.get(1), cache.get(2)];
    expect(a?.numberFormat).toEqual({ code: 'yyyy-mm-dd' });
    expect(a?.isDate).toBe(true);
    expect(cache.numberFormat(164)).toBe(cache.numberFormat(164));
    // equal but separate looks (a look is per format index), all pointing at the one code that was read
    expect([b?.numberFormat?.code, c?.numberFormat?.code]).toEqual(['yyyy-mm-dd', 'yyyy-mm-dd']);
  });

  it('a code of more than 255 characters is opaque: kept as written, not scanned (so not a date, whatever letters it holds), told of once however many formats use it', () => {
    const told: number[] = [];
    const long = `yyyy\\$${'0\\$'.repeat(200)}`;
    expect(long.length).toBeGreaterThan(255);
    const cache = new LookCache(styles(tables(long, 50)), (id) => told.push(id));
    for (let i = 0; i < 50; i++) {
      const l = cache.get(i);
      expect(l?.isDate).toBe(false);
      expect(l?.numberFormat?.code).toBe(long);
    }
    expect(told).toEqual([164]);
  });

  it('a code of exactly 255 characters is read as usual, with its escapes and its date letters', () => {
    const code = `d\\-${'0'.repeat(252)}`;
    expect(code).toHaveLength(255);
    const l = new LookCache(styles(tables(code, 1))).get(0);
    expect(l?.isDate).toBe(true);
    expect(l?.numberFormat?.code).toBe(`d-${'0'.repeat(252)}`);
  });

  it('works without anyone to tell (the default), and for ids that are built in or unknown', () => {
    const cache = new LookCache(styles(`${tables('x'.repeat(300), 1)}`));
    expect(cache.get(0)?.numberFormat?.code).toBe('x'.repeat(300));
    const builtin = new LookCache(styles('<cellXfs><xf numFmtId="14"/><xf numFmtId="0"/><xf numFmtId="999"/></cellXfs>'));
    expect(builtin.get(0)?.numberFormat?.code).toBe('m/d/yyyy');
    expect(builtin.get(1)?.numberFormat).toBeUndefined();
    expect(builtin.get(2)?.numberFormat).toBeUndefined();
    expect(builtin.numberFormat(999)).toEqual({ isDate: false, code: undefined, opaque: false });
  });
});
