import { DEFAULT_THEME_COLORS } from './colors';
import { isDateFormat, lookOf, unescapeLiterals } from './cell-style';
import { parseStyles } from './styles';
import { el, MAIN_NS } from './test-support';

const styles = (inner: string) => parseStyles(el(`<styleSheet ${MAIN_NS}>${inner}</styleSheet>`), DEFAULT_THEME_COLORS);

const look = (xf: string, tables = '', isDefault = false) => {
  const s = styles(`${tables}<cellXfs>${xf}</cellXfs>`);
  return lookOf(s.cellXfs[0], s, isDefault);
};

describe('lookOf', () => {
  it('maps fonts, with the colour resolved', () => {
    const l = look('<xf fontId="1"/>', '<fonts><font/><font><b/><i/><u val="double"/><strike/><sz val="12"/><name val="Arial"/><color theme="4"/></font></fonts>');
    expect(l.style).toEqual({ bold: true, italic: true, underline: true, strikethrough: true, fontFamily: 'Arial', fontSize: 12, fontColor: '#4472c4' });
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

  it('maps solid fills only, with the colour resolved', () => {
    const fill = (inner: string) => `<fills><fill>${inner}</fill></fills>`;
    expect(look('<xf/>', fill('<patternFill patternType="solid"><fgColor theme="5" tint="0.4"/></patternFill>')).style?.backgroundColor).toBe('#f4b183');
    expect(look('<xf/>', fill('<patternFill patternType="solid"><fgColor rgb="FF112233"/></patternFill>')).style?.backgroundColor).toBe('#112233');
    expect(look('<xf/>', fill('<patternFill patternType="solid"/>')).style).toBeUndefined();
    expect(look('<xf/>', fill('<patternFill patternType="lightGray"><fgColor rgb="FF112233"/></patternFill>')).style).toBeUndefined();
    expect(look('<xf/>', fill('<gradientFill><stop position="0"><color rgb="FF112233"/></stop></gradientFill>')).style).toBeUndefined();
  });

  it('maps alignment: known horizontal and vertical values (center is middle), wrap, rotation', () => {
    const a = (attrs: string) => look(`<xf><alignment ${attrs}/></xf>`).style;
    expect(a('horizontal="center" vertical="center" wrapText="1"')).toEqual({ horizontalAlign: 'center', verticalAlign: 'middle', wrapText: true });
    expect(a('horizontal="right" vertical="top"')).toEqual({ horizontalAlign: 'right', verticalAlign: 'top' });
    expect(a('horizontal="left" vertical="bottom"')).toEqual({ horizontalAlign: 'left', verticalAlign: 'bottom' });
    expect(a('horizontal="general" vertical="distributed"')).toBeUndefined();
    expect(a('horizontal="fill"')).toBeUndefined();
  });

  it('maps rotation to -90..90: 91 to 180 turn clockwise, 255 (stacked) is not modelled', () => {
    const rot = (v: string) => look(`<xf><alignment textRotation="${v}"/></xf>`).style?.textRotation;
    expect([rot('0'), rot('45'), rot('90')]).toEqual([0, 45, 90]);
    expect([rot('91'), rot('135'), rot('180')]).toEqual([-1, -45, -90]);
    expect(rot('255')).toBeUndefined();
    expect(rot('181')).toBeUndefined();
    expect(look('<xf><alignment/></xf>').style).toBeUndefined();
  });

  it('maps borders: known styles, others thin, the colour when it resolves', () => {
    const b = (sides: string) => look('<xf borderId="0"/>', `<borders><border>${sides}</border></borders>`).style?.border;
    expect(b('<left style="thin"><color rgb="FFFF0000"/></left><right style="double"/><top style="dashed"/><bottom style="hair"><color auto="1"/></bottom>')).toEqual({
      left: { style: 'thin', color: '#ff0000' },
      right: { style: 'double' },
      top: { style: 'dashed' },
      bottom: { style: 'thin' },
    });
    expect(b('<left style="mediumDashDot"/>')).toEqual({ left: { style: 'thin' } });
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
    for (const unit of ['[', '"', '["', '[]"', '"[']) {
      const t0 = Date.now();
      isDateFormat(unit.repeat(1_000_000) + 'y');
      expect(Date.now() - t0).toBeLessThan(1000);
    }
  });
});

// The number format codes of a file as the editor gets them (`unescapeLiterals`): the backslash goes before a character Excel
// shows without quotes and stays before everything else, and nothing in quotes is touched. The first four rows are what LibreOffice
// writes for the S-FIO-01 file; the rest is the table of the S2b review (each of these changed meaning when every escape was dropped).
const CODES: Array<[raw: string, editor: string, isDate: boolean, why: string]> = [
  // S-FIO-01 and the other escapes of the plain characters
  ['\\$#,##0.00', '$#,##0.00', false, 'LibreOffice escapes the currency sign'],
  ['yyyy\\-mm\\-dd', 'yyyy-mm-dd', true, 'and the date separators'],
  ['_(\\$* #,##0_)', '_($* #,##0_)', false, 'an accounting code: `_(` and `* ` are pairs'],
  ['[$\\$-409]#,##0.00', '[$$-409]#,##0.00', false, 'a locale group is unescaped too'],
  ['\\(0\\)', '(0)', false, 'parentheses'],
  ['0\\+0', '0+0', false, 'plus'],
  ['h\\:mm', 'h:mm', true, 'colon'],
  ['d\\/m\\/yyyy', 'd/m/yyyy', true, 'slash'],
  ['0\\!', '0!', false, 'exclamation mark'],
  ['0\\^2', '0^2', false, 'caret'],
  ['0\\&0', '0&0', false, 'ampersand'],
  ["0\\'", "0'", false, 'apostrophe'],
  ['0\\~', '0~', false, 'tilde'],
  ['\\{0\\}', '{0}', false, 'braces'],
  ['\\<0\\>', '<0>', false, 'angle brackets'],
  ['0\\=0', '0=0', false, 'equals'],
  ['0\\ kg', '0 kg', false, 'a space'],
  // the review's table: the escape is what keeps the character literal
  ['0.0\\"', '0.0\\"', false, 'inches: unescaping would leave an unbalanced quote'],
  ['0\\ \\m', '0 \\m', false, '`m` would become a month token'],
  ['0\\ \\d', '0 \\d', false, '`d` would become a day token (and a date in a 1904 workbook)'],
  ['\\*0\\*', '\\*0\\*', false, 'would become fill syntax'],
  ['\\@', '\\@', false, 'would become the text placeholder'],
  ['0\\.00', '0\\.00', false, 'a literal dot, not the decimal point'],
  ['0\\\\', '0\\\\', false, 'a literal backslash: the next character would be escaped'],
  ['"C:\\Temp"0', '"C:\\Temp"0', false, 'a backslash in quotes is a backslash'],
  // the others that are kept: letters, digits, `,` `%` `#` `?` `_` `E`
  ['0\\,0', '0\\,0', false, 'a literal comma, not the thousands separator'],
  ['0\\%', '0\\%', false, 'a literal percent sign (no scaling)'],
  ['0\\#', '0\\#', false, 'a literal number sign'],
  ['0\\?', '0\\?', false, 'a literal question mark'],
  ['0\\_', '0\\_', false, 'a literal underscore'],
  ['0\\E+0', '0\\E+0', false, 'a literal E'],
  ['0\\5', '0\\5', false, 'a literal digit'],
  ['\\y\\m\\d', '\\y\\m\\d', false, 'letters shown as they are, no date'],
  ['0\\h', '0\\h', false, 'an hour sign, shown'],
  // the pieces that are not escapes
  ['_\\)', '_\\)', false, '`_x` takes its character as it is'],
  ['*\\-0', '*\\-0', false, '`*x` takes its character as it is'],
  ['"a\\"x\\-c"', '"a\\"x-c"', false, 'a backslash does not escape a quote in a run: the quote closes it, and the escape after it is read'],
  ['"a\\-c"0', '"a\\-c"0', false, 'inside a run nothing changes'],
  ['0\\', '0\\', false, 'a backslash at the end stays'],
  ['[Red]\\-0', '[Red]-0', false, 'a colour group and an escape'],
  ['[$\\\\-409]0', '[$\\\\-409]0', false, 'an escaped backslash in a group is one escape, not an escaped `-`'],
  ['[xx \\-0', '[xx -0', false, 'an unclosed group is ordinary text'],
  ['0 "xx \\-0', '0 "xx -0', false, 'an unclosed quote is ordinary text'],
];

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
    const t0 = Date.now();
    for (const xf of s.cellXfs) lookOf(xf, s, false);
    expect(Date.now() - t0).toBeLessThan(2_000);
  });
});

describe('unescapeLiterals and isDateFormat are linear on long codes', () => {
  const units = ['\\', '\\-', '"', '[', ']', '[$\\$', '"\\', '_', '*', '\\"', '[]"', '"[', '\\\\', '0\\ \\d'];
  it.each(units)('a million of %j', (unit) => {
    const code = unit.repeat(1_000_000) + 'y';
    const t0 = Date.now();
    unescapeLiterals(code);
    isDateFormat(code);
    expect(Date.now() - t0).toBeLessThan(3_000);
  }, 30_000);
});
