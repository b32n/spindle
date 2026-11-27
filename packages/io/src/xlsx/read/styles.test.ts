import { DEFAULT_THEME_COLORS } from './colors';
import { numberFormatCode, parseStyles } from './styles';
import { el, MAIN_NS } from './test-support';

const parse = (body: string) => parseStyles(el(`<styleSheet ${MAIN_NS}>${body}</styleSheet>`), DEFAULT_THEME_COLORS);

describe('parseStyles: number formats', () => {
  it('reads custom formats (entities decoded) and skips entries without a usable id', () => {
    const s = parse('<numFmts count="3"><numFmt numFmtId="164" formatCode="0.0&quot;x&quot;"/><numFmt numFmtId="x" formatCode="bad"/><numFmt numFmtId="165"/></numFmts>');
    expect(s.numFmts).toEqual([{ id: 164, code: '0.0"x"' }, { id: 165, code: '' }]);
  });
  it('maps ids to codes: General, custom, built-in en-US, or unknown', () => {
    const s = parse('<numFmts><numFmt numFmtId="14" formatCode="dd/mm/yyyy"/><numFmt numFmtId="170" formatCode="0.00"/></numFmts>');
    expect(numberFormatCode(s, 0)).toBe('General');
    expect(numberFormatCode(s, 14)).toBe('dd/mm/yyyy'); // the file overrides the built-in
    expect(numberFormatCode(s, 170)).toBe('0.00');
    expect(numberFormatCode(s, 9)).toBe('0%');
    expect(numberFormatCode(s, 28)).toBeUndefined(); // locale-specific built-in
  });
});

describe('parseStyles: fonts, fills, borders', () => {
  const s = parse(`
    <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color theme="4" tint="0.5"/></font></fonts>
    <fills count="6">
      <fill><patternFill patternType="none"/></fill>
      <fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill>
      <fill><patternFill patternType="darkGrid"><fgColor theme="5"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>
      <fill><patternFill/></fill>
      <fill><gradientFill degree="90"><stop position="0"><color rgb="FFFFFFFF"/></stop><stop position="1"><color theme="4"/></stop><stop/></gradientFill></fill>
      <fill><gradientFill type="path" left="0.5" right="0.5" top="0.25" bottom="0.75"><stop position="0"><color indexed="10"/></stop></gradientFill></fill>
    </fills>
    <borders count="3">
      <border><left/><right/><top/><bottom/><diagonal/></border>
      <border diagonalUp="1" diagonalDown="0" outline="0"><left style="thin"><color rgb="FF000000"/></left><right style="medium"><color theme="1"/></right><top style="dashed"/><bottom style="double"><color auto="1"/></bottom><diagonal style="hair"><color indexed="12"/></diagonal></border>
      <border><start style="thin"/><end style="thick"/></border>
    </borders>`);
  it('keeps fonts in file order with resolved colours', () => {
    expect(s.fonts).toHaveLength(2);
    expect(s.fonts[0]).toEqual({ size: 11, name: 'Calibri' });
    expect(s.fonts[1].bold).toBe(true);
    expect(s.fonts[1].color?.resolved).toMatch(/^#[0-9a-f]{6}$/);
  });
  it('reads pattern fills (solid, patterned, and the empty form as none)', () => {
    expect(s.fills[0]).toEqual({ type: 'pattern', pattern: 'none' });
    expect(s.fills[1]).toEqual({
      type: 'pattern',
      pattern: 'solid',
      fg: { rgb: 'FFC6EFCE', resolved: '#c6efce' },
      bg: { indexed: 64, resolved: '#000000' },
    });
    expect(s.fills[2]).toMatchObject({ pattern: 'darkGrid', fg: { theme: 5, resolved: '#ed7d31' }, bg: { resolved: '#ffffff' } });
    expect(s.fills[3]).toEqual({ type: 'pattern', pattern: 'none' });
  });
  it('reads gradient fills (linear and path) with stops', () => {
    expect(s.fills[4]).toEqual({
      type: 'gradient',
      gradientType: 'linear',
      degree: 90,
      stops: [
        { position: 0, color: { rgb: 'FFFFFFFF', resolved: '#ffffff' } },
        { position: 1, color: { theme: 4, resolved: '#4472c4' } },
        { position: 0, color: {} },
      ],
    });
    expect(s.fills[5]).toEqual({
      type: 'gradient',
      gradientType: 'path',
      left: 0.5,
      right: 0.5,
      top: 0.25,
      bottom: 0.75,
      stops: [{ position: 0, color: { indexed: 10, resolved: '#ff0000' } }],
    });
  });
  it('reads borders: styles with colours, diagonal and flags; sides without a style are absent', () => {
    expect(s.borders[0]).toEqual({});
    expect(s.borders[1]).toEqual({
      left: { style: 'thin', color: { rgb: 'FF000000', resolved: '#000000' } },
      right: { style: 'medium', color: { theme: 1, resolved: '#000000' } },
      top: { style: 'dashed' },
      bottom: { style: 'double', color: { auto: true } },
      diagonal: { style: 'hair', color: { indexed: 12, resolved: '#0000ff' } },
      diagonalUp: true,
      diagonalDown: false,
      outline: false,
    });
    expect(s.borders[2]).toEqual({ left: { style: 'thin' }, right: { style: 'thick' } }); // start/end read as left/right
  });
});

describe('parseStyles: formats, named styles, dxfs, palette', () => {
  const s = parse(`
    <colors><indexedColors><rgbColor rgb="FF010203"/><rgbColor rgb="bogus"/></indexedColors></colors>
    <fills><fill><patternFill patternType="solid"><fgColor indexed="0"/><bgColor indexed="1"/></patternFill></fill></fills>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="3">
      <xf numFmtId="14" fontId="1" fillId="2" borderId="3" xfId="0" applyNumberFormat="1" applyFont="0" applyFill="true" applyBorder="1" applyAlignment="1" applyProtection="1" quotePrefix="1" pivotButton="0">
        <alignment horizontal="center" vertical="top" textRotation="45" wrapText="1" indent="2" relativeIndent="-1" justifyLastLine="1" shrinkToFit="0" readingOrder="2"/>
        <protection locked="0" hidden="1"/>
      </xf>
      <xf/>
      <xf numFmtId="x" fontId="1.5"><alignment/><protection/></xf>
    </cellXfs>
    <cellStyles count="3"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Hidden one" xfId="1" hidden="1"/><cellStyle xfId="2"/><cellStyle name="NoId"/></cellStyles>
    <dxfs count="4">
      <dxf><font><b/><color rgb="FF9C0006"/></font><numFmt numFmtId="14" formatCode="d/m"/><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill><alignment horizontal="left"/><border><left style="thin"/></border><protection locked="1"/></dxf>
      <dxf><numFmt numFmtId="9"/><fill><gradientFill><stop position="0"><color theme="0"/></stop></gradientFill></fill></dxf>
      <dxf><numFmt numFmtId="x"/></dxf><dxf><numFmt numFmtId="28"/></dxf>
    </dxfs>`);
  it('reads cell formats with flags, alignment and protection as written', () => {
    expect(s.cellStyleXfs).toEqual([{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }]);
    expect(s.cellXfs[0]).toEqual({
      numFmtId: 14,
      fontId: 1,
      fillId: 2,
      borderId: 3,
      xfId: 0,
      applyNumberFormat: true,
      applyFont: false,
      applyFill: true,
      applyBorder: true,
      applyAlignment: true,
      applyProtection: true,
      quotePrefix: true,
      pivotButton: false,
      alignment: { horizontal: 'center', vertical: 'top', textRotation: 45, wrapText: true, indent: 2, relativeIndent: -1, justifyLastLine: true, shrinkToFit: false, readingOrder: 2 },
      protection: { locked: false, hidden: true },
    });
  });
  it('defaults missing or unreadable ids to 0 and keeps empty alignment/protection', () => {
    expect(s.cellXfs[1]).toEqual({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 });
    expect(s.cellXfs[2]).toEqual({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, alignment: {}, protection: {} });
  });
  it('reads named styles, skipping one without a name', () => {
    expect(s.cellStyles).toEqual([
      { name: 'Normal', xfId: 0, builtinId: 0 },
      { name: 'Hidden one', xfId: 1, hidden: true },
      { name: 'NoId', xfId: 0 },
    ]);
  });
  it('reads dxfs: only what each overrides, with the built-in code for a format without one', () => {
    expect(s.dxfs[0]).toEqual({
      font: { bold: true, color: { rgb: 'FF9C0006', resolved: '#9c0006' } },
      numFmt: { id: 14, code: 'd/m' },
      fill: { type: 'pattern', pattern: 'none', bg: { rgb: 'FFFFC7CE', resolved: '#ffc7ce' } },
      alignment: { horizontal: 'left' },
      border: { left: { style: 'thin' } },
      protection: { locked: true },
    });
    expect(s.dxfs[1]).toEqual({
      numFmt: { id: 9, code: '0%' },
      fill: { type: 'gradient', gradientType: 'linear', stops: [{ position: 0, color: { theme: 0, resolved: '#ffffff' } }] },
    });
    expect(s.dxfs[2]).toEqual({}); // an unreadable id is no format
    expect(s.dxfs[3]).toEqual({ numFmt: { id: 28, code: '' } });
  });
  it('uses the workbook palette, keeping the default for an entry that is not a colour', () => {
    expect(s.indexedColors).toEqual(['#010203', '#ffffff']);
    expect(s.fills[0]).toMatchObject({ fg: { indexed: 0, resolved: '#010203' }, bg: { indexed: 1, resolved: '#ffffff' } });
  });
  it('keeps default for entries beyond the 64 defaults and has no palette when none is written', () => {
    const many = parse(`<colors><indexedColors>${'<rgbColor rgb="bogus"/>'.repeat(66)}</indexedColors></colors>`);
    expect(many.indexedColors).toHaveLength(66);
    expect(many.indexedColors?.[65]).toBe('#000000');
    expect(parse('').indexedColors).toBeUndefined();
  });
});

describe('parseStyles: absent parts', () => {
  it('gives empty tables for a missing root or an empty stylesheet', () => {
    const empty = { numFmts: [], fonts: [], fills: [], borders: [], cellStyleXfs: [], cellXfs: [], cellStyles: [], dxfs: [] };
    expect(parseStyles(undefined, DEFAULT_THEME_COLORS)).toEqual(empty);
    expect(parse('')).toEqual(empty);
  });
});
