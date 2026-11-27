import { builtinNumberFormat, parseStyles } from './styles-scan';
import { parseXml } from './xml';

const sheet = (inner: string) =>
  parseXml(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${inner}</styleSheet>`) ?? undefined;

const ALL = new Set([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);

describe('builtinNumberFormat', () => {
  it('knows the en-US codes of the built-in ids and nothing for the locale-specific ones', () => {
    expect(builtinNumberFormat(9)).toBe('0%');
    expect(builtinNumberFormat(14)).toBe('m/d/yyyy');
    expect(builtinNumberFormat(49)).toBe('@');
    expect(builtinNumberFormat(27)).toBeUndefined();
    expect(builtinNumberFormat(0)).toBeUndefined();
  });
});

describe('parseStyles number formats', () => {
  const styles = parseStyles(
    sheet(
      '<numFmts count="2"><numFmt numFmtId="164" formatCode="0.00"/><numFmt numFmtId="165" formatCode="&quot;x&quot;0"/></numFmts>' +
        '<cellXfs><xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="164"/><xf numFmtId="165"/><xf numFmtId="14"/><xf numFmtId="27"/><xf/></cellXfs>'
    )
  );

  it('lists the distinct formats of the used cell formats, General excluded, codes decoded', () => {
    expect(styles.numberFormats(new Set([0, 1, 2, 3, 4, 5, 6]))).toEqual([
      { code: '0.00', label: '0.00' },
      { code: '"x"0', label: '"x"0' },
      { code: 'm/d/yyyy', label: 'm/d/yyyy' },
      { code: undefined, label: 'built-in format 27' },
    ]);
  });

  it('only looks at the cell formats in use and ignores ids that do not exist', () => {
    expect(styles.numberFormats(new Set([0, 3, 99]))).toEqual([{ code: '"x"0', label: '"x"0' }]);
    expect(styles.numberFormats(new Set())).toEqual([]);
  });

  it('is empty for a missing or empty style sheet', () => {
    expect(parseStyles(undefined).numberFormats(ALL)).toEqual([]);
    expect(parseStyles(sheet('')).usedFacts(ALL)).toEqual({ patternFill: 0, gradientFill: 0, diagonalBorder: 0, namedStyles: [] });
  });

  it('falls back to an empty code for a custom format without one', () => {
    const s = parseStyles(sheet('<numFmts><numFmt numFmtId="170"/></numFmts><cellXfs><xf numFmtId="170"/></cellXfs>'));
    expect(s.numberFormats(new Set([0]))).toEqual([{ code: '', label: '' }]);
  });
});

describe('parseStyles used facts', () => {
  const fonts =
    '<fonts>' +
    '<font><color theme="1"/></font>' + // 0 plain black text: not a loss
    '<font><color theme="1" tint="0.5"/></font>' + // 1 tinted
    '<font><color theme="4"/></font>' + // 2
    '<font><color indexed="8"/></font>' + // 3 default black
    '<font><color indexed="10"/></font>' + // 4
    '<font><color rgb="FFFF0000"/></font>' + // 5 rgb: readable
    '<font><color indexed="64"/></font>' + // 6 system foreground
    '<font/>' + // 7
    '</fonts>';
  const fills =
    '<fills>' +
    '<fill><patternFill patternType="none"/></fill>' + // 0
    '<fill><patternFill patternType="gray125"/></fill>' + // 1
    '<fill><patternFill patternType="solid"><fgColor rgb="FF112233"/><bgColor indexed="64"/></patternFill></fill>' + // 2
    '<fill><patternFill patternType="solid"><fgColor theme="5"/><bgColor indexed="65"/></patternFill></fill>' + // 3
    '<fill><patternFill patternType="lightGray"><fgColor rgb="FF000000"/></patternFill></fill>' + // 4
    '<fill><gradientFill degree="90"><stop position="0"><color rgb="FF000000"/></stop><stop position="1"><color theme="3"/></stop></gradientFill></fill>' + // 5
    '<fill><gradientFill><stop position="0"><color rgb="FF000000"/></stop></gradientFill></fill>' + // 6
    '<fill><patternFill><bgColor indexed="30"/></patternFill></fill>' + // 7: no type, indexed bg
    '</fills>';
  const borders =
    '<borders>' +
    '<border><left/><diagonal/></border>' + // 0
    '<border><left style="thin"><color theme="2"/></left><right/></border>' + // 1 themed side with a style
    '<border><left><color theme="2"/></left></border>' + // 2 colour but no style: invisible
    '<border><top style="thin"><color auto="1"/></top><bottom style="thin"><color rgb="FF000000"/></bottom></border>' + // 3
    '<border><diagonal style="thin" diagonalUp="1"/></border>' + // 4
    '<border><diagonal/></border>' + // 5
    '</borders>';
  const xfs = (list: Array<[number, number, number, number?]>) =>
    `<cellXfs>${list.map(([f, fi, b, x]) => `<xf fontId="${f}" fillId="${fi}" borderId="${b}" xfId="${x ?? 0}"/>`).join('')}</cellXfs>`;

  it('counts the used cell formats with a pattern or gradient fill or a diagonal border', () => {
    const s = parseStyles(
      sheet(
        fonts + fills + borders +
          xfs([
            [0, 0, 0], // 0 nothing
            [1, 0, 0], // 1 tinted theme font
            [2, 0, 0], // 2 theme font
            [3, 0, 0], // 3 indexed 8: default
            [4, 0, 0], // 4 indexed font
            [5, 2, 0], // 5 rgb font, rgb fill: readable
            [6, 0, 0], // 6 indexed 64
            [7, 3, 0], // 7 theme fill
            [0, 4, 0], // 8 pattern fill (rgb colour)
            [0, 5, 0], // 9 gradient with a theme stop
            [0, 0, 1], // 10 themed border side
            [0, 0, 2], // 11 colour on a side with no style
            [0, 7, 3], // 12 indexed bg, auto border colour
          ])
      )
    );
    const facts = s.usedFacts(ALL);
    expect(facts.patternFill).toBe(1); // xf 8 (xf 12 has no pattern type)
    expect(facts.gradientFill).toBe(1);
    expect(facts.diagonalBorder).toBe(0);
  });

  it('counts only the formats in use', () => {
    const s = parseStyles(sheet(fonts + fills + borders + xfs([[0, 0, 0], [2, 0, 0], [0, 5, 0], [0, 0, 4]])));
    expect(s.usedFacts(new Set([0, 1])).gradientFill).toBe(0);
    expect(s.usedFacts(new Set([2])).gradientFill).toBe(1);
    expect(s.usedFacts(new Set([3])).diagonalBorder).toBe(1);
    expect(s.usedFacts(new Set([99]))).toEqual({ patternFill: 0, gradientFill: 0, diagonalBorder: 0, namedStyles: [] });
  });

  it('reports diagonal borders only when they have a style', () => {
    const s = parseStyles(sheet(fonts + fills + borders + xfs([[0, 0, 5], [0, 0, 4]])));
    expect(s.usedFacts(new Set([0])).diagonalBorder).toBe(0);
    expect(s.usedFacts(new Set([1])).diagonalBorder).toBe(1);
  });

  it('names the named styles in use, never Normal, each once', () => {
    const s = parseStyles(
      sheet(
        fonts + fills + borders +
          xfs([[0, 0, 0, 0], [0, 0, 0, 1], [0, 0, 0, 1], [0, 0, 0, 2], [0, 0, 0, 3], [0, 0, 0, 9]]) +
          '<cellStyles><cellStyle name="Normal" xfId="0"/><cellStyle name="Good" xfId="1"/><cellStyle name="Bad" xfId="2"/><cellStyle xfId="3"/></cellStyles>'
      )
    );
    expect(s.usedFacts(new Set([0, 1, 2, 3, 4, 5])).namedStyles).toEqual(['Good', 'Bad']);
    expect(s.usedFacts(new Set([0])).namedStyles).toEqual([]);
  });
});
