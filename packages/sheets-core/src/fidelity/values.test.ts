import {
  MAX_MISMATCHES,
  compareValues,
  errorCode,
  parseAddress,
  readFileFormulas,
  readSheetFormulas,
  sameValue,
  sharedStrings,
  show,
  unescapeXml,
  unreadableValues,
  type FileFormulas,
} from '../../../../scripts/xlsx-fidelity/values';

describe('unescapeXml', () => {
  it('decodes entities, character references and _xHHHH_ escapes', () => {
    expect(unescapeXml('a &lt;b&gt; &amp; &quot;c&quot; &apos;d&apos; &#65;&#x42; _x0043_')).toBe(
      'a <b> & "c" \'d\' AB C'
    );
    expect(unescapeXml('&#1114112; &#x110000;')).toBe('� �');
    expect(unescapeXml('plain')).toBe('plain');
    // decoded once: an escaped ampersand never starts another escape
    expect(unescapeXml('&amp;#65;')).toBe('&#65;');
    expect(unescapeXml('&amp;_x0041_')).toBe('&A'); // a real _xHHHH_ escape still decodes
    expect(unescapeXml('_x005F_x0041_')).toBe('_x0041_'); // escaped underscore, not decoded twice
    expect(unescapeXml('&amp;amp; &#38;lt;')).toBe('&amp; &lt;');
  });
});

describe('sharedStrings', () => {
  it('reads plain, rich-text and empty items, skipping phonetic runs', () => {
    const xml =
      '<sst><si><t>one</t></si><si><r><t>tw</t></r><r><t xml:space="preserve">o </t></r></si><si/><si><t>a &amp; b</t><rPh sb="0"><t>ignored</t></rPh></si><si><t/></si></sst>';
    expect(sharedStrings(xml)).toEqual(['one', 'two ', '', 'a & b', '']);
    expect(sharedStrings(undefined)).toEqual([]);
  });
});

describe('parseAddress', () => {
  it('parses plain addresses and rejects anything else', () => {
    expect(parseAddress('A1')).toEqual([0, 0]);
    expect(parseAddress('AB12')).toEqual([11, 27]);
    expect(parseAddress('XFD1048576')).toEqual([1048575, 16383]);
    for (const bad of ['', 'a1', '1A', 'A', 'A0A', '$A$1', 'A1:B2'])
      expect(parseAddress(bad)).toBeNull();
  });
});

const sheet = (cells: string) =>
  `<worksheet><sheetData><row r="1">${cells}</row></sheetData></worksheet>`;

describe('readSheetFormulas', () => {
  const strings = ['zero', 'one'];
  const read = (cells: string) => readSheetFormulas(sheet(cells), strings, 'S', 0);

  it('reads the cached value of each formula cell by type', () => {
    const cells = read(
      [
        '<c r="A1"><f>1+1</f><v>2</v></c>',
        '<c r="B1" t="str"><f>"a"&amp;"b"</f><v>a &amp; b</v></c>',
        '<c r="C1" t="b"><f>TRUE</f><v>1</v></c>',
        '<c r="D1" t="b"><f>FALSE</f><v>0</v></c>',
        '<c r="E1" t="e"><f>1/0</f><v>#DIV/0!</v></c>',
        '<c r="F1" t="s"><f>X</f><v>1</v></c>',
        '<c r="G1" t="inlineStr"><f>X</f><is><t>in</t></is></c>',
        '<c r="H1" t="d"><f>X</f><v>2026-01-01T00:00:00Z</v></c>',
        '<c r="I1"><f t="shared" si="0"/><v>5</v></c>',
        '<c r="J1" t="str"><f t="array" ref="J1:J2">X</f><v></v></c>',
      ].join('')
    );
    expect(cells.map((c) => [c.address, c.expected, c.skip])).toEqual([
      ['A1', 2, null],
      ['B1', 'a & b', null],
      ['C1', true, null],
      ['D1', false, null],
      ['E1', '#DIV/0!', null],
      ['F1', 'one', null],
      ['G1', 'in', null],
      ['H1', '2026-01-01T00:00:00Z', null],
      ['I1', 5, null],
      ['J1', '', null],
    ]);
    expect(cells[0]).toMatchObject({ sheet: 'S', occurrence: 0 });
  });

  it('skips cells without a formula or an address, marks uncached and volatile formulas', () => {
    const cells = read(
      [
        '<c r="A1"><v>1</v></c>',
        '<c r="B1"/>',
        '<c><f>1</f><v>1</v></c>',
        '<c r="C1"><f>1</f></c>',
        '<c r="D1"><f>1</f><v></v></c>',
        '<c r="E1"><f>1</f><v>abc</v></c>',
        '<c r="F1" t="s"><f>1</f><v>99</v></c>',
        '<c r="G1"><f>NOW()</f><v>46000.5</v></c>',
        '<c r="H1"><f>rand() + 1</f></c>',
        '<c r="I1"><f>TODAY()</f><v>1</v></c>',
        '<c r="J1"><f>RANDBETWEEN(1,2)</f><v>1</v></c>',
        '<c r="K1"><f>OPERAND(1)</f><v>1</v></c>',
        '<c r="L1" t="inlineStr"><f>1</f></c>',
        '<c r="M1"><f>1</f><v/></c>',
      ].join('')
    );
    expect(cells.map((c) => [c.address, c.expected, c.skip])).toEqual([
      ['C1', null, 'uncached'],
      ['D1', null, 'uncached'],
      ['E1', null, 'uncached'],
      ['F1', null, 'uncached'],
      ['G1', 46000.5, 'volatile'],
      ['H1', null, 'volatile'],
      ['I1', 1, 'volatile'],
      ['J1', 1, 'volatile'],
      ['K1', 1, null],
      ['L1', '', null],
      ['M1', null, 'uncached'],
    ]);
  });

  it('survives empty or damaged XML', () => {
    expect(readSheetFormulas('', [], 'S', 0)).toEqual([]);
    expect(readSheetFormulas('<c r="A1"><f>1', [], 'S', 0)).toEqual([]);
  });
});

const workbook = (sheets: string[]) => `<workbook><sheets>${sheets.join('')}</sheets></workbook>`;
const rels = (pairs: Array<[string, string]>) =>
  `<Relationships>${pairs.map(([id, target]) => `<Relationship Id="${id}" Type="t" Target="${target}"/>`).join('')}</Relationships>`;

describe('readFileFormulas', () => {
  it('maps sheets to parts through the workbook relationships, numbering repeated names', () => {
    const file = readFileFormulas({
      xml: {
        'xl/workbook.xml': workbook([
          '<sheet name="A &amp; B" sheetId="1" r:id="rId1"/>',
          '<sheet name="Dup" sheetId="2" r:id="rId2"/>',
          '<sheet name="Dup" sheetId="3" r:id="rId3"/>',
          '<sheet name="Gone" sheetId="4" r:id="rId4"/>',
          '<sheet name="NoRel" sheetId="5"/>',
        ]),
        'xl/_rels/workbook.xml.rels': rels([
          ['rId1', 'worksheets/sheet1.xml'],
          ['rId2', 'worksheets/sheet2.xml'],
          ['rId3', '/xl/worksheets/sheet3.xml'],
          ['rId4', 'worksheets/missing.xml'],
        ]),
        'xl/sharedStrings.xml': '<sst><si><t>s</t></si></sst>',
        'xl/worksheets/sheet1.xml': sheet('<c r="A1" t="s"><f>x</f><v>0</v></c>'),
        'xl/worksheets/sheet2.xml': sheet('<c r="A1"><f>x</f><v>2</v></c>'),
        'xl/worksheets/sheet3.xml': sheet('<c r="A1"><f>x</f><v>3</v></c>'),
      },
    });
    expect(file.sheets).toEqual(['A & B', 'Dup', 'Dup', 'Gone', 'NoRel']);
    expect(file.unreadableSheets).toBe(2);
    expect(file.cells.map((c) => [c.sheet, c.occurrence, c.address, c.expected])).toEqual([
      ['A & B', 0, 'A1', 's'],
      ['Dup', 0, 'A1', 2],
      ['Dup', 1, 'A1', 3],
    ]);
  });

  it('is empty for a package without a workbook', () => {
    expect(readFileFormulas({ xml: {} })).toEqual({ sheets: [], cells: [], unreadableSheets: 0 });
    expect(
      readFileFormulas({ xml: { 'xl/workbook.xml': workbook(['<sheet r:id="rId1"/>']) } })
    ).toEqual({ sheets: [''], cells: [], unreadableSheets: 1 });
  });
});

describe('comparing values', () => {
  it('compares numbers with a tolerance and everything else exactly', () => {
    expect(sameValue(1, 1)).toBe(true);
    expect(sameValue(1.7320508075688772, 1.73205080756888)).toBe(true);
    expect(sameValue(1e12, 1e12 + 100)).toBe(true);
    expect(sameValue(1, 1.001)).toBe(false);
    expect(sameValue(0, 1e-16)).toBe(true);
    expect(sameValue(0, 1e-10)).toBe(false);
    expect(sameValue(1e-12, 0)).toBe(false);
    expect(sameValue(1e-12, 1.0000000001e-12)).toBe(true);
    expect(sameValue(1e-12, 1.1e-12)).toBe(false);
    expect(sameValue('a', 'a')).toBe(true);
    expect(sameValue('a', 'b')).toBe(false);
    expect(sameValue(true, true)).toBe(true);
    expect(sameValue(true, 1)).toBe(false);
    expect(sameValue(null, null)).toBe(true);
    expect(sameValue(0, null)).toBe(false);
    expect(sameValue('1', 1)).toBe(false);
  });

  it('treats errors as equal when the code is the same, whatever diagnostic follows', () => {
    expect(sameValue('#NAME?', '#NAME? Function PI not found')).toBe(true);
    expect(sameValue('#N/A', '#N/A')).toBe(true);
    expect(sameValue('#N/A', '#NAME?')).toBe(false);
    expect(sameValue('#N/A', 'text')).toBe(false);
    expect(sameValue('text', '#N/A')).toBe(false);
  });

  it('finds the error code of a value', () => {
    expect(errorCode('#DIV/0! x')).toBe('#DIV/0!');
    expect(errorCode('#SPILL!')).toBe('#SPILL!');
    expect(errorCode('#hash')).toBeNull();
    expect(errorCode(5)).toBeNull();
    expect(errorCode(null)).toBeNull();
    expect(errorCode(true)).toBeNull();
  });

  it('shows values for the report', () => {
    expect([show(null), show('a"b'), show(true), show(false), show(1.5)]).toEqual([
      '(empty)',
      '"a\\"b"',
      'TRUE',
      'FALSE',
      '1.5',
    ]);
  });
});

const formulas = (
  cells: Array<
    [string, string, Parameters<typeof sameValue>[0], FileFormulas['cells'][number]['skip']?]
  >
): FileFormulas => ({
  sheets: ['S'],
  unreadableSheets: 0,
  cells: cells.map(([address, , expected, skip]) => ({
    sheet: 'S',
    occurrence: 0,
    address,
    expected,
    skip: skip ?? null,
  })),
});
const model = (cells: Record<string, unknown>) => ({ sheets: [{ name: 'S', cells }] }) as never;

describe('compareValues', () => {
  it('counts matches and mismatches, classifies them and lists the first ones', () => {
    const result = compareValues(
      formulas([
        ['A1', '', 1],
        ['A2', '', 2],
        ['A3', '', '#NAME?'],
        ['A4', '', 'x'],
        ['A5', '', 5, 'uncached'],
        ['A6', '', 6, 'volatile'],
        ['A7', '', true],
        ['A8', '', null],
      ]),
      model({
        A1: { value: 1 },
        A2: { value: '#NAME? nope' },
        A3: { value: 3 },
        A4: {},
        A7: { value: false },
      }),
      false
    );
    expect(result).toMatchObject({
      status: 'ok',
      formulaCells: 8,
      compared: 6,
      matched: 2,
      mismatched: 4,
      uncached: 1,
      volatile: 1,
      unreadableSheets: 0,
    });
    expect(result.mismatchClasses).toEqual({
      'error -> number': 1,
      'text -> empty': 1,
      'boolean -> boolean': 1,
      'number -> error': 1,
    });
    expect(result.firstMismatches).toEqual([
      { sheet: 'S', cell: 'A2', expected: '2', actual: '"#NAME? nope"' },
      { sheet: 'S', cell: 'A3', expected: '"#NAME?"', actual: '3' },
      { sheet: 'S', cell: 'A4', expected: '"x"', actual: '(empty)' },
      { sheet: 'S', cell: 'A7', expected: 'TRUE', actual: 'FALSE' },
    ]);
  });

  it('caps the mismatch list but counts them all', () => {
    const cells = Array.from(
      { length: MAX_MISMATCHES + 5 },
      (_, i) => [`A${i + 1}`, '', i] as [string, string, number]
    );
    const result = compareValues(formulas(cells), model({}), false);
    expect(result.mismatched).toBe(MAX_MISMATCHES + 5); // an empty cell is not the number 0
    expect(result.firstMismatches).toHaveLength(MAX_MISMATCHES);
  });

  it('finds the right sheet among repeated names and tolerates a missing sheet or malformed model', () => {
    const dup: FileFormulas = {
      sheets: ['S', 'S', 'Other'],
      unreadableSheets: 1,
      cells: [
        { sheet: 'S', occurrence: 1, address: 'A1', expected: 2, skip: null },
        { sheet: 'Other', occurrence: 0, address: 'A1', expected: 1, skip: null },
      ],
    };
    const m = {
      sheets: [
        { name: 'S', cells: { A1: { value: 1 } } },
        { name: 'S', cells: { A1: { value: 2 } } },
        5,
      ],
    } as never;
    expect(compareValues(dup, m, false)).toMatchObject({
      compared: 2,
      matched: 1,
      mismatched: 1,
      unreadableSheets: 1,
    });
    expect(compareValues(dup, { sheets: 'no' } as never, false)).toMatchObject({
      compared: 2,
      matched: 0,
      mismatched: 2,
    });
    expect(compareValues(dup, 5 as never, false)).toMatchObject({ compared: 2, mismatched: 2 });
    expect(
      compareValues(dup, { sheets: [{ name: 'S', cells: 5 }] } as never, false).mismatched
    ).toBe(2);
  });

  it('records why nothing was compared', () => {
    const file = formulas([['A1', '', 1]]);
    expect(compareValues(file, model({}), true)).toMatchObject({
      status: 'skipped-import-error',
      formulaCells: 1,
      compared: 0,
      mismatched: 0,
    });
    expect(compareValues(file, null, false).status).toBe('skipped-import-error');
    expect(unreadableValues()).toMatchObject({
      status: 'unreadable-package',
      formulaCells: 0,
      compared: 0,
    });
  });
});
