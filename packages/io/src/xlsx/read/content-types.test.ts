import { contentTypeOf, EMPTY_CONTENT_TYPES, parseContentTypes, partsWithContentType } from './content-types';

const NS = 'xmlns="http://schemas.openxmlformats.org/package/2006/content-types"';

describe('parseContentTypes', () => {
  const types = parseContentTypes(
    `<Types ${NS}>
      <Default Extension="XML" ContentType="application/xml"/>
      <Default Extension="xml" ContentType="ignored: the first default of an extension wins"/>
      <Default Extension="png" ContentType="image/png"/>
      <Default ContentType="no-extension"/>
      <Default Extension="bin"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="ignored"/>
      <Override PartName="xl/Styles.xml" ContentType="styles"/>
      <Override PartName="/x.xml"/>
      <Override ContentType="no-part"/>
    </Types>`,
  );
  it('keys defaults by lower-cased extension and overrides by lower-cased part name without the slash', () => {
    expect([...types.defaults]).toEqual([
      ['xml', 'application/xml'],
      ['png', 'image/png'],
    ]);
    expect([...types.overrides.keys()]).toEqual(['xl/workbook.xml', 'xl/styles.xml']);
  });
  it('answers an override first, then the extension default, case-insensitively', () => {
    expect(contentTypeOf(types, 'xl/workbook.xml')).toMatch(/sheet\.main\+xml$/);
    expect(contentTypeOf(types, 'XL/STYLES.XML')).toBe('styles');
    expect(contentTypeOf(types, 'xl/media/a.PNG')).toBe('image/png');
    expect(contentTypeOf(types, 'xl/media/a.gif')).toBeUndefined();
    expect(contentTypeOf(types, 'noextension')).toBeUndefined();
  });
  it('lists the parts whose type passes a test', () => {
    expect(partsWithContentType(types, (t) => t === 'styles')).toEqual(['xl/styles.xml']);
  });
  it('is empty for a missing part', () => {
    const none = parseContentTypes(undefined);
    expect(none.defaults.size + none.overrides.size).toBe(0);
    expect(EMPTY_CONTENT_TYPES.defaults.size).toBe(0);
  });
});

describe('parseContentTypes: problems and the entry limit', () => {
  const types = (n: number, extra = '') =>
    `<Types ${NS}><Default Extension="xml" ContentType="a"/>${Array.from({ length: n }, (_, i) => `<Override PartName="/p${i}.xml" ContentType="t"/>`).join('')}${extra}</Types>`;
  it('keeps the first entries up to the limit (defaults and overrides together) and reports it once', () => {
    const problems: string[] = [];
    const out = parseContentTypes(types(10), { max: 4, problem: (w) => problems.push(w) });
    expect(out.defaults.size + out.overrides.size).toBe(4);
    expect([...out.overrides.keys()]).toEqual(['p0.xml', 'p1.xml', 'p2.xml']);
    expect(problems).toEqual(['over']);
  });
  it('counts a default over the limit too', () => {
    const problems: string[] = [];
    const out = parseContentTypes(`<Types ${NS}><Override PartName="/a.xml" ContentType="t"/><Default Extension="png" ContentType="i"/></Types>`, { max: 1, problem: (w) => problems.push(w) });
    expect(out.defaults.size).toBe(0);
    expect(problems).toEqual(['over']);
  });
  it('is not over at exactly the limit, and repeats do not count', () => {
    const problems: string[] = [];
    const text = types(3, '<Default Extension="XML" ContentType="b"/><Override PartName="P0.xml" ContentType="b"/>');
    expect(parseContentTypes(text, { max: 4, problem: (w) => problems.push(w) }).overrides.size).toBe(3);
    expect(problems).toEqual([]);
  });
  it('reads what precedes damage and reports it', () => {
    const problems: string[] = [];
    const out = parseContentTypes(`<Types ${NS}><Override PartName="/a.xml" ContentType="t"/><Override`, { problem: (w) => problems.push(w) });
    expect([...out.overrides.keys()]).toEqual(['a.xml']);
    expect(problems).toEqual(['damaged']);
  });
  it('ignores elements that are not direct children', () => {
    expect(parseContentTypes(`<Types ${NS}><x><Default Extension="a" ContentType="t"/></x><Other Extension="b" ContentType="t"/></Types>`).defaults.size).toBe(0);
  });
});
