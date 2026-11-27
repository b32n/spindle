import { DEFAULT_THEME_COLORS } from './colors';
import { decodeStringEscapes, forEachSharedString, parseSharedStrings, readSharedString } from './shared-strings';
import { el, MAIN_NS } from './test-support';

const ctx = { themeColors: DEFAULT_THEME_COLORS };

describe('decodeStringEscapes', () => {
  it.each([
    ['a_x000D_b', 'a\rb'],
    ['_x0041_', 'A'],
    ['_x005F_x0041_', '_x0041_'], // the escape for a literal underscore guards the text that follows
    ['_x00zz_', '_x00zz_'],
    ['_xD800_', '_xD800_'], // a lone surrogate half is not decodable
    ['plain_x', 'plain_x'],
    ['no escapes', 'no escapes'],
  ])('%s -> %s', (input, expected) => expect(decodeStringEscapes(input)).toBe(expected));
});

describe('readSharedString', () => {
  it('reads plain text, keeping spaces and line breaks as written', () => {
    expect(readSharedString(el('<si><t xml:space="preserve">  a\n b  </t></si>'), ctx)).toEqual({ text: '  a\n b  ' });
    expect(readSharedString(el('<si><t/></si>'), ctx)).toEqual({ text: '' });
    expect(readSharedString(el('<si/>'), ctx)).toEqual({ text: '' });
  });
  it('keeps rich runs with their fonts, joins them into the text, and leaves out phonetic text', () => {
    const item = readSharedString(
      el(`<si><r><rPr><b/><sz val="11"/><color rgb="FFFF0000"/><rFont val="Calibri"/></rPr><t>Bold </t></r><r><t xml:space="preserve">plain</t></r><r><rPr/><t>!</t></r>
        <rPh sb="0" eb="1"><t>PHONETIC</t></rPh><phoneticPr fontId="1"/></si>`),
      ctx,
    );
    expect(item).toEqual({
      text: 'Bold plain!',
      runs: [
        { text: 'Bold ', font: { bold: true, size: 11, color: { rgb: 'FFFF0000', resolved: '#ff0000' }, name: 'Calibri' } },
        { text: 'plain' },
        { text: '!' },
      ],
    });
  });
  it('splits a text over several <t> and decodes escapes', () => {
    expect(readSharedString(el('<si><t>a</t><t>_x000A_b</t></si>'), ctx)).toEqual({ text: 'a\nb' });
  });
  it('resolves a theme colour in a run', () => {
    const run = readSharedString(el('<si><r><rPr><color theme="4"/></rPr><t>x</t></r></si>'), ctx).runs?.[0];
    expect(run?.font?.color?.resolved).toBe('#4472c4');
  });
});

describe('parseSharedStrings / forEachSharedString', () => {
  const root = el(`<sst ${MAIN_NS} count="3" uniqueCount="3"><si><t>a</t></si><junk/><si><r><t>b</t></r></si><si><t>c</t></si></sst>`);
  it('lists the items in order, skipping foreign elements', () => {
    expect(parseSharedStrings(root, ctx).map((s) => s.text)).toEqual(['a', 'b', 'c']);
  });
  it('pushes items with their index to a callback', () => {
    const seen: Array<[number, string]> = [];
    forEachSharedString(root, ctx, (i, item) => seen.push([i, item.text]));
    expect(seen).toEqual([[0, 'a'], [1, 'b'], [2, 'c']]);
  });
  it('is empty for a missing part', () => expect(parseSharedStrings(undefined, ctx)).toEqual([]));
});
