import { DEFAULT_THEME_COLORS } from './colors';
import { decodeStringEscapes, forEachSharedString, MAX_RUN_ELEMENTS, parseSharedStrings } from './shared-strings';
import { MAIN_NS } from './test-support';

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

/** The one item of a part that holds `si`. */
const one = (si: string, colors = ctx) => parseSharedStrings(`<sst ${MAIN_NS}>${si}</sst>`, colors)[0];

describe('shared string items', () => {
  it('reads plain text, keeping spaces and line breaks as written', () => {
    expect(one('<si><t xml:space="preserve">  a\n b  </t></si>')).toEqual({ text: '  a\n b  ' });
    expect(one('<si><t/></si>')).toEqual({ text: '' });
    expect(one('<si/>')).toEqual({ text: '' });
  });
  it('keeps rich runs with their fonts, joins them into the text, and leaves out phonetic text', () => {
    const item = one(`<si><r><rPr><b/><sz val="11"/><color rgb="FFFF0000"/><rFont val="Calibri"/></rPr><t>Bold </t></r><r><t xml:space="preserve">plain</t></r><r><rPr/><t>!</t></r>
        <rPh sb="0" eb="1"><t>PHONETIC</t></rPh><phoneticPr fontId="1"/></si>`);
    expect(item).toEqual({
      text: 'Bold plain!',
      runs: [
        { text: 'Bold ', font: { bold: true, size: 11, color: { rgb: 'FFFF0000', resolved: '#ff0000' }, name: 'Calibri' } },
        { text: 'plain' },
        { text: '!' },
      ],
    });
  });
  it('ignores the direct text of an item that has runs, and a t inside a phonetic run of a plain item', () => {
    expect(one('<si><t>direct</t><r><t>run</t></r></si>')).toEqual({ text: 'run', runs: [{ text: 'run' }] });
    expect(one('<si><t>a</t><rPh><t>PH</t></rPh></si>')).toEqual({ text: 'a' });
  });
  it('splits a text over several <t> and decodes escapes', () => {
    expect(one('<si><t>a</t><t>_x000A_b</t></si>')).toEqual({ text: 'a\nb' });
  });
  it('resolves a theme colour in a run', () => {
    const run = one('<si><r><rPr><color theme="4"/></rPr><t>x</t></r></si>').runs?.[0];
    expect(run?.font?.color?.resolved).toBe('#4472c4');
  });
  it('decodes entities in text and CDATA, in plain text and in runs', () => {
    expect(one('<si><t>a &amp; b<![CDATA[ <c>]]></t></si>')).toEqual({ text: 'a & b <c>' });
    expect(one('<si><r><t>&lt;x&gt;_x0041_</t></r></si>')).toEqual({ text: '<x>A', runs: [{ text: '<x>A' }] });
  });
  it('reads any prefix, and an <r> or <t> that is not a direct child is not a run or text', () => {
    expect(one('<x:si><x:r><x:t>a</x:t></x:r></x:si>')).toEqual({ text: 'a', runs: [{ text: 'a' }] });
    expect(one('<si><t>a</t><x><t>b</t><r><t>c</t></r></x></si>')).toEqual({ text: 'a' });
  });
  it('reads a run only as far as MAX_RUN_ELEMENTS elements: what comes after them is left out (text before them is kept)', () => {
    const props = '<b/>'.repeat(MAX_RUN_ELEMENTS * 2);
    expect(one(`<si><r><rPr>${props}</rPr><t>late</t></r></si>`)).toEqual({ text: '', runs: [{ text: '', font: { bold: true } }] });
    expect(one(`<si><r><t>early</t><rPr>${props}</rPr></r></si>`)).toEqual({ text: 'early', runs: [{ text: 'early', font: { bold: true } }] });
  });
});

describe('parseSharedStrings / forEachSharedString', () => {
  const xml = `<sst ${MAIN_NS} count="3" uniqueCount="3"><si><t>a</t></si><junk/><si><r><t>b</t></r></si><si><t>c</t></si></sst>`;
  it('lists the items in order, skipping foreign elements', () => {
    expect(parseSharedStrings(xml, ctx).map((s) => s.text)).toEqual(['a', 'b', 'c']);
  });
  it('pushes items with their index to a callback', () => {
    const seen: Array<[number, string]> = [];
    forEachSharedString(xml, ctx, (i, item) => seen.push([i, item.text]));
    expect(seen).toEqual([[0, 'a'], [1, 'b'], [2, 'c']]);
  });
  it('is empty for a missing part', () => expect(parseSharedStrings(undefined, ctx)).toEqual([]));
  it('is empty, and not a problem, for a part with no items or no text at all', () => {
    const problems: string[] = [];
    expect(parseSharedStrings(`<sst ${MAIN_NS}/>`, ctx, { problem: (w) => problems.push(w) })).toEqual([]);
    expect(problems).toEqual([]);
  });

  it('stops at the first item over maxItems, once, and keeps the ones before it', () => {
    const problems: string[] = [];
    const items = parseSharedStrings(`<sst>${'<si><t>x</t></si>'.repeat(10)}</sst>`, ctx, { maxItems: 3, problem: (w) => problems.push(w) });
    expect(items).toHaveLength(3);
    expect(problems).toEqual(['items']);
  });
  it('is not over at exactly maxItems', () => {
    const problems: string[] = [];
    expect(parseSharedStrings(`<sst>${'<si><t>x</t></si>'.repeat(3)}</sst>`, ctx, { maxItems: 3, problem: (w) => problems.push(w) })).toHaveLength(3);
    expect(problems).toEqual([]);
  });
  it('joins the runs past maxRuns into the text of their item, without fonts, once', () => {
    const problems: string[] = [];
    const runs = ['A', 'B', 'C', 'D'].map((t) => `<r><rPr><b/></rPr><t>${t}</t></r>`).join('');
    const items = parseSharedStrings(`<sst><si>${runs}</si><si><r><t>E</t></r><r><t>F</t></r></si></sst>`, ctx, { maxRuns: 2, problem: (w) => problems.push(w) });
    expect(items[0]).toEqual({ text: 'ABCD', runs: [{ text: 'A', font: { bold: true } }, { text: 'B', font: { bold: true } }] });
    // every run of the second item is past the limit: it is plain text
    expect(items[1]).toEqual({ text: 'EF' });
    expect(problems).toEqual(['runs']);
  });
  it('reads what precedes damage (an item cut off by it is not reported) and says so', () => {
    const problems: string[] = [];
    const items = parseSharedStrings('<sst><si><t>a</t></si><si><t>b</t></Oops></sst>', ctx, { problem: (w) => problems.push(w) });
    expect(items.map((i) => i.text)).toEqual(['a']);
    expect(problems).toEqual(['damaged']);
  });
  it('never throws on damage without a problem handler', () => {
    expect(parseSharedStrings('<sst><si><t>a', ctx)).toEqual([]);
  });
});
