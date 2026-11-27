import JSZip from 'jszip';
import { importXlsx } from '../reader';
import { parseXml } from '../xml';
import { MAIN_NS, REL_NS, relsXml } from './test-support';

// ---------------------------------------------------------------------------
// The workbook-level parts (styles, shared strings, workbook, theme, relationships, content types) are read as element trees, and
// each is walked by the readers of this directory. A hostile file can make any of them a few hundred thousand items long for a few
// hundred KB of zip: every shape below must cost time that grows with its size (each is timed, with a limit well above the cost it
// has now (a loaded machine must not fail it) and far below what a quadratic walk would take). The shapes of the sheet part itself are in `reader-hostile.test.ts`.
// ---------------------------------------------------------------------------

const N = 300_000;
const sheet = (cells: string) => `<worksheet ${MAIN_NS}><sheetData>${cells}</sheetData></worksheet>`;
const rep = (n: number, f: (i: number) => string): string => Array.from({ length: n }, (_, i) => f(i)).join('');

/** A one-sheet package whose parts are the defaults below, replaced by `parts`; compressed, as a hostile file would be. */
async function pack(parts: Record<string, string>): Promise<Uint8Array> {
  const defaults: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
    '_rels/.rels': relsXml([['rId1', 'officeDocument', 'xl/workbook.xml']]),
    'xl/workbook.xml': `<workbook ${MAIN_NS} ${REL_NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': relsXml([['rId1', 'worksheet', 'worksheets/sheet1.xml']]),
    'xl/worksheets/sheet1.xml': sheet('<row r="1"><c r="A1"><v>1</v></c></row>'),
  };
  const zip = new JSZip();
  for (const [name, text] of Object.entries({ ...defaults, ...parts })) zip.file(name, text);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 1 } });
}

const timed = async <T>(run: () => Promise<T>): Promise<{ value: T; ms: number }> => {
  const t0 = Date.now();
  const value = await run();
  return { value, ms: Date.now() - t0 };
};

describe('hostile workbook-level parts: each costs time in proportion to its size', () => {
  type Result = Awaited<ReturnType<typeof importXlsx>>;
  const shapes: Array<[name: string, parts: () => Record<string, string>, check: (r: Result) => void]> = [
    [
      '300,000 number formats and 300,000 cells that each use their own format (a table scan per format was quadratic)',
      () => ({
        'xl/styles.xml': `<styleSheet ${MAIN_NS}><numFmts>${rep(N, (i) => `<numFmt numFmtId="${200 + i}" formatCode="0.${i}"/>`)}</numFmts><cellXfs>${rep(N, (i) => `<xf numFmtId="${i % 2 ? 14 : 2}"/>`)}</cellXfs></styleSheet>`,
        'xl/worksheets/sheet1.xml': sheet(rep(N, (i) => `<row r="${i + 1}"><c r="A${i + 1}" s="${i}"><v>1</v></c></row>`)),
      }),
      (r) => expect(r.sheets[0].cells.size).toBe(N),
    ],
    ['300,000 shared strings', () => ({ 'xl/sharedStrings.xml': `<sst ${MAIN_NS}>${rep(N, (i) => `<si><t>${i}</t></si>`)}</sst>` }), (r) => expect(r.sharedStrings).toHaveLength(N)],
    ['one shared string of 100,000 rich runs', () => ({ 'xl/sharedStrings.xml': `<sst ${MAIN_NS}><si>${rep(100_000, () => '<r><rPr><b/></rPr><t>x</t></r>')}</si></sst>` }), (r) => expect(r.sharedStrings?.[0].runs).toHaveLength(100_000)],
    ['300,000 defined names', () => ({ 'xl/workbook.xml': `<workbook ${MAIN_NS} ${REL_NS}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets><definedNames>${rep(N, (i) => `<definedName name="n${i}">A1</definedName>`)}</definedNames></workbook>` }), (r) => expect(r.workbook?.definedNames).toHaveLength(N)],
    [
      '300,000 workbook relationships',
      () => ({ 'xl/_rels/workbook.xml.rels': relsXml([['rId1', 'worksheet', 'worksheets/sheet1.xml'], ...Array.from({ length: N }, (_, i) => [`x${i}`, 'styles', `s${i}.xml`] as [string, string, string])]) }),
      (r) => expect(r.sheets).toHaveLength(1),
    ],
    [
      '300,000 content type overrides of the kind a part is looked up by',
      () => ({ '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/>${rep(N, (i) => `<Override PartName="/xl/s${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`)}</Types>` }),
      (r) => expect(r.sheets).toHaveLength(1),
    ],
    [
      '300,000 colours in a theme scheme, 300,000 palette entries and 300,000 fonts',
      () => ({
        'xl/theme/theme1.xml': `<a:theme xmlns:a="urn:a"><a:themeElements><a:clrScheme>${rep(N, (i) => `<a:x${i % 5}/>`)}</a:clrScheme></a:themeElements></a:theme>`,
        'xl/styles.xml': `<styleSheet ${MAIN_NS}><colors><indexedColors>${rep(N, () => '<rgbColor rgb="FF0000FF"/>')}</indexedColors></colors><fonts>${rep(N, (i) => `<font><color indexed="${i}"/></font>`)}</fonts></styleSheet>`,
      }),
      (r) => expect(r.styles?.fonts).toHaveLength(N),
    ],
    [
      '100,000 fills, borders, named styles and dxfs (800,000 elements, under the budget of a million)',
      () => ({
        'xl/styles.xml': `<styleSheet ${MAIN_NS}><fills>${rep(N / 3, () => '<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>')}</fills><borders>${rep(N / 3, () => '<border><left style="thin"/></border>')}</borders><cellStyles>${rep(N / 3, (i) => `<cellStyle name="n${i}" xfId="0"/>`)}</cellStyles><dxfs>${rep(N / 3, () => '<dxf><font><b/></font></dxf>')}</dxfs></styleSheet>`,
      }),
      (r) => expect([r.styles?.fills.length, r.styles?.borders.length, r.styles?.cellStyles.length, r.styles?.dxfs.length]).toEqual([N / 3, N / 3, N / 3, N / 3]),
    ],
  ];

  it.each(shapes)('%s', async (_name, parts, check) => {
    const bytes = await pack(parts());
    const { value, ms } = await timed(() => importXlsx(bytes));
    check(value);
    expect(ms).toBeLessThan(60_000);
  }, 180_000);
});

describe('hostile XML in the element tree reader (parseXml): the shapes of the sax scanner, a million items each', () => {
  const N1 = 1_000_000;
  const open = '<a xmlns="x">';
  const shapes: Array<[string, () => string]> = [
    ['processing instructions', () => `${open}${'<?x?>'.repeat(N1)}</a>`],
    ['comments', () => `${open}${'<!---->'.repeat(N1)}</a>`],
    ['CDATA sections', () => `${open}${'<![CDATA[]]>'.repeat(N1)}</a>`],
    ['declarations', () => `${open}${'<!a>'.repeat(N1)}</a>`],
    ['unterminated comments', () => `${open}${'<!--'.repeat(N1)}`],
    ['unterminated processing instructions', () => `${open}${'<?x'.repeat(N1)}`],
    ['unterminated CDATA', () => `${open}${'<![CDATA['.repeat(N1)}`],
    ['attributes on one element', () => `${open}<b ${'c="1" '.repeat(N1)}/></a>`],
    ['attributes without a value', () => `${open}<b ${'c '.repeat(N1)}/></a>`],
    ['`xmlns` declarations on one element', () => `${open}<b ${Array.from({ length: 200_000 }, (_, i) => `xmlns:p${i}="u"`).join(' ')}/></a>`],
    ['entities in text', () => `${open}${'&amp;'.repeat(2_000_000)}</a>`],
    ['entities in an attribute value', () => `${open}<b c="${'&#65;'.repeat(2_000_000)}"/></a>`],
    ['an attribute value of 20 million characters', () => `${open}<b c="${'x'.repeat(20_000_000)}"/></a>`],
    ['nested elements', () => `${open}${'<b>'.repeat(N1)}`],
    ['stray end tags', () => `${open}${'</b>'.repeat(N1)}</a>`],
    ['`]` after a DOCTYPE subset opener', () => `<!DOCTYPE a [${']'.repeat(N1)}${open}</a>`],
    ['`]` and blanks after a DOCTYPE subset opener', () => `<!DOCTYPE a [${']  '.repeat(N1)}${open}</a>`],
    ['DOCTYPE openers', () => `${'<!DOCTYPE a ['.repeat(N1)}${open}</a>`],
    ['blanks in a tag', () => `${open}<b${' '.repeat(4_000_000)}/></a>`],
  ];

  it.each(shapes)('%s', (_name, xml) => {
    const text = xml();
    const t0 = Date.now();
    parseXml(text);
    expect(Date.now() - t0).toBeLessThan(30_000);
  }, 120_000);
});
