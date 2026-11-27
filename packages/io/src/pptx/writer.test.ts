import { readFileSync, readdirSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import type { PptxLayout, PptxPlaceholder, PptxSlide, PresentationImportResult } from '@b32nio/spindle-shared/export/pptx';
import { DEFAULT_COLOR_MAP } from './color';
import { importPptx } from './reader';
import { FALLBACK_THEME } from './theme';
import { exportPptx } from './writer';
import { lintPptxPackage } from './writer-lint';

const FIXTURES = path.resolve(__dirname, '../../../../fixtures/pptx');
const fixture = (name: string): Uint8Array => new Uint8Array(readFileSync(path.join(FIXTURES, name)));

const level = (extra: Record<string, unknown> = {}) => ({ fontSize: 24, ...extra });
const placeholder = (over: Partial<PptxPlaceholder> = {}): PptxPlaceholder => ({
  type: 'title',
  frame: { x: 10, y: 20, w: 300, h: 80 },
  levels: [level(), {}, {}, {}, {}, {}, {}, {}, {}],
  ...over,
});
const layout = (over: Partial<PptxLayout> = {}): PptxLayout => ({
  partName: 'ppt/slideLayouts/slideLayout1.xml',
  name: 'Title',
  masterPart: 'ppt/slideMasters/slideMaster1.xml',
  placeholders: [placeholder()],
  elements: [],
  staticShapes: 0,
  ...over,
});
const slide = (over: Partial<PptxSlide> = {}): PptxSlide => ({ pptxId: '256', partName: 'ppt/slides/slide1.xml', layoutPart: 'ppt/slideLayouts/slideLayout1.xml', ...over });
const deck = (over: Partial<PresentationImportResult> = {}): PresentationImportResult => ({
  slideSize: { w: 1280, h: 720 },
  theme: FALLBACK_THEME,
  colorMap: DEFAULT_COLOR_MAP,
  slides: [slide()],
  layouts: [layout()],
  fonts: { used: [], unavailable: [] },
  opaqueParts: [],
  warnings: [],
  ...over,
});

const read = async (bytes: Uint8Array) => {
  const zip = await JSZip.loadAsync(bytes);
  return { zip, text: (name: string) => zip.file(name)?.async('string') as Promise<string> };
};

describe('exportPptx: the package', () => {
  it('writes a package that passes every lint and that the reader opens again', async () => {
    const { bytes, warnings } = await exportPptx(deck({ slides: [slide(), slide({ pptxId: '300', name: 'Second' })] }));
    expect(warnings).toEqual([]);
    expect(await lintPptxPackage(bytes)).toEqual([]);
    const back = await importPptx(bytes);
    expect(back.slideSize).toEqual({ w: 1280, h: 720 });
    expect(back.slides.map((s) => [s.pptxId, s.name])).toEqual([['256', undefined], ['300', 'Second']]);
    expect(back.layouts[0]).toMatchObject({ name: 'Title', placeholders: [{ type: 'title', frame: { x: 10, y: 20, w: 300, h: 80 } }] });
    expect(back.theme).toEqual(FALLBACK_THEME);
    expect(back.warnings).toEqual([]);
  });

  it('is deterministic: the same deck gives the same bytes, and the entries carry the fixed zip date', async () => {
    const a = (await exportPptx(deck())).bytes;
    const b = (await exportPptx(deck())).bytes;
    expect(Buffer.from(a).equals(Buffer.from(b))).toBe(true);
    const { zip } = await read(a);
    expect(Object.keys(zip.files)[0]).toBe('[Content_Types].xml');
    // the DOS date of the first entry's local header (offset 12): 1980-01-01
    expect(Buffer.from(a).readUInt16LE(12)).toBe(0x21);
  });

  it('claims no producer it is not: the application is Spindle, and the clock only appears when the host passes one', async () => {
    const plain = await read((await exportPptx(deck())).bytes);
    expect(await plain.text('docProps/app.xml')).toContain('<Application>Spindle</Application>');
    expect(await plain.text('docProps/app.xml')).not.toMatch(/Microsoft|PowerPoint/);
    expect(await plain.text('docProps/core.xml')).not.toContain('dcterms:');
    const meta = await read((await exportPptx(deck(), { metadata: { title: 'A & B', creator: 'Ada <a>', description: 'd', modified: new Date('2026-03-04T05:06:07.890Z') } })).bytes);
    const core = await meta.text('docProps/core.xml');
    expect(core).toContain('<dc:title>A &amp; B</dc:title>');
    expect(core).toContain('<dc:creator>Ada &lt;a&gt;</dc:creator>');
    expect(core).toContain('<dc:description>d</dc:description>');
    expect(core).toContain('<dcterms:created xsi:type="dcterms:W3CDTF">2026-03-04T05:06:07Z</dcterms:created>');
    expect(core).toContain('<dcterms:modified xsi:type="dcterms:W3CDTF">2026-03-04T05:06:07Z</dcterms:modified>');
  });

  it('sets the slide size in EMU, the hidden flag and the slide name, and escapes names', async () => {
    const out = await read((await exportPptx(deck({ slideSize: { w: 960, h: 720 }, slides: [slide({ hidden: true, name: 'A "q" <b> & c\u0001' })] }))).bytes);
    expect(await out.text('ppt/presentation.xml')).toContain('<p:sldSz cx="9144000" cy="6858000"/>');
    const s = await out.text('ppt/slides/slide1.xml');
    expect(s).toContain('show="0"');
    expect(s).toContain('name="A &quot;q&quot; &lt;b&gt; &amp; c"');
    const back = await importPptx((await exportPptx(deck({ slides: [slide({ hidden: true, name: 'x' })] }))).bytes);
    expect(back.slides[0]).toMatchObject({ hidden: true, name: 'x' });
  });

  it('keeps usable slide ids and allocates the rest: no id twice, none below 256, none above the limit', async () => {
    const ids = ['256', '256', 'abc', '10', '2147483648', '400', '0257'];
    const { bytes } = await exportPptx(deck({ slides: ids.map((id) => slide({ pptxId: id })) }));
    const back = await importPptx(bytes);
    // 256 kept; the second 256 -> next free (257 is taken by '0257' which is kept as 257 only if it comes first: it comes later) ...
    const got = back.slides.map((s) => s.pptxId);
    expect(new Set(got).size).toBe(got.length);
    expect(got.every((id) => Number(id) >= 256 && Number(id) <= 2147483647)).toBe(true);
    expect(got[0]).toBe('256');
    expect(got[5]).toBe('400');
    expect(await lintPptxPackage(bytes)).toEqual([]);
  });

  it('writes sections with the final slide ids and leaves out slides the deck does not have', async () => {
    const { bytes } = await exportPptx(
      deck({ slides: [slide({ pptxId: '256' }), slide({ pptxId: '5' }), slide({ pptxId: '300' })], sections: [{ name: 'One & two', slideIds: ['256', '5', '999'] }, { name: 'Three', slideIds: ['300'] }, { name: 'Empty', slideIds: [] }] })
    );
    const back = await importPptx(bytes);
    const final = back.slides.map((s) => s.pptxId);
    expect(back.sections).toEqual([
      { name: 'One & two', slideIds: [final[0], final[1]] },
      { name: 'Three', slideIds: [final[2]] },
      { name: 'Empty', slideIds: [] },
    ]);
  });

  it('writes a blank layout when the deck has none, and gives a slide with an unknown layout the first one', async () => {
    const none = await exportPptx(deck({ layouts: [], slides: [slide({ layoutPart: undefined })] }));
    expect(await lintPptxPackage(none.bytes)).toEqual([]);
    expect((await importPptx(none.bytes)).layouts).toMatchObject([{ name: 'Blank', layoutType: 'blank', placeholders: [] }]);
    const odd = await exportPptx(deck({ layouts: [layout({ partName: 'x/y.xml', name: 'First' }), layout({ partName: 'ppt/slideLayouts/slideLayout1.xml', name: 'Second' })], slides: [slide({ layoutPart: 'nowhere.xml' }), slide({ layoutPart: 'ppt/slideLayouts/slideLayout1.xml' })] }));
    expect(await lintPptxPackage(odd.bytes)).toEqual([]);
    const back = await importPptx(odd.bytes);
    // the layout whose part name is not the usual one is renumbered, the other keeps its own
    expect(back.layouts.map((l) => [l.partName, l.name])).toEqual([['ppt/slideLayouts/slideLayout2.xml', 'First'], ['ppt/slideLayouts/slideLayout1.xml', 'Second']]);
    expect(back.slides.map((s) => s.layoutPart)).toEqual(['ppt/slideLayouts/slideLayout2.xml', 'ppt/slideLayouts/slideLayout1.xml']);
  });

  it('renumbers two layouts that claim the same part name', async () => {
    const { bytes } = await exportPptx(deck({ layouts: [layout({ name: 'A' }), layout({ name: 'B' }), layout({ partName: 'ppt/slideLayouts/slideLayout2.xml', name: 'C' })] }));
    const back = await importPptx(bytes);
    expect(back.layouts.map((l) => l.name)).toEqual(['A', 'B', 'C']);
    expect(new Set(back.layouts.map((l) => l.partName)).size).toBe(3);
    expect(await lintPptxPackage(bytes)).toEqual([]);
  });

  it('keeps the first of two layouts that claim the same part name for the slides that name it', async () => {
    const dup = 'ppt/slideLayouts/slideLayout1.xml';
    const { bytes } = await exportPptx(deck({ layouts: [layout({ name: 'A' }), layout({ name: 'B' })], slides: [slide({ layoutPart: dup })] }));
    const back = await importPptx(bytes);
    expect(back.layouts.map((l) => l.name)).toEqual(['A', 'B']);
    expect(back.slides[0].layoutPart).toBe(back.layouts[0].partName);
  });

  it('keeps the first slide of a repeated source id in the sections', async () => {
    const { bytes } = await exportPptx(deck({ slides: [slide({ name: 'One' }), slide({ name: 'Two', partName: 'ppt/slides/slide2.xml' })], sections: [{ name: 'S', slideIds: ['256'] }] }));
    const back = await importPptx(bytes);
    expect(back.sections).toEqual([{ name: 'S', slideIds: [back.slides[0].pptxId] }]);
    expect(back.slides[1].pptxId).not.toBe(back.slides[0].pptxId);
  });

  it('keeps the colour map and the theme with its east Asian and complex fonts', async () => {
    const colorMap = { ...DEFAULT_COLOR_MAP, bg1: 'dk1' as const, tx1: 'lt1' as const };
    const theme = { ...FALLBACK_THEME, name: 'Mine & "yours"', fonts: { major: 'Georgia', minor: 'Verdana', majorEastAsian: 'MS Mincho', minorEastAsian: 'SimSun', majorComplex: 'Arial', minorComplex: 'Tahoma' } };
    const back = await importPptx((await exportPptx(deck({ colorMap, theme }))).bytes);
    expect(back.colorMap).toEqual(colorMap);
    expect(back.theme).toEqual(theme);
  });
});

describe('exportPptx: layouts', () => {
  it('states a placeholder in full: levels, body properties, anchor, prompt, type, idx and name', async () => {
    const ph: PptxPlaceholder[] = [
      {
        type: 'body',
        idx: 1,
        name: 'Content',
        frame: { x: 1, y: 2, w: 3, h: 4, rot: 10, flipH: true, flipV: true },
        anchor: 'middle',
        bodyProps: { insets: { l: 1, t: 2, r: 3, b: 4 }, anchorCtr: true, vert: 'vert', wrap: 'none', columns: 2, columnSpacing: 12, rot: 5, autofit: { kind: 'shrink', fontScale: 0.625, lnSpcReduction: 0.2 } },
        levels: [
          level({ bold: true, italic: false, color: { kind: 'theme', slot: 'accent1', lumMod: 0.5, lumOff: 0.25, alpha: 0.5 }, font: 'major', align: 'justify', marginLeft: 20, indent: -10, lineHeight: 1.2, spaceBefore: 6, spaceAfter: 3, bullet: { kind: 'char', char: '•' }, bulletColor: { kind: 'rgb', hex: '#FF0000' }, bulletSizePct: 0.8, bulletFont: 'Arial', tabStops: [{ pos: 50, align: 'l' }, { pos: 99, align: 'dec' }] }),
          level({ font: 'minor', align: 'center', lineSpacingPts: 14, spaceBeforePct: 0.2, spaceAfterPct: 0.1, bullet: { kind: 'auto', scheme: 'arabicPeriod', startAt: 3 } }),
          level({ font: 'Calibri', align: 'right', bullet: { kind: 'none' } }),
          { align: 'left' },
          {}, {}, {}, {}, {},
        ],
        prompt: 'Click & edit\n\nsecond <line>',
      },
      { type: 'obj', levels: [] },
      { type: 'sldNum', idx: 12, levels: [{}, {}, {}, {}, {}, {}, {}, {}, {}], bodyProps: { autofit: { kind: 'none' } } },
      { type: 'dt', levels: [], bodyProps: { autofit: { kind: 'resize' } } },
      { type: 'ftr', levels: [], bodyProps: { autofit: { kind: 'shrink' } }, anchor: 'bottom' },
      { type: 'pic', levels: [], bodyProps: { anchor: 'top' } },
    ];
    const { bytes } = await exportPptx(deck({ layouts: [layout({ placeholders: ph })] }));
    expect(await lintPptxPackage(bytes)).toEqual([]);
    const back = (await importPptx(bytes)).layouts[0].placeholders;
    // a placeholder's own anchor reads back as a body property too
    expect(back[0]).toEqual({ ...ph[0], bodyProps: { ...ph[0].bodyProps, anchor: 'middle' } });
    // the unnamed ones get a name; the rest is as stated
    expect(back[1]).toMatchObject({ type: 'obj' });
    expect(back[1].levels).toEqual(Array.from({ length: 9 }, () => ({})));
    expect(back[2]).toMatchObject({ type: 'sldNum', idx: 12, bodyProps: { autofit: { kind: 'none' } } });
    expect(back[3].bodyProps).toEqual({ autofit: { kind: 'resize' } });
    expect(back[4]).toMatchObject({ anchor: 'bottom', bodyProps: { autofit: { kind: 'shrink' }, anchor: 'bottom' } });
    expect(back[5]).toMatchObject({ type: 'pic', anchor: 'top' });
  });

  it('keeps the layout type, the hidden master shapes and a solid background, and drops a layout type the schema does not know', async () => {
    const { bytes } = await exportPptx(
      deck({ layouts: [layout({ layoutType: 'titleOnly', hideMasterShapes: true, background: { kind: 'theme', slot: 'accent2', lumMod: 0.6 }, masterName: 'M & N' }), layout({ partName: 'ppt/slideLayouts/slideLayout2.xml', layoutType: 'bogus' })] })
    );
    const back = (await importPptx(bytes)).layouts;
    expect(back[0]).toMatchObject({ layoutType: 'titleOnly', hideMasterShapes: true, background: { kind: 'theme', slot: 'accent2', lumMod: 0.6 }, masterName: 'M & N' });
    expect(back[1].layoutType).toBeUndefined();
    expect(back[1].background).toBeUndefined();
    expect(back[1].masterName).toBe('M & N');
  });

  it('puts the five standard placeholders on the master, scaled to the slide', async () => {
    const out = await read((await exportPptx(deck({ slideSize: { w: 960, h: 720 } }))).bytes);
    const master = await out.text('ppt/slideMasters/slideMaster1.xml');
    expect(master.match(/<p:ph /g)).toHaveLength(5);
    expect(master).toContain('<p:sldLayoutId id="2147483649" r:id="rId1"/>');
    expect(master).toContain('<p:clrMap bg1="lt1" tx1="dk1"');
  });
});

describe('exportPptx: what is not exported yet', () => {
  const shape = { kind: 'shape' as const, pptxId: '2', name: 's' };
  it('says so for every kind the writer leaves out, so a deck is never reported clean when it was emptied', async () => {
    const slides = [
      slide({
        elements: [
          shape,
          { kind: 'picture', pptxId: '3', name: 'p' },
          { kind: 'table', pptxId: '4', name: 't', columns: [], rows: [] },
          { kind: 'connector', pptxId: '5', name: 'c' },
          { kind: 'group', pptxId: '6', name: 'g', children: [shape, { kind: 'opaque', opaqueKind: 'chart', pptxId: '7', name: 'o', parts: [] }] },
        ],
        background: { kind: 'none' },
        notes: { props: {}, paragraphs: [] },
        threads: [{ id: '1', format: 'legacy', anchor: { kind: 'slide' }, comments: [] }, { id: '2', format: 'modern', anchor: { kind: 'slide' }, comments: [] }],
        opaque: { transition: '<p:transition/>', timing: '<p:timing/>' },
      }),
      slide({ pptxId: '257', error: 'the part is damaged' }),
    ];
    const { warnings } = await exportPptx(deck({ slides, layouts: [layout({ elements: [shape] })] }));
    expect(warnings).toEqual([
      '1 slide could not be read at import and is written empty.',
      'Not exported yet: 2 shapes and text boxes (K12b).',
      'Not exported yet: 1 pictures (K13).',
      'Not exported yet: 1 tables (K13).',
      'Not exported yet: 1 connectors (K13).',
      'Not exported yet: 1 groups (K13).',
      'Not exported yet: 1 charts, diagrams, objects, media and ink (K14).',
      'Not exported yet: 1 slide backgrounds (K12b).',
      'Not exported yet: 1 speaker notes (K13).',
      'Not exported yet: 2 comment threads (K13).',
      'Not exported yet: 1 transitions (K14).',
      'Not exported yet: 1 animation timelines (K14).',
      'Not exported yet: 1 static layout shape (K12b).',
    ]);
    const two = await exportPptx(deck({ slides: [slide({ error: 'e' }), slide({ pptxId: '257', error: 'e' })], layouts: [layout({ elements: [shape, shape] })] }));
    expect(two.warnings).toEqual(['2 slides could not be read at import and are written empty.', 'Not exported yet: 2 static layout shapes (K12b).']);
  });
});

describe('exportPptx: several masters', () => {
  it('says that the masters were merged', async () => {
    const { warnings } = await exportPptx(deck({ layouts: [layout({ masterPart: 'ppt/slideMasters/slideMaster1.xml' }), layout({ partName: 'ppt/slideLayouts/slideLayout2.xml', masterPart: 'ppt/slideMasters/slideMaster2.xml' })] }));
    expect(warnings).toEqual(['2 slide masters were merged into one (their layouts keep their placeholders).']);
    expect((await exportPptx(deck())).warnings).toEqual([]);
  });
});

describe('exportPptx over the corpus', () => {
  const decks = readdirSync(FIXTURES).filter((f) => /^(?!edge-(?:encrypted|legacy|no-presentation|truncated)).*\.(pptx|pptm)$/.test(f));

  it.each(decks)('%s: the written package passes every lint and reads back with the same slides, layout names and theme', async (name) => {
    const first = await importPptx(fixture(name));
    const { bytes } = await exportPptx(first);
    expect(await lintPptxPackage(bytes)).toEqual([]);
    const second = await importPptx(bytes);
    expect(second.slides).toHaveLength(first.slides.length);
    expect(second.slides.map((s) => [s.hidden, s.name])).toEqual(first.slides.map((s) => [s.hidden, s.name]));
    expect(second.slideSize).toEqual(first.slideSize);
    expect(second.theme).toEqual(first.theme);
    expect(second.colorMap).toEqual(first.colorMap);
    expect(second.sections).toEqual(first.sections?.map((s) => ({ ...s, slideIds: s.slideIds.map((id) => second.slides[first.slides.findIndex((x) => x.pptxId === id)].pptxId) })));
    // every layout comes back with the same placeholders (the master of origin is the one thing a single-master deck loses)
    expect(second.layouts.map((l) => l.placeholders.map((p) => ({ ...p, name: p.name })))).toEqual(first.layouts.map((l) => l.placeholders));
  });
});
