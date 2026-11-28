import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import type { PptxElement, PptxFill, PptxLayout, PptxLevelStyle, PptxParagraph, PptxRun, PptxShape, PptxSlide, PresentationImportResult } from '@b32nio/spindle-shared/export/pptx';
import { DEFAULT_COLOR_MAP } from './color';
import { importPptx } from './reader';
import { FALLBACK_THEME } from './theme';
import { exportPptx } from './writer';
import { lintPptxPackage } from './writer-lint';
import { sharedLevels } from './writer-body';
import { allocateElementIds } from './writer-shapes';

const FIXTURES = path.resolve(__dirname, '../../../../fixtures/pptx');
const fixture = (name: string): Uint8Array => new Uint8Array(readFileSync(path.join(FIXTURES, name)));

const layout = (over: Partial<PptxLayout> = {}): PptxLayout => ({ partName: 'ppt/slideLayouts/slideLayout1.xml', name: 'Blank', layoutType: 'blank', masterPart: 'ppt/slideMasters/slideMaster1.xml', placeholders: [], elements: [], staticShapes: 0, ...over });
const slide = (over: Partial<PptxSlide> = {}): PptxSlide => ({ pptxId: '256', partName: 'ppt/slides/slide1.xml', layoutPart: 'ppt/slideLayouts/slideLayout1.xml', ...over });
const deck = (slides: PptxSlide[], layouts: PptxLayout[] = [layout()]): PresentationImportResult => ({
  slideSize: { w: 1280, h: 720 },
  theme: FALLBACK_THEME,
  colorMap: DEFAULT_COLOR_MAP,
  slides,
  layouts,
  fonts: { used: [], unavailable: [] },
  opaqueParts: [],
  warnings: [],
});
const shape = (over: Partial<PptxShape> = {}): PptxShape => ({ kind: 'shape', pptxId: '2', name: 'Shape', frame: { x: 10, y: 20, w: 300, h: 100 }, ...over });
const para = (runs: PptxRun[], style: PptxLevelStyle = {}, over: Partial<PptxParagraph> = {}): PptxParagraph => ({ level: 0, style, runs, ...over });
const text = (t: string, over: Partial<PptxRun> = {}): PptxRun => ({ kind: 'text', text: t, ...over });
const rgb = (hex: string, alpha?: number) => ({ kind: 'rgb' as const, hex, ...(alpha !== undefined ? { alpha } : {}) });

/** Write one slide holding `elements`; returns its XML, its relationships, the warnings and the model read back. */
async function write(elements: PptxElement[], extra: Partial<PptxSlide> = {}, layouts?: PptxLayout[]) {
  const { bytes, warnings } = await exportPptx(deck([slide({ elements, ...extra })], layouts));
  expect(await lintPptxPackage(bytes)).toEqual([]);
  const zip = await JSZip.loadAsync(bytes);
  const part = async (n: string) => (await zip.file(n)?.async('string')) as string;
  const back = await importPptx(bytes);
  return { xml: await part('ppt/slides/slide1.xml'), rels: await part('ppt/slides/_rels/slide1.xml.rels'), part, warnings, back, shapes: (back.slides[0].elements ?? []) as PptxShape[] };
}

describe('slide shapes', () => {
  it('writes a shape with a preset, adjustments, fill and outline that reads back equal', async () => {
    const s = shape({
      name: 'Box & "co"',
      altText: 'alt <text>',
      geometry: { kind: 'preset', prst: 'roundRect', adjustments: { adj: 12500.4 }, known: true },
      fill: { kind: 'solid', color: { kind: 'theme', slot: 'accent1', lumMod: 0.75, alpha: 0.5 } },
      line: { width: 2, fill: { kind: 'solid', color: rgb('#FF0000') }, dash: 'dash', cap: 'round', join: 'bevel', head: { type: 'triangle', w: 'med', len: 'lg' }, tail: { type: 'oval' } },
    });
    const { shapes, xml } = await write([s]);
    expect(xml).toContain('<a:gd name="adj" fmla="val 12500"/>');
    expect(shapes).toEqual([{ ...s, geometry: { kind: 'preset', prst: 'roundRect', adjustments: { adj: 12500 }, known: true } }]);
  });

  it('writes the unknown preset under its own name and the editor\'s unknown: prefix without it', async () => {
    const { shapes, xml } = await write([shape({ geometry: { kind: 'preset', prst: 'mathPlus', known: false } }), shape({ pptxId: '3', geometry: { kind: 'preset', prst: 'unknown:cloudCallout', known: false } })]);
    expect(xml).toContain('prst="mathPlus"');
    expect(xml).toContain('prst="cloudCallout"');
    expect(shapes.map((s) => s.geometry)).toMatchObject([{ prst: 'mathPlus' }, { prst: 'cloudCallout' }]);
  });

  it('writes a preset name that is not an identifier as a rectangle and says so', async () => {
    const { xml, warnings } = await write([shape({ geometry: { kind: 'preset', prst: 'bad"/><x', adjustments: { 'a b': 1, adj: Infinity }, known: false } })]);
    expect(xml).toContain('prst="rect"');
    expect(xml).not.toContain('<x');
    expect(warnings).toContain('Changed on export: shapes with an unusable preset name written as rectangles (1).');
  });

  it('writes custom geometry without its outline until K13 and says so', async () => {
    const { shapes, warnings } = await write([shape({ geometry: { kind: 'custom' }, text: { props: {}, paragraphs: [para([text('kept')])] } })]);
    expect(shapes[0].geometry).toBeUndefined();
    expect(shapes[0].text?.paragraphs[0].runs[0].text).toBe('kept');
    expect(warnings).toEqual(['Changed on export: shapes with custom geometry written without their outline, until K13 (1).']);
  });

  it('writes text boxes, hidden shapes and placeholders', async () => {
    const lay = layout({ placeholders: [{ type: 'title', idx: undefined, frame: { x: 1, y: 2, w: 3, h: 4 }, levels: [{ fontSize: 40 }, {}, {}, {}, {}, {}, {}, {}, {}] }, { type: 'body', idx: 1, frame: { x: 5, y: 6, w: 7, h: 8 }, levels: [{ fontSize: 20 }, {}, {}, {}, {}, {}, {}, {}, {}] }] });
    const { shapes, xml } = await write(
      [
        shape({ textBox: true, hidden: true, name: 'tb' }),
        shape({ pptxId: '3', name: 'T', placeholder: { type: 'title' }, text: { props: {}, paragraphs: [para([text('Hello')], { fontSize: 40 })] } }),
        shape({ pptxId: '4', name: 'B', placeholder: { type: 'obj', idx: 1 }, text: { props: {}, paragraphs: [para([text('Body')], { fontSize: 20 })] } }),
      ],
      {},
      [lay]
    );
    expect(xml).toContain('txBox="1"');
    expect(xml).toContain('hidden="1"');
    expect(xml).toContain('<p:ph type="title"/>');
    expect(xml).toContain('<p:ph idx="1"/>');
    expect(shapes[0]).toMatchObject({ textBox: true, hidden: true });
    expect(shapes[1]).toMatchObject({ placeholder: { type: 'title' }, text: { paragraphs: [{ style: { fontSize: 40 } }] } });
    expect(shapes[2].placeholder).toEqual({ type: 'obj', idx: 1 });
  });

  it('writes a shape without a frame, a name or text', async () => {
    const { shapes } = await write([{ kind: 'shape', pptxId: '2', name: '' }]);
    expect(shapes).toEqual([{ kind: 'shape', pptxId: '2', name: '' }]);
  });
});

describe('fills and outlines', () => {
  const fills: [string, PptxFill][] = [
    ['none', { kind: 'none' }],
    ['solid', { kind: 'solid', color: rgb('#00FF00', 0.25) }],
    ['linear gradient', { kind: 'gradient', shape: 'linear', angle: 45, scaled: true, rotWithShape: true, stops: [{ pos: 0, color: rgb('#000000') }, { pos: 0.5, color: { kind: 'theme', slot: 'accent2', lumOff: 0.2 } }, { pos: 1, color: rgb('#FFFFFF', 0.5) }] }],
    ['linear gradient without angle', { kind: 'gradient', shape: 'linear', angle: 0, stops: [{ pos: 0, color: rgb('#000000') }] }],
    ['radial gradient', { kind: 'gradient', shape: 'radial', stops: [{ pos: 0, color: rgb('#111111') }, { pos: 1, color: rgb('#222222') }] }],
    ['rect gradient', { kind: 'gradient', shape: 'rect', stops: [{ pos: 0, color: rgb('#111111') }] }],
    ['shape gradient', { kind: 'gradient', shape: 'shape', stops: [{ pos: 0, color: rgb('#111111') }] }],
    ['pattern', { kind: 'pattern', prst: 'ltUpDiag', fg: rgb('#123456'), bg: rgb('#ABCDEF') }],
    ['pattern without colours', { kind: 'pattern', prst: 'pct50' }],
  ];
  it.each(fills)('%s: shape fill, outline fill and slide background read back as written', async (_n, fill) => {
    const { shapes, back } = await write([shape({ fill, line: { fill } })], { background: fill });
    expect(shapes[0].fill).toEqual(fill);
    expect(shapes[0].line?.fill).toEqual(fill);
    expect(back.slides[0].background).toEqual(fill);
  });

  it('writes a linear gradient that has no angle as angle 0, and says every kind of change once, in order', async () => {
    const stops = [{ pos: 0, color: rgb('#000000') }];
    const pic: PptxFill = { kind: 'picture' };
    const { shapes, warnings } = await write([shape({ fill: { kind: 'gradient', shape: 'linear', stops }, line: { fill: pic, dash: 'custom' } })]);
    expect(shapes[0].fill).toEqual({ kind: 'gradient', shape: 'linear', angle: 0, stops });
    expect(warnings).toEqual(['Changed on export: custom dash patterns written as solid lines (1).', 'Changed on export: picture fills left out, until K13 (1).']);
  });

  it('leaves picture fills out and says so once per use', async () => {
    const pic: PptxFill = { kind: 'picture', part: 'ppt/media/a.png' };
    const { shapes, warnings, xml } = await write([shape({ fill: pic, line: { fill: pic, width: 1 } })], { background: pic });
    expect(shapes[0].fill).toBeUndefined();
    expect(xml).not.toContain('<p:bg>');
    expect(warnings).toEqual(['Changed on export: picture fills left out, until K13 (3).']);
  });

  it('writes an empty outline and a custom dash as a solid line, with a warning', async () => {
    const { shapes, warnings } = await write([shape({ line: {} }), shape({ pptxId: '3', line: { dash: 'custom', width: 1 } })]);
    expect(shapes[0].line).toEqual({});
    expect(shapes[1].line).toEqual({ width: 1 });
    expect(warnings).toEqual(['Changed on export: custom dash patterns written as solid lines (1).']);
  });
});

describe('text', () => {
  const body = (paragraphs: PptxParagraph[], props = {}) => shape({ text: { props, paragraphs } });

  it('writes body properties, run marks, fields and breaks that read back equal', async () => {
    const runs: PptxRun[] = [
      text('plain'),
      text('rich', { bold: true, italic: false, underline: 'dbl', strike: 'double', size: 18.5, color: { kind: 'theme', slot: 'accent3' }, highlight: rgb('#FFFF00'), font: 'major', fontEastAsian: 'minor', fontComplex: 'Arial', baseline: 0.3, caps: 'small', spacing: 1.5, lang: 'de-DE' }),
      text('under', { underline: 'sng', strike: 'single', caps: 'all', baseline: -0.25, font: 'Georgia' }),
      { kind: 'break', text: '' },
      { kind: 'field', text: '7', fieldType: 'slidenum', lang: 'en-US' },
      { kind: 'field', text: 'x' },
    ];
    const props = { insets: { l: 1, t: 2, r: 3, b: 4 }, anchor: 'bottom' as const, anchorCtr: true, vert: 'vert270', wrap: 'none' as const, columns: 2, columnSpacing: 12, rot: 15, autofit: { kind: 'shrink' as const, fontScale: 0.625, lnSpcReduction: 0.2 } };
    const s = body([para(runs, { align: 'center' }, { endSize: 14 })], props);
    const { shapes, xml } = await write([s]);
    expect(xml.match(/<a:fld id="\{F1E1D000-0000-4000-8000-00000000000[12]\}"/g)).toHaveLength(2);
    expect(shapes[0].text).toEqual(s.text);
  });

  it.each([
    ['none', { kind: 'none' as const }],
    ['shrink', { kind: 'shrink' as const }],
    ['resize', { kind: 'resize' as const }],
  ])('writes the %s autofit mode', async (_n, autofit) => {
    const { shapes } = await write([body([para([text('a')])], { autofit })]);
    expect(shapes[0].text?.props.autofit).toEqual(autofit);
    expect((await write([body([para([text('a')])], { anchor: 'top' })])).shapes[0].text?.props.anchor).toBe('top');
    expect((await write([body([para([text('a')])], { anchor: 'middle', anchorCtr: false })])).shapes[0].text?.props).toEqual({ anchor: 'middle', anchorCtr: false });
  });

  it('writes the paragraph level style: spacing, bullets, tabs, indents, direction and levels', async () => {
    const styles: PptxLevelStyle[] = [
      { align: 'right', marginLeft: 20, indent: -10, lineHeight: 1.2, spaceBefore: 6, spaceAfter: 3, bullet: { kind: 'char', char: '•' }, bulletColor: rgb('#FF0000'), bulletSizePct: 0.8, bulletFont: 'Arial', tabStops: [{ pos: 40, align: 'l' }, { pos: 80, align: 'dec' }], fontSize: 20, bold: true, italic: true, color: rgb('#333333'), font: 'minor' },
      { align: 'justify', lineSpacingPts: 18, spaceBeforePct: 0.2, spaceAfterPct: 0.1, bullet: { kind: 'auto', scheme: 'arabicPeriod', startAt: 3 } },
      { align: 'left', bullet: { kind: 'none' } },
      {},
    ];
    const paragraphs = styles.map((style, i) => para([text(`p${i}`)], style, { level: i === 3 ? 8 : i, ...(i === 1 ? { rtl: true } : {}) }));
    const { shapes } = await write([body(paragraphs)]);
    expect(shapes[0].text?.paragraphs).toEqual(paragraphs);
  });

  it('puts what the paragraphs of a level share into the list style and only the difference on a paragraph', async () => {
    const a = para([text('a')], { fontSize: 20, bold: true, align: 'left' });
    const b = para([text('b')], { fontSize: 20, bold: true, align: 'right' });
    const c = para([text('c')], { fontSize: 20, spaceBefore: 4, spaceBeforePct: undefined }, { level: 1 });
    expect(sharedLevels([a, b, c])[0]).toEqual({ fontSize: 20, bold: true });
    expect(sharedLevels([a, b, c])[1]).toEqual({ fontSize: 20, spaceBefore: 4 });
    expect(sharedLevels([a, b, c])[2]).toEqual({});
    const { xml, shapes } = await write([body([a, b, c])]);
    expect(xml).toContain('<a:lvl1pPr><a:defRPr sz="2000" b="1"></a:defRPr></a:lvl1pPr>');
    expect(xml).toContain('<a:pPr algn="r"/>');
    expect(shapes[0].text?.paragraphs).toEqual([a, b, c]);
  });

  it('writes a paragraph level outside 0 to 8 at the nearest level', async () => {
    const { shapes } = await write([body([para([text('a')], { fontSize: 9 }, { level: 12 }), para([text('b')], {}, { level: -3 })])]);
    expect(shapes[0].text?.paragraphs.map((p) => p.level)).toEqual([8, 0]);
    expect(shapes[0].text?.paragraphs[0].style).toEqual({ fontSize: 9 });
  });

  it('writes a body without paragraphs and an empty paragraph', async () => {
    const empty = await write([body([])]);
    expect(empty.xml).toContain('<a:p/>');
    const { shapes } = await write([body([para([], {}, { endSize: 11 })])]);
    expect(shapes[0].text?.paragraphs).toEqual([{ level: 0, style: {}, runs: [], endSize: 11 }]);
  });

  it('escapes text and drops characters XML cannot carry', async () => {
    const { xml, shapes } = await write([body([para([text('a<b>&"\u0001c\u0000d')])])]);
    expect(xml).toContain('a&lt;b&gt;&amp;&quot;cd');
    expect(shapes[0].text?.paragraphs[0].runs[0].text).toBe('a<b>&"cd');
  });

  it('writes an underline outside the format as single and says so', async () => {
    const { shapes, warnings } = await write([body([para([text('a', { underline: 'zigzag' })])])]);
    expect(shapes[0].text?.paragraphs[0].runs[0].underline).toBe('sng');
    expect(warnings).toEqual(['Changed on export: underline styles outside the format written as single (1).']);
  });
});

describe('hyperlinks', () => {
  it('writes shape and run links as relationships shared by target', async () => {
    const link = { kind: 'url' as const, href: 'https://example.com/a?b=1&c=2', tooltip: 'go' };
    const s = shape({ link, text: { props: {}, paragraphs: [para([text('x', { link }), text('y', { link: { kind: 'url', href: 'mailto:a@b.c' } })])] } });
    const { shapes, rels } = await write([s]);
    expect(rels.match(/relationships\/hyperlink/g)).toHaveLength(2);
    expect(rels).toContain('Target="https://example.com/a?b=1&amp;c=2" TargetMode="External"');
    expect(shapes[0].link).toEqual(link);
    expect(shapes[0].text?.paragraphs[0].runs.map((r) => r.link)).toEqual([link, { kind: 'url', href: 'mailto:a@b.c' }]);
  });

  it('leaves out an address the link rules do not allow', async () => {
    const bad = (href: string) => shape({ pptxId: String(href.length + 2), link: { kind: 'url', href } });
    const { shapes, warnings, rels } = await write([bad('javascript:alert(1)'), bad('/relative'), bad('java\tscript:x'), { ...bad('data:text/html,x'), link: { kind: 'action', target: 'macro:run' } }]);
    expect(shapes.map((s) => s.link)).toEqual([undefined, undefined, undefined, undefined]);
    expect(rels).not.toContain('hyperlink');
    expect(warnings).toEqual(['Changed on export: hyperlinks with an address the editor does not allow left out (4).']);
  });

  it('writes a jump to another slide as a relationship to the slide as it is numbered now, and other actions as written', async () => {
    const jump = { kind: 'action' as const, target: 'ppaction://hlinksldjump', slidePart: 'ppt/slides/slide2.xml' };
    const first = slide({ elements: [shape({ link: jump }), shape({ pptxId: '3', link: { kind: 'action', target: 'ppaction://hlinksldjump', slidePart: 'ppt/slides/gone.xml' } }), shape({ pptxId: '4', link: { kind: 'action', target: 'ppaction://hlinkshowjump?jump=nextslide' } }), shape({ pptxId: '5', link: { kind: 'action', target: 'file:///tmp/x.txt' } })] });
    const second = slide({ pptxId: '257', partName: 'ppt/slides/slide2.xml' });
    const { bytes, warnings } = await exportPptx(deck([first, second]));
    expect(warnings).toEqual([]);
    expect(await lintPptxPackage(bytes)).toEqual([]);
    const back = await importPptx(bytes);
    expect((back.slides[0].elements as PptxShape[]).map((s) => s.link)).toEqual([
      { kind: 'action', target: 'ppaction://hlinksldjump', slidePart: 'ppt/slides/slide2.xml' },
      { kind: 'action', target: 'ppaction://hlinksldjump' },
      { kind: 'action', target: 'ppaction://hlinkshowjump?jump=nextslide' },
      { kind: 'action', target: 'file:///tmp/x.txt' },
    ]);
  });
});

describe('element ids', () => {
  const el = (pptxId: string): PptxElement => shape({ pptxId });
  it('keeps usable ids and gives the others the next free number', () => {
    const list = [el('5'), el('5'), el(''), el('1'), el('abc'), el('7'), el('99999999999'), el('3')];
    const ids = allocateElementIds(list);
    expect(list.map((e) => ids.get(e))).toEqual([5, 2, 4, 6, 8, 7, 9, 3]);
  });
  it('numbers the children of groups with the rest and keeps what it reserves', () => {
    const [a, b] = [el('2'), el('4')];
    const group: PptxElement = { kind: 'group', pptxId: '2', name: 'g', children: [a, b] };
    const c = el('3');
    const ids = allocateElementIds([group, c], 3);
    expect([group, a, b, c].map((e) => ids.get(e))).toEqual([5, 6, 4, 7]);
  });
  it('writes the ids of a slide whose shapes repeat them, unique', async () => {
    const { shapes } = await write([el('2'), el('2'), el('2')]);
    expect(shapes.map((s) => s.pptxId)).toEqual(['2', '3', '4']);
  });
});

describe('backgrounds and layout shapes', () => {
  it('writes a slide background before the shape tree', async () => {
    const { xml, back } = await write([], { background: { kind: 'solid', color: rgb('#112233') } });
    expect(xml.indexOf('<p:bg>')).toBeLessThan(xml.indexOf('<p:spTree>'));
    expect(back.slides[0].background).toEqual({ kind: 'solid', color: rgb('#112233') });
  });

  it('writes a layout\'s static shapes after its placeholders, with ids of their own and hyperlinks', async () => {
    const link = { kind: 'url' as const, href: 'https://example.com' };
    const lay = layout({
      placeholders: [{ type: 'title', frame: { x: 0, y: 0, w: 100, h: 50 }, levels: [{}, {}, {}, {}, {}, {}, {}, {}, {}] }],
      elements: [shape({ pptxId: '2', name: 'Logo', fill: { kind: 'solid', color: rgb('#FF0000') }, link }), shape({ pptxId: '9', name: 'Rule' })],
      staticShapes: 2,
    });
    const { bytes } = await exportPptx(deck([slide()], [lay]));
    expect(await lintPptxPackage(bytes)).toEqual([]);
    const back = await importPptx(bytes);
    expect(back.layouts[0].elements).toMatchObject([{ name: 'Logo', fill: { kind: 'solid' }, link }, { name: 'Rule' }]);
    const zip = await JSZip.loadAsync(bytes);
    expect(await zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels')?.async('string')).toContain('relationships/hyperlink');
  });
});

describe('round trip over the corpus', () => {
  // decks made of shapes and text only: everything the reader models comes back, frames within half a pixel
  it.each(['text-basic.pptx', 'text-rich.pptx', 'rtl-cjk.pptx', 'shapes-presets.pptx', 'themes.pptx', 'sizes-4x3.pptx', 'fills-lines.pptx'])('%s: shapes and text read back equal', async (name) => {
    const first = await importPptx(fixture(name));
    const second = await importPptx((await exportPptx(first)).bytes);
    const strip = (e: PptxElement) => (e.kind === 'shape' && !(e.geometry?.kind === 'custom') && e.fill?.kind !== 'picture' && e.line?.fill?.kind !== 'picture' ? [e] : []);
    first.slides.forEach((s, i) => {
      const want = (s.elements ?? []).flatMap(strip);
      const got = (second.slides[i].elements ?? []).filter((e) => want.some((w) => w.name === e.name));
      expect(got.map((e) => e.kind === 'shape' && e.text)).toEqual(want.map((e) => e.kind === 'shape' && e.text));
      want.forEach((w, k) => {
        const g = got[k] as PptxShape;
        for (const key of ['x', 'y', 'w', 'h'] as const) expect(Math.abs((g.frame?.[key] ?? 0) - (w.frame?.[key] ?? 0))).toBeLessThanOrEqual(0.5);
      });
    });
    expect(second.layouts.map((l) => l.elements.length)).toEqual(first.layouts.map((l) => l.elements.length));
  });
});
