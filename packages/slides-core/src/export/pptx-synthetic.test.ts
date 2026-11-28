import type { PptxElement, PptxLayout, PptxShape, PresentationImportResult } from '@b32nio/spindle-shared/export/pptx';
import { hashBytes } from '@b32nio/spindle-shared/opaque';
import { DeckImpl } from '../deck';
import type { SlideElement } from '../scene/types';
import { exportToPptx, importFromPptx } from './pptx';

// The mapping with hand-built interchange values: what the reader hands over is replaced, the writer is the real one but its input is captured.
const mocks = {
  read: jest.fn(),
  write: jest.fn(),
};
jest.mock('@b32nio/spindle-io/pptx', () => {
  const actual = jest.requireActual('@b32nio/spindle-io/pptx');
  return {
    ...actual,
    importPptx: (...a: unknown[]) => mocks.read(...a),
    exportPptx: (...a: unknown[]) => mocks.write(...a),
    detectPptxLoss: actual.detectPptxLoss,
  };
});

const BYTES = new Uint8Array([1, 2, 3]);
const red = { kind: 'rgb' as const, hex: '#FF0000' };
const solid = { kind: 'solid' as const, color: red };
const line = { fill: solid, width: 2 };
const frame = (x = 10, y = 20, w = 100, h = 50) => ({ x, y, w, h });
const rectGeom = { kind: 'preset' as const, prst: 'rect', known: true };
const text = (s: string) => ({ props: {}, paragraphs: [{ level: 0, style: {}, runs: [{ kind: 'text' as const, text: s }] }] });
const shape = (id: string, extra: Partial<PptxShape> = {}): PptxShape => ({ kind: 'shape', pptxId: id, name: `S${id}`, frame: frame(), ...extra });

const base = (slides: PresentationImportResult['slides'], layouts: PptxLayout[] = []): PresentationImportResult => ({
  slideSize: { w: 1280, h: 720 },
  theme: { name: 'T', colors: {} as never, fonts: { major: 'A', minor: 'B', majorEastAsian: 'C' } },
  colorMap: {} as never,
  slides,
  layouts,
  fonts: { used: [], unavailable: [] },
  opaqueParts: [],
  warnings: ['reader warning'],
});
const slide = (elements: PptxElement[], extra: Record<string, unknown> = {}) => ({ pptxId: '256', partName: 'ppt/slides/slide1.xml', elements, ...extra });

async function importOf(result: PresentationImportResult, options = {}) {
  mocks.read.mockResolvedValueOnce(result);
  const deck = new DeckImpl();
  const out = await importFromPptx(BYTES, deck, options);
  return { deck, data: deck.getData(), ...out };
}
const els = (data: ReturnType<DeckImpl['getData']>, i = 0): SlideElement[] => data.slides[i].elements;

beforeEach(() => {
  mocks.read.mockReset();
  mocks.write.mockReset();
});

describe('importFromPptx', () => {
  it('refuses a file with no slides', async () => {
    mocks.read.mockResolvedValueOnce(base([]));
    await expect(importFromPptx(BYTES, new DeckImpl())).rejects.toThrow('no slides');
  });

  it('passes limits to the reader but not the comment time', async () => {
    await importOf(base([slide([])]), { limits: { maxEntries: 5 }, opaqueLimits: { maxPartBytes: 5 }, now: new Date(0) });
    expect(mocks.read).toHaveBeenCalledWith(BYTES, { limits: { maxEntries: 5 }, opaqueLimits: { maxPartBytes: 5 } });
  });

  it('maps slide properties, layouts, theme, sections, hidden and failed slides', async () => {
    const layout: PptxLayout = {
      partName: 'ppt/slideLayouts/slideLayout1.xml',
      name: 'Title',
      masterName: 'M',
      masterPart: 'm',
      background: red,
      placeholders: [
        { type: 'ctrTitle', idx: 0, frame: { ...frame(), rot: 5 }, anchor: 'middle', prompt: 'Click', levels: [{ fontSize: 30, bold: true, italic: false, color: red, align: 'center', font: 'major', bullet: { kind: 'char', char: '*' }, marginLeft: 5, indent: -5, lineHeight: 1.1, spaceBefore: 3 }, {}, {}] },
        { type: 'body', frame: frame(), bodyProps: { anchor: 'bottom' }, levels: [{}, {}] },
        { type: 'body', idx: 2, frame: frame(), levels: [{ fontSize: 9 }] },
        { type: 'dt', frame: frame(), levels: [] },
        { type: 'sldNum', levels: [] },
      ],
      elements: [shape('3', { textBox: true, text: text('Static') })],
      staticShapes: 1,
    };
    const r = base(
      [
        slide([shape('2', { textBox: true, text: text('Hello') })], {
          hidden: true,
          layoutPart: layout.partName,
          background: { kind: 'solid', color: red },
          notes: text('speaker'),
          opaque: { transition: '<p:transition/>' },
        }),
        slide([], { pptxId: '257', layoutPart: 'unknown.xml', background: { kind: 'picture' }, notes: text(''), opaque: {} }),
        { pptxId: '258', partName: 'x', error: 'bad xml' },
      ],
      [layout, { partName: 'l2', name: 'Plain', masterPart: 'm', placeholders: [], elements: [], staticShapes: 0 }]
    );
    r.sections = [{ name: 'S', slideIds: ['256'] }];
    const { data, warnings } = await importOf(r);
    expect(data.slides).toHaveLength(3);
    expect(data.slides[0]).toMatchObject({ hidden: true, layoutRef: layout.partName, background: solid, opaque: { transition: '<p:transition/>' } });
    expect(data.slides[0].notes).toBeDefined();
    expect(data.slides[1]).not.toHaveProperty('layoutRef');
    expect(data.slides[1]).not.toHaveProperty('background');
    expect(data.slides[1]).not.toHaveProperty('notes');
    expect(data.slides[1]).not.toHaveProperty('opaque');
    expect(data.slides[2].elements[0]).toMatchObject({ type: 'text' });
    expect(warnings).toEqual(expect.arrayContaining(['reader warning', expect.stringContaining('Slide 3 could not be read (bad xml)'), expect.stringContaining('picture fill'), expect.stringContaining('1 slide section')]));
    expect(data.theme.fonts).toEqual({ major: 'A', minor: 'B' });
    const l = data.layouts!.find((x) => x.id === layout.partName)!;
    expect(l.placeholders.map((p) => p.type)).toEqual(['centerTitle', 'body', 'body']);
    expect(l.placeholders[0]).toMatchObject({ frame: { rotation: 5 }, prompt: 'Click', style: { vAlign: 'middle' } });
    expect(l.placeholders[1].style).toEqual({ vAlign: 'bottom' });
    expect(l.textStyles!.centerTitle![0]).toMatchObject({ fontSize: 40, bold: true, italic: false, align: 'center', font: 'major', marginLeft: 5, firstLineIndent: -5, lineHeight: 1.1, spaceBefore: 4 });
    expect(l.textStyles!.body).toEqual([{ fontSize: 12 }]);
    expect(l.background).toEqual(solid);
    expect(l.elements).toHaveLength(1);
    expect(data.layouts!.find((x) => x.id === 'l2')).not.toHaveProperty('elements');
  });

  it('says a single unreadable picture is a placeholder, and reads a slide with no element list', async () => {
    const { data, warnings } = await importOf(base([{ pptxId: '1', partName: 'p' }, slide([{ kind: 'picture', pptxId: '1', name: 'x' }])]));
    expect(data.slides[0].elements).toEqual([]);
    expect(warnings).toContain('1 picture could not be shown and is a labelled placeholder.');
  });

  it('falls back to the built-in layouts when the file has none', async () => {
    const { data } = await importOf(base([slide([])]));
    expect(data.layouts!.length).toBeGreaterThan(0);
  });

  it('maps shapes, text boxes, placeholders, geometry and adjustments', async () => {
    const { data, warnings } = await importOf(
      base([
        slide([
          shape('1', { placeholder: { type: 'title' }, text: text('T') }),
          shape('2', { placeholder: { type: 'dt' }, text: text('date') }),
          shape('3'),
          shape('4', { text: text('plain box') }),
          shape('5', { text: text('filled box'), fill: solid }),
          shape('6', { geometry: { kind: 'preset', prst: 'roundRect', adjustments: { adj: 25000 }, known: true }, fill: solid, line, frame: { ...frame(), flipH: true, flipV: true } }),
          shape('7', { geometry: { kind: 'preset', prst: 'roundRect', adjustments: { adj: 15000 }, known: true } }),
          shape('8', { geometry: { kind: 'preset', prst: 'roundRect', adjustments: { other: 1 }, known: true } }),
          shape('9', { geometry: { kind: 'preset', prst: 'rect', adjustments: { adj: 1 }, known: true } }),
          shape('10', { geometry: { kind: 'preset', prst: 'pie', adjustments: { adj1: 5400000, adj2: 16200000 }, known: true } }),
          shape('11', { geometry: { kind: 'preset', prst: 'mathPlus', adjustments: { adj1: 10000 }, known: true } }),
          shape('12', { geometry: { kind: 'preset', prst: 'weirdThing', adjustments: { a: 5, b: NaN, c: 7 }, known: false } }),
          shape('13', { geometry: { kind: 'preset', prst: 'weirdOther', adjustments: { a: NaN }, known: false } }),
          shape('14', { geometry: { kind: 'preset', prst: 'rtTriangle', known: true } }),
          shape('15', { geometry: { kind: 'custom', paths: [{ w: 10, h: 10, fill: 'lighten', stroke: false, extrusionOk: true, commands: [{ op: 'move', x: 0, y: 0 }, { op: 'close' }] }, { w: 1, h: 1, fill: 'bogus', commands: [] }] } }),
          shape('16', { geometry: { kind: 'custom' } }),
          shape('17', { geometry: { kind: 'custom', paths: [] } }),
          shape('18', { fill: { kind: 'picture' } }),
          shape('19', { hidden: true }),
          shape('', { text: text('no id'), textBox: true }),
          shape('20', { frame: undefined }),
          shape('21', { textBox: true, line }),
          shape('22', { textBox: true, text: text('lined'), line }),
        ]),
      ])
    );
    const e = els(data);
    expect(e[0]).toMatchObject({ type: 'text', placeholder: { type: 'title', idx: 0 }, pptxId: '1' });
    expect(e[1]).toMatchObject({ type: 'text' });
    expect(e[1]).not.toHaveProperty('placeholder');
    expect(e[2]).toMatchObject({ type: 'shape', shape: 'rect', fill: { kind: 'none' } });
    expect(e[3]).toMatchObject({ type: 'text' });
    expect(e[4]).toMatchObject({ type: 'shape', shape: 'rect', fill: solid });
    expect(e[5]).toMatchObject({ type: 'shape', shape: 'roundRect', adjustments: [0.25], flipH: true, flipV: true, stroke: { width: 2 } });
    expect(e[6]).not.toHaveProperty('adjustments');
    expect(e[7]).not.toHaveProperty('adjustments');
    expect(e[8]).not.toHaveProperty('adjustments');
    expect(e[9]).toMatchObject({ adjustments: [0.25, 0.75] });
    expect(e[10]).toMatchObject({ shape: 'mathPlus', adjustments: [0.1] });
    expect(e[11]).toMatchObject({ shape: 'unknown:weirdThing', adjustments: [5, 7] });
    expect(e[12]).not.toHaveProperty('adjustments');
    expect(e[13]).toMatchObject({ shape: 'rightTriangle' });
    expect((e[14] as { path: { paths: unknown[] } }).path.paths).toEqual([
      { w: 10, h: 10, fill: 'lighten', stroke: false, extrusionOk: true, commands: [{ op: 'move', x: 0, y: 0 }, { op: 'close' }] },
      { w: 1, h: 1, commands: [] },
    ]);
    expect(e[15]).toMatchObject({ shape: 'rect' });
    expect(e[16]).toMatchObject({ shape: 'rect' });
    expect(e[17]).toMatchObject({ type: 'shape', fill: { kind: 'none' } });
    expect(e[18]).not.toHaveProperty('pptxId');
    expect(e[20]).toMatchObject({ type: 'text', stroke: { width: 2 } });
    expect(e[21]).toMatchObject({ type: 'text', stroke: { width: 2 } });
    expect(e[19]).toMatchObject({ x: 0, y: 0, w: 0, h: 0 });
    expect(warnings).toEqual(expect.arrayContaining(['1 hidden shape not imported.', expect.stringContaining('2 custom outlines')]));
  });

  it('maps pictures, tables, opaque content and the placeholders pictures fill', async () => {
    const png = 'data:image/png;base64,AA==';
    const { data, warnings } = await importOf(
      base([
        slide([
          { kind: 'picture', pptxId: '1', name: 'P', frame: { ...frame(), flipH: true, flipV: true }, dataUrl: png, crop: { l: 1, t: 2, r: 3, b: 4 }, alpha: 0.5, placeholder: { type: 'pic', idx: 3 } },
          { kind: 'picture', pptxId: '2', name: 'Q', dataUrl: png, alpha: 1 },
          { kind: 'picture', pptxId: '3', name: 'Gone', unsupported: 'emf' },
          { kind: 'picture', pptxId: '', name: 'Svg', dataUrl: 'data:text/plain,hi' },
          { kind: 'picture', pptxId: '30', name: 'Plain' },
          { kind: 'opaque', opaqueKind: 'chart', pptxId: '4', name: 'C', altText: 'alt', title: 'Sales', poster: png, parts: [{ hash: 'h', role: 'chart' }], props: { a: 'b' }, frame: frame() },
          { kind: 'opaque', opaqueKind: 'ink', pptxId: '', name: 'I', parts: [] },
          { kind: 'table', pptxId: '5', name: 'T', columns: [], rows: [] },
          {
            kind: 'table', pptxId: '6', name: 'T2', frame: undefined, columns: [100, 100], flags: { firstRow: true },
            rows: [
              { height: 20, cells: [{ text: text('a'), gridSpan: 2, anchor: 'middle', margins: { l: 1, t: undefined }, fill: solid, borders: { l: { fill: solid, width: 1 }, t: null } }, { text: text('b'), hMerge: true }] },
              { height: 30, cells: [{ text: text('c'), rowSpan: 2, fill: { kind: 'picture' }, borders: { r: {} } }, { text: text('d'), vMerge: true, gridSpan: 1, rowSpan: 1 }] },
              { height: 0, cells: [] },
            ],
          },
          { kind: 'table', pptxId: '', name: 'Z', columns: [0, 0], rows: [{ height: 0, cells: [{ text: text('z') }] }] },
        ]),
      ])
    );
    const e = els(data);
    expect(e[0]).toMatchObject({ type: 'image', src: png, flipH: true, flipV: true, crop: { l: 1, t: 2, r: 3, b: 4 }, opacity: 0.5, placeholder: { type: 'image', idx: 3 } });
    expect(e[1]).not.toHaveProperty('opacity');
    expect(e[2]).toMatchObject({ type: 'shape' });
    expect(e[3]).toMatchObject({ type: 'shape' });
    expect(e[5]).toMatchObject({ type: 'opaque', kind: 'chart', alt: 'alt', title: 'Sales', poster: png, props: { a: 'b' } });
    expect(e[6]).toMatchObject({ type: 'opaque', kind: 'ink' });
    const table = e[7] as Extract<SlideElement, { type: 'table' }>;
    expect(table).toMatchObject({ rows: 3, cols: 2, w: 200, h: 50, styleFlags: { firstRow: true } });
    expect(table.cells[0][0]).toMatchObject({ gridSpan: 2, bodyStyle: { vAlign: 'middle', insets: { l: 1, t: 4.8 } } });
    expect(table.colFractions).toEqual([0.5, 0.5]);
    expect((e[8] as Extract<SlideElement, { type: 'table' }>).colFractions).toEqual([0.5, 0.5]);
    expect(warnings).toEqual(expect.arrayContaining([expect.stringContaining('1 empty table'), expect.stringContaining('3 pictures could not be shown')]));
  });

  it('builds groups, binds connectors to the shapes they name and keeps a free end where it was', async () => {
    const conn = (id: string, extra: Record<string, unknown> = {}): PptxElement => ({ kind: 'connector', pptxId: id, name: id, frame: frame(0, 0, 100, 50), line, ...extra } as PptxElement);
    const { data } = await importOf(
      base([
        slide([
          shape('1', { geometry: { kind: 'preset', prst: 'ellipse', known: true } }),
          shape('2', { geometry: rectGeom }),
          { kind: 'group', pptxId: '3', name: 'G', children: [shape('4'), { kind: 'group', pptxId: '5', name: 'H', children: [shape('6')] }, { kind: 'group', pptxId: '9', name: 'hid', hidden: true, children: [shape('10')] }, { kind: 'group', pptxId: '', name: 'anon', children: [shape('13')] }] },
          conn('7', { start: { id: '1', idx: 6 }, end: { id: '2', idx: 1 }, geometry: { kind: 'preset', prst: 'bentConnector3', adjustments: { adj1: 25000 }, known: true }, line: { ...line, head: { type: 'triangle' }, tail: { type: 'oval' } } }),
          conn('8', { start: { id: '1', idx: 0 }, geometry: { kind: 'preset', prst: 'curvedConnector3', known: true }, frame: { ...frame(0, 0, 100, 50), flipH: true, flipV: true }, line: { fill: { kind: 'none' }, head: { type: 'stealth' }, tail: { type: 'nope' } } }),
          conn('11', { end: { id: '2', idx: 3 }, line: {}, geometry: undefined }),
          conn('12', { start: { id: 'zzz', idx: 0 }, end: { id: '2', idx: 99 }, geometry: { kind: 'custom' }, line: { fill: solid, head: { type: 'diamond' }, tail: { type: 'arrow' } } }),
          conn('', {}),
        ]),
      ])
    );
    const e = els(data);
    const byP = (id: string) => e.find((x) => x.pptxId === id)!;
    expect(byP('4').groupId).toBeDefined();
    expect(byP('4').parentGroupId).toBeUndefined();
    expect(byP('6')).toMatchObject({ parentGroupId: byP('4').groupId });
    expect(byP('6').groupId).not.toBe(byP('4').groupId);
    expect(byP('4').groupId).toMatch(/^pptx-group-3-/);
    expect(byP('6').groupId).toMatch(/^pptx-group-5-/);
    expect(byP('13').groupId).toBeDefined();
    expect(byP('13').groupId).not.toMatch(/^pptx-group-/);
    expect(byP('10')).toBeUndefined();
    expect(byP('7')).toMatchObject({ type: 'line', connector: 'elbow', adjustments: [0.25], startBind: { elementId: byP('1').id, anchor: 'e' }, endBind: { elementId: byP('2').id, anchor: 'w' }, startArrow: 'triangle', endArrow: 'circle' });
    expect(byP('8')).toMatchObject({ connector: 'curved', startBind: { anchor: 'n' }, endPoint: { x: 0, y: 0 }, startArrow: 'triangle', stroke: { color: { alpha: 0 } } });
    expect(byP('8')).not.toHaveProperty('endArrow');
    expect(byP('11')).toMatchObject({ endBind: { anchor: 'e' }, startPoint: { x: 0, y: 0 }, stroke: { color: { slot: 'dk1' } } });
    expect(byP('12')).toMatchObject({ startArrow: 'circle', endArrow: 'arrow' });
    expect(byP('12')).not.toHaveProperty('connector');
    expect(byP('12')).not.toHaveProperty('startBind');
    expect(byP('12')).not.toHaveProperty('endBind');
  });

  it('turns comments into threads anchored to the element, or to nothing for a slide comment', async () => {
    const { data } = await importOf(
      base([
        slide([shape('1', { textBox: true, text: text('x') })], {
          threads: [
            { id: 't1', format: 'modern', anchor: { kind: 'element', pptxId: '1' }, resolved: true, comments: [{ author: 'Ann', initials: 'A', body: 'hi', createdAt: '2024-01-01T00:00:00Z' }, { author: 'Bo', body: 'yo' }] },
            { id: 't2', format: 'legacy', anchor: { kind: 'slide' }, comments: [{ author: 'Cy', body: 'slide' }] },
            { id: 't3', format: 'modern', anchor: { kind: 'element', pptxId: 'gone' }, comments: [] },
          ],
        }),
      ]),
      { now: new Date('2025-05-05T00:00:00Z') }
    );
    const [t1, t2, t3] = data.threads!;
    expect(t1).toMatchObject({ status: 'resolved', anchor: { elementId: els(data)[0].id }, createdAt: '2024-01-01T00:00:00Z' });
    expect(t1.comments[0]).toMatchObject({ authorName: 'Ann', authorInitials: 'A', authorId: 'pptx:Ann' });
    expect(t1.comments[1].createdAt).toBe('2025-05-05T00:00:00.000Z');
    expect(t2).toMatchObject({ status: 'open', anchor: { elementId: '' } });
    expect(t3.createdAt).toBe('2025-05-05T00:00:00.000Z');
  });

  it('keeps the opaque parts and text structure', async () => {
    const r = base([slide([shape('1', { textBox: true, text: { props: { vert: 'wordArtVert' }, paragraphs: [{ level: 0, style: {}, runs: [{ kind: 'text', text: 'a' }, { kind: 'break', text: '' }, { kind: 'text', text: 'b' }] }] } })])]);
    const bytes = new Uint8Array([1]);
    r.opaqueParts = [{ hash: await hashBytes(bytes), contentType: 'x', bytes }];
    const { data, warnings } = await importOf(r);
    expect(data.opaqueParts).toHaveLength(1);
    expect(warnings).toEqual(expect.arrayContaining(['1 line break became paragraph breaks.', expect.stringContaining('vertical text direction')]));
  });
});

describe('exportToPptx', () => {
  const captured = async (deck: DeckImpl, options = {}) => {
    mocks.write.mockResolvedValueOnce({ bytes: new Uint8Array([9]), warnings: ['w'] });
    const out = await exportToPptx(deck, options);
    return { out, p: mocks.write.mock.calls[0][0] as PresentationImportResult, opts: mocks.write.mock.calls[0][1] };
  };
  const el = (over: Record<string, unknown>) => ({ id: 'e', containerId: 's1', index: 'a0', x: 1, y: 2, w: 30, h: 40, rotation: 0, ...over }) as SlideElement;
  const deckOf = (elements: SlideElement[], extra: Record<string, unknown> = {}, slideExtra: Record<string, unknown> = {}) => {
    const deck = new DeckImpl('d', 'My deck');
    const data = deck.getData();
    deck.setData({
      ...data,
      slides: [{ id: 's1', index: 'a0', elements: elements.map((e, i) => ({ ...e, id: e.id === 'e' ? `e${i}` : e.id, index: `a${i}` })) as SlideElement[], ...slideExtra }],
      ...extra,
    });
    return deck;
  };

  it('writes the metadata, the slide properties and returns the writer\'s warnings', async () => {
    const deck = deckOf([], { threads: undefined }, { hidden: true, layoutRef: 'title', background: { kind: 'solid', color: red }, opaque: { timing: '<p:timing/>' }, notes: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'n' }] }] } });
    const { out, p, opts } = await captured(deck, { metadata: { creator: 'me' } });
    expect(out).toEqual({ bytes: new Uint8Array([9]), warnings: ['w'] });
    expect(opts).toEqual({ metadata: { title: 'My deck', creator: 'me' } });
    expect(p.slides[0]).toMatchObject({ pptxId: '256', partName: 'ppt/slides/slide1.xml', hidden: true, layoutPart: 'title', background: { kind: 'solid' }, opaque: { timing: '<p:timing/>' } });
    expect(p.slides[0].notes).toBeDefined();
    expect(p.layouts.length).toBeGreaterThan(0);
    expect(p.colorMap.bg1).toBe('lt1');
    mocks.write.mockClear();
    const plain = await captured(deckOf([], {}, { layoutRef: 'nope', notes: { type: 'doc', content: [{ type: 'paragraph' }] } }));
    expect(plain.p.slides[0]).not.toHaveProperty('layoutPart');
    expect(plain.p.slides[0]).not.toHaveProperty('notes');
    expect(plain.opts).toEqual({ metadata: { title: 'My deck' } });
  });

  it('writes shapes, text, images, tables and opaque content', async () => {
    const rich = { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'hi' }] }] };
    const deck = deckOf([
      el({ type: 'text', richText: rich, bodyStyle: {}, fill: { kind: 'solid', color: red }, stroke: { color: red, width: 1 }, placeholder: { type: 'body', idx: 1 }, rotation: 10 }),
      el({ type: 'text', richText: rich, placeholder: { type: 'title', idx: 0 } }),
      el({ type: 'text', richText: rich }),
      el({ type: 'shape', shape: 'roundRect', adjustments: [0.3], fill: { kind: 'none' }, flipH: true, richText: rich, placeholder: { type: 'image', idx: 0 } }),
      el({ type: 'shape', shape: 'pie', adjustments: [0.25, 0.5], fill: { kind: 'none' }, richText: { type: 'doc', content: [{ type: 'paragraph' }] } }),
      el({ type: 'shape', shape: 'roundRect', adjustments: [0.15], fill: { kind: 'none' } }),
      el({ type: 'shape', shape: 'rect', adjustments: [1], fill: { kind: 'none' } }),
      el({ type: 'shape', shape: 'unknown:weird', adjustments: [3, NaN], fill: { kind: 'none' } }),
      el({ type: 'shape', shape: 'unknown:one', adjustments: [3], fill: { kind: 'none' } }),
      el({ type: 'shape', shape: 'path', path: { paths: [{ w: 1, h: 1, fill: 'darken', stroke: false, extrusionOk: true, commands: [{ op: 'close' }] }, { w: 2, h: 2, fill: 'norm', commands: [] }] }, fill: { kind: 'none' } }),
      el({ type: 'shape', shape: 'path', adjustments: [1], fill: { kind: 'none' } }),
      el({ type: 'shape', shape: 'ellipse', fill: undefined as never, placeholder: { type: 'body', idx: 2 } }),
      el({ type: 'image', src: 'data:image/png;base64,AA==', naturalW: 1, naturalH: 1, flipV: true, crop: { l: 0, t: 0, r: 0, b: 5 }, opacity: 0.4, placeholder: { type: 'image', idx: 2 } }),
      el({ type: 'image', src: 'data:image/png;base64,AA==', naturalW: 1, naturalH: 1, crop: { l: 0, t: 0, r: 0, b: 0 }, opacity: 1 }),
      el({ type: 'image', src: 'data:image/png;base64,AA==', naturalW: 1, naturalH: 1 }),
      el({ type: 'image', src: 'data:image/png;base64,AA==', naturalW: 1, naturalH: 1, placeholder: { type: 'image', idx: 0 } }),
      el({ type: 'opaque', kind: 'chart', partRefs: [{ hash: 'h', role: 'chart' }], alt: 'a', title: 't', poster: 'p', props: { k: 'v' } }),
      el({ type: 'opaque', kind: 'ink', partRefs: [] }),
    ]);
    const { p } = await captured(deck);
    const e = p.slides[0].elements!;
    expect(e).toHaveLength(18);
    expect(e[0]).toMatchObject({ kind: 'shape', placeholder: { type: 'body', idx: 1 }, frame: { rot: 10 }, fill: { kind: 'solid' }, line: { width: 1 } });
    expect(e[1]).toMatchObject({ placeholder: { type: 'title' } });
    expect((e[1] as PptxShape).placeholder).not.toHaveProperty('idx');
    expect(e[2]).toMatchObject({ textBox: true, name: 'Text 3' });
    expect(e[3]).toMatchObject({ geometry: { prst: 'roundRect', adjustments: { adj: 30000 } }, frame: { flipH: true }, placeholder: { type: 'pic' } });
    expect(e[4]).toMatchObject({ geometry: { prst: 'pie', adjustments: { adj1: 5400000, adj2: 10800000 } } });
    expect(e[4]).not.toHaveProperty('text');
    expect((e[5] as PptxShape).geometry).not.toHaveProperty('adjustments');
    expect((e[6] as PptxShape).geometry).not.toHaveProperty('adjustments');
    expect(e[7]).toMatchObject({ geometry: { prst: 'weird', known: false, adjustments: { adj1: 3 } } });
    expect(e[8]).toMatchObject({ geometry: { adjustments: { adj: 3 } } });
    expect(e[9]).toMatchObject({ geometry: { kind: 'custom', paths: [{ fill: 'darken', stroke: false, extrusionOk: true }, { w: 2 }] } });
    expect((e[9] as PptxShape).geometry).toMatchObject({ paths: [expect.anything(), expect.not.objectContaining({ fill: expect.anything() })] });
    expect(e[10]).toMatchObject({ geometry: { prst: 'rect' } });
    expect(e[11]).toMatchObject({ placeholder: { type: 'body', idx: 2 } });
    expect(e[11]).not.toHaveProperty('fill');
    expect(e[12]).toMatchObject({ kind: 'picture', crop: { b: 5 }, alpha: 0.4, frame: { flipV: true }, placeholder: { type: 'pic', idx: 2 } });
    expect(e[13]).not.toHaveProperty('crop');
    expect(e[13]).not.toHaveProperty('alpha');
    expect(e[14]).not.toHaveProperty('crop');
    expect((e[15] as { placeholder: object }).placeholder).toEqual({ type: 'pic' });
    expect(e[16]).toMatchObject({ kind: 'opaque', opaqueKind: 'chart', altText: 'a', title: 't', poster: 'p', props: { k: 'v' } });
    expect(e[17]).toMatchObject({ opaqueKind: 'ink' });
    expect(new Set(e.map((x) => x.pptxId)).size).toBe(18);
  });

  it('writes tables with their borders resolved over the grid', async () => {
    const cell = (extra = {}) => ({ richText: { type: 'doc', content: [{ type: 'paragraph' }] }, ...extra });
    const deck = deckOf([
      el({
        type: 'table', rows: 2, cols: 2, colFractions: [0.25, 0.75], rowFractions: [0.5, 0.5], rowHeights: [10, 0], styleFlags: { firstRow: true, bandRow: false },
        cells: [
          [cell({ gridSpan: 2, bodyStyle: { insets: { l: 1, t: 2, r: 3, b: 4 }, vAlign: 'bottom' }, fill: { kind: 'solid', color: red }, borders: { l: null, t: { color: red, width: 3 } } }), cell({ hMerge: true })],
          [cell({ rowSpan: 1 }), cell({ vMerge: true })],
        ],
      }),
      el({ type: 'table', rows: 1, cols: 1, colFractions: [1], rowFractions: [1], cells: [[cell()]], border: { color: red, width: 2 } }),
    ]);
    const { p } = await captured(deck);
    const t = p.slides[0].elements![0] as Extract<PptxElement, { kind: 'table' }>;
    expect(t.columns).toEqual([7.5, 22.5]);
    expect(t.rows.map((r) => r.height)).toEqual([10, 20]);
    expect(t.flags).toEqual({ firstRow: true });
    expect(t.rows[0].cells[0]).toMatchObject({ gridSpan: 2, anchor: 'bottom', margins: { l: 1, b: 4 }, fill: { kind: 'solid' }, borders: { l: null, t: { width: 3 } } });
    expect(t.rows[0].cells[1]).toMatchObject({ hMerge: true });
    expect(t.rows[1].cells[1]).toMatchObject({ vMerge: true });
    const single = p.slides[0].elements![1] as Extract<PptxElement, { kind: 'table' }>;
    expect(single).not.toHaveProperty('flags');
    expect(single.rows[0].cells[0].borders!.l).toMatchObject({ width: 2 });
  });

  it('keeps a group\'s own pptxId through the editor so a timing target still resolves, and never reuses one', async () => {
    const box = (id: string, over: Record<string, unknown>) => el({ id, type: 'shape', shape: 'rect', fill: { kind: 'none' }, w: 10, h: 10, ...over });
    const deck = deckOf([
      box('a', { pptxId: '5', groupId: 'pptx-group-8-xyz', parentGroupId: 'pptx-group-9-xyz' }),
      box('b', { groupId: 'pptx-group-8-xyz', parentGroupId: 'pptx-group-9-xyz' }),
      box('c', { groupId: 'pptx-group-5-copy' }),
      box('d', { groupId: 'pptx-group-8-other' }),
      box('e2', { groupId: 'plain' }),
    ] as SlideElement[]);
    const { p } = await captured(deck);
    const groups = p.slides[0].elements!.filter((x) => x.kind === 'group');
    const ids = groups.map((g) => g.pptxId);
    expect(ids).toContain('9');
    const outer = groups.find((g) => g.pptxId === '9') as Extract<PptxElement, { kind: 'group' }>;
    expect((outer.children[0] as Extract<PptxElement, { kind: 'group' }>).pptxId).toBe('8');
    // '5' is an element's id, and '8' was taken by the first group: the others get free ids
    const all = JSON.stringify(p.slides[0].elements).match(/"pptxId":"(\d+)"/g)!;
    expect(new Set(all).size).toBe(all.length);
  });

  it('rebuilds groups with their box, resolves connectors against their shapes and writes comments', async () => {
    const rich = { type: 'doc', content: [{ type: 'paragraph' }] };
    const box = (id: string, over: Record<string, unknown>) => el({ id, type: 'shape', shape: 'ellipse', fill: { kind: 'none' }, ...over });
    const deck = deckOf(
      [
        box('a', { x: 0, y: 0, w: 10, h: 10, pptxId: '7' }),
        box('b', { x: 100, y: 100, w: 10, h: 10, groupId: 'g1', parentGroupId: 'g0' }),
        box('c', { x: 200, y: 0, w: 10, h: 10, groupId: 'g1', parentGroupId: 'g0' }),
        box('d', { x: 0, y: 300, w: 10, h: 10, groupId: 'g0', shape: 'rect' }),
        el({ id: 'tx', type: 'text', richText: rich, x: 300, y: 300, groupId: 'g2' }),
        el({ id: 'l1', type: 'line', stroke: { color: red, width: 1 }, connector: 'elbow', adjustments: [0.25], startArrow: 'arrow', endArrow: 'circle', startBind: { elementId: 'a', anchor: 'e' }, endBind: { elementId: 'tx', anchor: 'n' }, x: 0, y: 0, w: 1, h: 1 }),
        el({ id: 'l2', type: 'line', stroke: { color: red, width: 1 }, connector: 'curved', startArrow: 'none', endArrow: 'triangle', startBind: { elementId: 'gone', anchor: 'n' }, endBind: { elementId: 'a', anchor: 'nw' }, flipH: true }),
        el({ id: 'l3', type: 'line', stroke: { color: red, width: 1 }, connector: 'elbow', adjustments: [0.5], rotation: 5 }),
        el({ id: 'l5', type: 'line', stroke: { color: red, width: 1 }, startBind: { elementId: 'd', anchor: 'nw' } }),
        el({ id: 'l4', type: 'line', stroke: { color: red, width: 1 }, startBind: { elementId: 'a', anchor: 'e' } }),
      ].map((e) => ({ ...e, id: (e as { id: string }).id })) as SlideElement[],
      {
        threads: [
          { id: 't1', anchor: { slideId: 's1', elementId: 'a' }, status: 'resolved', createdAt: 'x', comments: [{ id: 'c1', authorId: 'u', authorName: 'Ann', authorInitials: 'A', body: 'hi', createdAt: '2024-01-01T00:00:00Z' }, { id: 'c2', authorId: 'u', authorName: 'Bo', body: 'yo', createdAt: 'y' }] },
          { id: 't2', anchor: { slideId: 's1', elementId: '' }, status: 'open', createdAt: 'x', comments: [] },
          { id: 't3', anchor: { slideId: 'other', elementId: '' }, status: 'open', createdAt: 'x', comments: [] },
        ],
      }
    );
    // element ids are kept as given above (deckOf renames only the default 'e')
    const { p } = await captured(deck);
    const top = p.slides[0].elements!;
    const kinds = top.map((x) => x.kind);
    expect(kinds).toEqual(['shape', 'group', 'group', 'connector', 'connector', 'connector', 'connector', 'connector']);
    const g0 = top[1] as Extract<PptxElement, { kind: 'group' }>;
    expect(g0.children.map((x) => x.kind)).toEqual(['group', 'shape']);
    expect(g0.frame).toEqual({ x: 0, y: 0, w: 210, h: 310 });
    const g1 = g0.children[0] as Extract<PptxElement, { kind: 'group' }>;
    expect(g1.frame).toEqual({ x: 100, y: 0, w: 110, h: 110 });
    const l1 = top[3] as Extract<PptxElement, { kind: 'connector' }>;
    expect(l1).toMatchObject({ start: { id: '7', idx: 6 }, geometry: { prst: 'bentConnector3', adjustments: { adj1: 25000 } }, line: { head: { type: 'arrow' }, tail: { type: 'oval' } } });
    expect(l1.end).toMatchObject({ idx: 0 });
    const l2 = top[4] as Extract<PptxElement, { kind: 'connector' }>;
    expect(l2.geometry).toMatchObject({ prst: 'curvedConnector3' });
    expect(l2.start).toBeUndefined();
    expect(l2.end).toEqual({ id: '7', idx: 1 });
    expect(l2.line!.head).toBeUndefined();
    expect(l2.line!.tail).toMatchObject({ type: 'triangle' });
    expect((top[6] as Extract<PptxElement, { kind: 'connector' }>).start).toBeUndefined();
    const l3 = top[5] as Extract<PptxElement, { kind: 'connector' }>;
    expect(l3.geometry).not.toHaveProperty('adjustments');
    expect(l3.frame).toMatchObject({ rot: 5 });
    expect(p.slides[0].threads).toEqual([
      { id: 't1', format: 'legacy', anchor: { kind: 'element', pptxId: '7' }, position: { x: 0, y: 0 }, resolved: true, comments: [{ author: 'Ann', initials: 'A', body: 'hi', createdAt: '2024-01-01T00:00:00Z' }, { author: 'Bo', body: 'yo', createdAt: 'y' }] },
      { id: 't2', format: 'legacy', anchor: { kind: 'slide' }, comments: [] },
    ]);
  });
});

describe('exported layouts', () => {
  it('writes placeholder text levels over the placeholder\'s own style, and static shapes only', async () => {
    mocks.write.mockResolvedValueOnce({ bytes: new Uint8Array(), warnings: [] });
    const deck = new DeckImpl();
    const data = deck.getData();
    deck.setData({
      ...data,
      layouts: [
        {
          id: 'L', name: 'N', masterName: 'M', background: { kind: 'solid', color: red },
          placeholders: [{ type: 'subtitle', idx: 1, frame: { x: 1, y: 2, w: 3, h: 4, rotation: 7 }, prompt: 'p', style: { fontSize: 30, vAlign: 'bottom' } }, { type: 'body', idx: 2, frame: { x: 1, y: 2, w: 3, h: 4, rotation: 0 } }],
          textStyles: { subtitle: [{ bold: true, bullet: { kind: 'char', char: '-', color: red, sizePct: 50 }, marginLeft: 4, firstLineIndent: -4, lineHeight: 1.2, spaceBefore: 6, italic: true, align: 'left', font: 'minor', color: red }, { bullet: { kind: 'auto', scheme: 'romanLower', startAt: 2 } }, { bullet: { kind: 'none' } }, { bullet: { kind: 'char' } }, { bullet: { kind: 'auto' } }] },
          elements: [
            { id: 'x', containerId: 'L', index: '', type: 'text', x: 0, y: 0, w: 5, h: 5, rotation: 0, richText: { type: 'doc', content: [{ type: 'paragraph' }] } },
            { id: 'y', containerId: 'L', index: 'a1', type: 'image', x: 0, y: 0, w: 5, h: 5, rotation: 0, src: 'data:image/png;base64,AA==', naturalW: 1, naturalH: 1 },
          ] as SlideElement[],
        },
        { id: 'plain', name: 'P', background: { kind: 'none' }, placeholders: [] },
      ],
      slides: [],
    });
    await exportToPptx(deck);
    const [l, plain] = (mocks.write.mock.calls[0][0] as PresentationImportResult).layouts;
    expect(l).toMatchObject({ partName: 'L', masterName: 'M', background: red, staticShapes: 1 });
    expect(l.placeholders[0]).toMatchObject({ type: 'subTitle', idx: 1, anchor: 'bottom', prompt: 'p', frame: { rot: 7 } });
    expect(l.placeholders[0].levels[0]).toMatchObject({ fontSize: 22.5, bold: true, italic: true, bullet: { kind: 'char', char: '-' }, bulletSizePct: 0.5, marginLeft: 4, indent: -4, spaceBefore: 4.5 });
    expect(l.placeholders[0].levels[1].bullet).toEqual({ kind: 'auto', scheme: 'romanLcPeriod', startAt: 2 });
    expect(l.placeholders[0].levels[2].bullet).toEqual({ kind: 'none' });
    expect(l.placeholders[0].levels[3].bullet).toEqual({ kind: 'char', char: '•' });
    expect(l.placeholders[0].levels[4].bullet).toEqual({ kind: 'auto', scheme: 'arabicPeriod' });
    expect(l.placeholders[1]).not.toHaveProperty('anchor');
    expect(plain).not.toHaveProperty('background');
    expect(plain).not.toHaveProperty('masterName');
  });
});
