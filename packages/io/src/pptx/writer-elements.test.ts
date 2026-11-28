import JSZip from 'jszip';
import { composeGroupChain } from '@b32nio/spindle-shared';
import type { PptxElement, PptxFrame, PptxGroup, PptxLayout, PptxPicture, PptxShape, PptxSlide, PptxTable, PptxTableCell, PptxTextBody, PresentationImportResult } from '@b32nio/spindle-shared/export/pptx';
import { DEFAULT_COLOR_MAP } from './color';
import { importPptx } from './reader';
import { FALLBACK_THEME } from './theme';
import { exportPptx } from './writer';
import { toChildSpace } from './writer-elements';
import { lintPptxPackage } from './writer-lint';

const layout: PptxLayout = { partName: 'ppt/slideLayouts/slideLayout1.xml', name: 'Blank', layoutType: 'blank', masterPart: 'ppt/slideMasters/slideMaster1.xml', placeholders: [], elements: [], staticShapes: 0 };
const slide = (over: Partial<PptxSlide> = {}): PptxSlide => ({ pptxId: '256', partName: 'ppt/slides/slide1.xml', layoutPart: layout.partName, ...over });
const deck = (slides: PptxSlide[]): PresentationImportResult => ({ slideSize: { w: 1280, h: 720 }, theme: FALLBACK_THEME, colorMap: DEFAULT_COLOR_MAP, slides, layouts: [layout], fonts: { used: [], unavailable: [] }, opaqueParts: [], warnings: [] });
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
const JPEG = 'data:image/jpeg;base64,/9j/4AA=';
const rgb = (hex: string) => ({ kind: 'rgb' as const, hex });
const body = (t: string): PptxTextBody => ({ props: {}, paragraphs: [{ level: 0, style: {}, runs: [{ kind: 'text', text: t }] }] });
const shape = (over: Partial<PptxShape> = {}): PptxShape => ({ kind: 'shape', pptxId: '2', name: 'Shape', frame: { x: 10, y: 20, w: 300, h: 100 }, ...over });
const picture = (over: Partial<PptxPicture> = {}): PptxPicture => ({ kind: 'picture', pptxId: '3', name: 'Pic', frame: { x: 5, y: 6, w: 100, h: 50 }, part: 'ppt/media/photo.png', mime: 'image/png', dataUrl: PNG, ...over });

/** Write one slide holding `elements` (and `extra`); returns the zip, the parts' text, the warnings and the model read back. */
async function write(elements: PptxElement[], extra: Partial<PptxSlide> = {}, more: PptxSlide[] = []) {
  const { bytes, warnings } = await exportPptx(deck([slide({ elements, ...extra }), ...more]));
  expect(await lintPptxPackage(bytes)).toEqual([]);
  const zip = await JSZip.loadAsync(bytes);
  const part = async (n: string) => (await zip.file(n)?.async('string')) as string;
  const back = await importPptx(bytes);
  return { zip, part, xml: await part('ppt/slides/slide1.xml'), rels: await part('ppt/slides/_rels/slide1.xml.rels'), warnings, back, els: (back.slides[0].elements ?? []) as PptxElement[] };
}

describe('pictures', () => {
  it('writes the image as a media part and reads back equal: crop, transparency, outline, link, geometry and alternative text', async () => {
    const pic = picture({
      altText: 'a <photo>',
      crop: { l: 10, t: 0, r: 25.5, b: 5 },
      alpha: 0.5,
      geometry: { kind: 'preset', prst: 'ellipse', known: true },
      line: { width: 2, fill: { kind: 'solid', color: rgb('#FF0000') } },
      link: { kind: 'url', href: 'https://example.com/' },
      hidden: true,
    });
    const { zip, xml, rels, els, part } = await write([pic]);
    expect(await zip.file('ppt/media/photo.png')?.async('base64')).toBe(PNG.split(',')[1]);
    expect(await part('[Content_Types].xml')).toContain('<Default Extension="png" ContentType="image/png"/>');
    expect(rels).toContain('Target="../media/photo.png"');
    expect(xml).toContain('<a:srcRect l="10000" r="25500" b="5000"/>');
    expect(xml).toContain('<a:alphaModFix amt="50000"/>');
    expect(els).toEqual([pic]);
  });

  it('shares one media part between pictures with the same image, and names parts it cannot keep', async () => {
    const { zip, els } = await write([
      picture(),
      picture({ pptxId: '4', name: 'Again' }),
      picture({ pptxId: '5', name: 'Other part', part: 'ppt/media/photo.png', dataUrl: JPEG, mime: 'image/jpeg' }),
      picture({ pptxId: '6', name: 'Wrong type', part: 'ppt/media/x.gif', dataUrl: JPEG, mime: 'image/jpeg' }),
      picture({ pptxId: '7', name: 'No part', part: undefined, dataUrl: 'data:image/jpeg;base64,/9j/4AAA' }),
      picture({ pptxId: '8', name: 'Elsewhere', part: 'ppt/other/p.png', dataUrl: PNG.replace('ggg=', 'ggA=') }),
    ]);
    const media = Object.keys(zip.files).filter((n) => /^ppt\/media\/.+\./.test(n));
    expect(media.sort()).toEqual(['ppt/media/image1.jpeg', 'ppt/media/image2.jpeg', 'ppt/media/image1.png', 'ppt/media/photo.png'].sort());
    expect(els.map((e) => (e as PptxPicture).part)).toEqual(['ppt/media/photo.png', 'ppt/media/photo.png', 'ppt/media/image1.jpeg', 'ppt/media/image1.jpeg', 'ppt/media/image2.jpeg', 'ppt/media/image1.png']);
  });

  it('does not let two part names differ only by case', async () => {
    const { zip } = await write([picture({ part: 'ppt/media/Photo.png' }), picture({ pptxId: '4', part: 'ppt/media/photo.png', dataUrl: PNG.replace('ggg=', 'ggA=') })]);
    expect(Object.keys(zip.files).filter((n) => /^ppt\/media\/.+\./.test(n)).sort()).toEqual(['ppt/media/Photo.png', 'ppt/media/image1.png']);
  });

  it('keeps a jpg part name and writes its content type', async () => {
    const { zip, part } = await write([picture({ part: 'ppt/media/photo.JPG', dataUrl: JPEG, mime: 'image/jpeg' })]);
    expect(zip.file('ppt/media/photo.JPG')).not.toBeNull();
    expect(await part('[Content_Types].xml')).toContain('Extension="jpg" ContentType="image/jpeg"');
  });

  it('writes a picture placeholder', async () => {
    const { els, xml } = await write([picture({ placeholder: { type: 'pic', idx: 13 } }), picture({ pptxId: '4', placeholder: { type: 'obj' } })]);
    expect(xml).toContain('<p:ph type="pic" idx="13"/>');
    expect((els[0] as PptxPicture).placeholder).toEqual({ type: 'pic', idx: 13 });
  });

  it('leaves out a picture without an image PowerPoint can show, and says so once per picture', async () => {
    const { els, warnings } = await write([
      picture({ pptxId: '4', dataUrl: undefined, unsupported: 'emf' }),
      picture({ pptxId: '5', dataUrl: 'data:image/svg+xml;base64,PHN2Zy8+', mime: 'image/svg+xml' }),
      picture({ pptxId: '6', dataUrl: 'data:image/png;base64,@@@' }),
      picture({ pptxId: '8', dataUrl: 'data:image/png;base64,AAAAA' }),
      picture({ pptxId: '7', dataUrl: 'not a data url' }),
    ]);
    expect(els).toEqual([]);
    expect(warnings).toEqual(['Changed on export: pictures with no embedded image PowerPoint can show (linked, missing, EMF, WMF or SVG without a PNG) left out (5).']);
  });

  it('writes a picture with no frame, geometry or outline', async () => {
    const { els } = await write([{ kind: 'picture', pptxId: '3', name: '', dataUrl: PNG, mime: 'image/png', part: 'ppt/media/a.png' }]);
    expect(els).toEqual([{ kind: 'picture', pptxId: '3', name: '', part: 'ppt/media/a.png', mime: 'image/png', dataUrl: PNG }]);
  });
});

describe('tables', () => {
  const cell = (t: string, over: Partial<PptxTableCell> = {}): PptxTableCell => ({ text: body(t), ...over });
  const table = (over: Partial<PptxTable> = {}): PptxTable => ({
    kind: 'table',
    pptxId: '4',
    name: 'Table',
    altText: 'numbers',
    frame: { x: 10, y: 20, w: 300, h: 80 },
    columns: [100, 200],
    rows: [
      { height: 40, cells: [cell('a', { gridSpan: 2, fill: { kind: 'solid', color: rgb('#112233') }, anchor: 'middle', margins: { l: 4, t: 2, r: 4, b: 2 } }), cell('', { hMerge: true })] },
      { height: 40, cells: [cell('c', { rowSpan: 2, vert: 'vert270', borders: { l: { width: 1, fill: { kind: 'solid', color: rgb('#FF0000') }, dash: 'dash' }, r: null, t: { width: 2 }, b: { width: 3 } } }), cell('d', { anchor: 'bottom' })] },
    ],
    flags: { firstRow: true, bandRow: true },
    styleId: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}',
    styleName: 'Medium Style 2 - Accent 1',
    ...over,
  });

  it('writes spans, borders, margins, anchors, fills, flags and the style, and reads back equal', async () => {
    const t = table();
    const { els, xml, part } = await write([t]);
    expect(els).toEqual([t]);
    expect(xml).toContain('gridSpan="2"');
    expect(xml).toContain('<a:lnR><a:noFill/></a:lnR>');
    expect(await part('ppt/tableStyles.xml')).toContain('<a:tblStyle styleId="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" styleName="Medium Style 2 - Accent 1"/>');
  });

  it('writes a table without a frame, a style or flags at the size of its grid, and a style without a name', async () => {
    const t = table({ frame: undefined, flags: undefined, styleId: undefined, styleName: undefined, altText: undefined, hidden: true });
    const { els, xml } = await write([t, table({ pptxId: '5', frame: { x: 0, y: 0, w: 1, h: 1 }, styleId: '{X}', styleName: undefined, flags: undefined }), table({ pptxId: '6', frame: { x: 0, y: 0, w: 1, h: 1 }, styleId: '{X}', styleName: 'Named' })]);
    expect(xml).toContain('<a:ext cx="2857500" cy="762000"/>');
    expect(els[0]).toMatchObject({ frame: { x: 0, y: 0, w: 300, h: 80 }, hidden: true });
  });

  it('keeps the first name of a table style several tables use', async () => {
    const { part } = await write([table({ styleName: undefined }), table({ pptxId: '5', styleName: 'Named' }), table({ pptxId: '6', styleName: 'Later' })]);
    expect(await part('ppt/tableStyles.xml')).toContain('styleName="Named"');
  });

  it('writes a table with no rows', async () => {
    const { els } = await write([table({ columns: [], rows: [] })]);
    expect(els).toHaveLength(1);
  });
});

describe('connectors', () => {
  it('glues the ends to the new ids of the shapes, and leaves an end free when its shape is not on the slide', async () => {
    const a = shape({ pptxId: 'first', name: 'A' });
    const b = shape({ pptxId: 'first', name: 'B (same id)' });
    const c = shape({ pptxId: 'x', name: 'C' });
    const cxn = { kind: 'connector' as const, pptxId: '7', name: 'Elbow', altText: 'link', frame: { x: 1, y: 2, w: 30, h: 40, flipV: true }, geometry: { kind: 'preset' as const, prst: 'bentConnector3', adjustments: { adj1: 50000 }, known: false }, line: { width: 2, tail: { type: 'triangle' } }, start: { id: 'first', idx: 3 }, end: { id: '99', idx: 0 } };
    const { els, xml } = await write([a, b, c, cxn]);
    const written = els.find((e) => e.kind === 'connector');
    expect(written).toEqual({ ...Object.fromEntries(Object.entries(cxn).filter(([k]) => k !== 'end')), start: { id: '2', idx: 3 } });
    expect(xml).toMatch(/<a:stCxn id="2" idx="3"\/>/);
    expect(xml).not.toContain('endCxn');
    const free = (await write([{ kind: 'connector', pptxId: '7', name: 'Free' }])).els;
    expect(free).toEqual([{ kind: 'connector', pptxId: '7', name: 'Free' }]);
  });
});

describe('groups', () => {
  const group = (over: Partial<PptxGroup> = {}): PptxGroup => ({ kind: 'group', pptxId: '20', name: 'Group', altText: 'g', frame: { x: 100, y: 100, w: 200, h: 100 }, children: [shape({ pptxId: '21', name: 'In', frame: { x: 120, y: 110, w: 50, h: 30 } })], ...over });
  const close = (a: PptxFrame | undefined, b: PptxFrame | undefined) => {
    expect(a).toBeDefined();
    for (const k of ['x', 'y', 'w', 'h', 'rot'] as const) expect(Math.abs((a?.[k] ?? 0) - (b?.[k] ?? 0))).toBeLessThan(0.01);
    expect(!!a?.flipH).toBe(!!b?.flipH);
    expect(!!a?.flipV).toBe(!!b?.flipV);
  };

  it('writes a group with its children in the group space and reads the same slide-space frames back, rotated, flipped and nested', async () => {
    const inner = group({ pptxId: '30', name: 'Inner', frame: { x: 150, y: 120, w: 80, h: 40, rot: 30, flipH: true }, children: [shape({ pptxId: '31', name: 'Deep', frame: { x: 160, y: 125, w: 20, h: 10, rot: 10, flipV: true } })] });
    const outer = group({ frame: { x: 100, y: 100, w: 200, h: 100, rot: 45, flipV: true }, children: [shape({ pptxId: '21', name: 'In', frame: { x: 120, y: 110, w: 50, h: 30, rot: 15 } }), inner] });
    const { els } = await write([outer]);
    const g = els[0] as PptxGroup;
    close(g.frame, outer.frame);
    close(g.children[0].frame, outer.children[0].frame);
    const i = g.children[1] as PptxGroup;
    close(i.frame, inner.frame);
    close(i.children[0].frame, inner.children[0].frame);
    expect(g).toMatchObject({ name: 'Group', altText: 'g', pptxId: '20' });
  });

  it('writes a hidden group with no frame: its children keep the space they are in, also inside a group with a frame', async () => {
    const loose = group({ frame: undefined, hidden: true, children: [shape({ pptxId: '21', frame: { x: 120, y: 110, w: 50, h: 30 } })] });
    const nested = group({ pptxId: '40', children: [group({ pptxId: '41', frame: undefined, children: [shape({ pptxId: '42', frame: { x: 110, y: 105, w: 10, h: 10 } })] })] });
    const { els } = await write([loose, nested]);
    expect((els[0] as PptxGroup).frame).toBeUndefined();
    expect((els[0] as PptxGroup).hidden).toBe(true);
    expect((els[0] as PptxGroup).children[0].frame).toEqual({ x: 120, y: 110, w: 50, h: 30 });
    const deep = ((els[1] as PptxGroup).children[0] as PptxGroup).children[0];
    close(deep.frame, { x: 110, y: 105, w: 10, h: 10 });
  });

  it('writes an empty group and children without a frame', async () => {
    const { els } = await write([group({ children: [] }), group({ pptxId: '50', children: [{ kind: 'shape', pptxId: '51', name: 'none' }] })]);
    expect((els[0] as PptxGroup).children).toEqual([]);
    expect((els[1] as PptxGroup).children[0].frame).toBeUndefined();
  });

  it('gives every element of a group its own id and leaves out opaque children', async () => {
    const { xml, els } = await write([group({ children: [shape({ pptxId: '2', name: 'dup' }), { kind: 'opaque', opaqueKind: 'chart', pptxId: '9', name: 'chart', parts: [] }] })]);
    expect((els[0] as PptxGroup).children).toHaveLength(1);
    expect(xml.match(/<p:cNvPr id="\d+"/g)).toHaveLength(3);
  });

  it('maps a frame into the child space of the group so the reader composes it back, for any rotation and flips', () => {
    const rot = [0, 30, 90, 135, 270, 359];
    let seed = 7;
    const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
    for (const r of rot) for (const fh of [false, true]) for (const fv of [false, true]) {
      const g: PptxFrame = { x: rnd() * 300, y: rnd() * 300, w: 50 + rnd() * 200, h: 50 + rnd() * 200, ...(r ? { rot: r } : {}), ...(fh ? { flipH: true } : {}), ...(fv ? { flipV: true } : {}) };
      const f: PptxFrame = { x: rnd() * 500, y: rnd() * 500, w: 10 + rnd() * 100, h: 10 + rnd() * 100, rot: rnd() * 360, ...(rnd() > 0.5 ? { flipH: true } : {}) };
      const c = toChildSpace(f, g);
      const xfrm = { off: { x: g.x, y: g.y }, ext: { w: g.w, h: g.h }, chOff: { x: g.x, y: g.y }, chExt: { w: g.w, h: g.h }, rotation: g.rot ?? 0, flipH: g.flipH ?? false, flipV: g.flipV ?? false };
      const back = composeGroupChain({ x: c.x, y: c.y, w: c.w, h: c.h, rotation: c.rot ?? 0, flipH: c.flipH ?? false, flipV: c.flipV ?? false }, [xfrm]);
      expect(back.x).toBeCloseTo(f.x, 6);
      expect(back.y).toBeCloseTo(f.y, 6);
      expect(back.w).toBeCloseTo(f.w, 6);
      expect(((back.rotation - (f.rot ?? 0)) % 360 + 360) % 360 < 1e-6 || Math.abs(((back.rotation - (f.rot ?? 0)) % 360 + 360) % 360 - 360) < 1e-6).toBe(true);
      expect(back.flipH).toBe(f.flipH ?? false);
      expect(back.flipV).toBe(f.flipV ?? false);
    }
  });
});

describe('custom geometry', () => {
  it('writes every path command, the fill mode and the stroke flags, and reads back equal', async () => {
    const geometry: PptxShape['geometry'] = {
      kind: 'custom',
      paths: [
        { w: 1000, h: 500, commands: [{ op: 'move', x: 0, y: 0 }, { op: 'line', x: 1000, y: 0 }, { op: 'cubic', x1: 1, y1: 2, x2: 3, y2: 4, x: 5, y: 6 }, { op: 'quad', x1: 7, y1: 8, x: 9, y: 10 }, { op: 'arc', wR: 100, hR: 50, stAng: 90, swAng: -45.5 }, { op: 'close' }] },
        { w: 10, h: 10, fill: 'darken', stroke: false, extrusionOk: true, commands: [{ op: 'move', x: 1, y: 1 }] },
      ],
    };
    const { els, warnings } = await write([shape({ geometry })]);
    expect((els[0] as PptxShape).geometry).toEqual(geometry);
    expect(warnings).toEqual([]);
  });

  it('writes whole numbers for coordinates that are not, and 0 for those that are not numbers', async () => {
    const geometry: PptxShape['geometry'] = { kind: 'custom', paths: [{ w: 0, h: 99.6, commands: [{ op: 'move', x: 1.4, y: Number.NaN }] }] };
    const { els } = await write([shape({ geometry })]);
    expect((els[0] as PptxShape).geometry).toEqual({ kind: 'custom', paths: [{ w: 1, h: 100, commands: [{ op: 'move', x: 1, y: 0 }] }] });
  });
});

describe('notes', () => {
  it('writes the notes master and a notes slide per slide with notes, and reads the text and its formatting back', async () => {
    const notes: PptxTextBody = {
      props: {},
      paragraphs: [
        { level: 0, style: { fontSize: 12 }, runs: [{ kind: 'text', text: 'Say ' }, { kind: 'text', text: 'this', bold: true, link: { kind: 'url', href: 'https://example.com/n' } }] },
        { level: 1, style: { fontSize: 12 }, runs: [{ kind: 'text', text: 'then that' }] },
      ],
    };
    const { zip, part, rels, back } = await write([shape()], { notes }, [slide({ pptxId: '257', elements: [] }), slide({ pptxId: '258', elements: [], notes: body('second') })]);
    expect(Object.keys(zip.files).filter((n) => /notes(Slide|Master)\d/.test(n) && !n.includes('_rels'))).toEqual(['ppt/notesSlides/notesSlide1.xml', 'ppt/notesSlides/notesSlide2.xml', 'ppt/notesMasters/notesMaster1.xml']);
    expect(rels).toContain('relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"');
    expect(await part('ppt/presentation.xml')).toContain('<p:notesMasterIdLst><p:notesMasterId r:id="');
    expect(await part('ppt/notesSlides/_rels/notesSlide2.xml.rels')).toContain('Target="../slides/slide3.xml"');
    expect(back.slides[0].notes).toEqual(notes);
    expect(back.slides[1].notes).toBeUndefined();
    expect(back.slides[2].notes?.paragraphs[0].runs[0].text).toBe('second');
  });

  it('writes no notes master for a deck without notes', async () => {
    const { zip, part } = await write([shape()]);
    expect(Object.keys(zip.files).some((n) => /notes/i.test(n))).toBe(false);
    expect(await part('ppt/presentation.xml')).not.toContain('notesMaster');
  });
});

describe('comments', () => {
  it('writes the authors once and each slide\'s threads with their replies, positions and times', async () => {
    const t1 = { id: '256-0.1', format: 'legacy' as const, anchor: { kind: 'slide' as const }, position: { x: 12.5, y: 40 }, comments: [{ author: 'Ana', initials: 'AB', body: 'First <&>', createdAt: '2024-03-01T10:20:30.000Z' }, { author: 'Bo', body: 'Reply\nline', createdAt: '2024-03-02T08:00:00.000Z' }, { author: 'Ana', initials: 'AB', body: 'Again' }] };
    const t2 = { id: '257-1.1', format: 'legacy' as const, anchor: { kind: 'slide' as const }, position: { x: 1, y: 2 }, comments: [{ author: 'Bo', body: 'Other slide' }] };
    const { part, back } = await write([shape()], { threads: [t1] }, [slide({ pptxId: '257', elements: [], threads: [t2] })]);
    const authors = await part('ppt/commentAuthors.xml');
    expect(authors).toContain('name="Ana" initials="AB" lastIdx="2"');
    expect(authors).toContain('name="Bo" initials="" lastIdx="2"');
    expect(await part('ppt/comments/comment1.xml')).toContain('p15:parentCm authorId="0" idx="1"');
    expect(back.slides[0].threads).toMatchObject([{ format: 'legacy', position: { x: 12.5, y: 40 }, comments: t1.comments }]);
    expect(back.slides[1].threads).toMatchObject([{ position: { x: 1, y: 2 }, comments: t2.comments }]);
  });

  it('writes a modern comment as a legacy one at the slide and says so', async () => {
    const modern = { id: 'm1', format: 'modern' as const, anchor: { kind: 'element' as const, pptxId: '2' }, resolved: true, comments: [{ author: 'Cy', body: 'Hi' }] };
    const { back, warnings } = await write([shape()], { threads: [modern] });
    expect(back.slides[0].threads).toMatchObject([{ format: 'legacy', anchor: { kind: 'slide' }, comments: [{ author: 'Cy', body: 'Hi' }] }]);
    expect(warnings).toEqual(['Changed on export: modern comments written as legacy comments on the slide (status and the shape a comment was anchored to are not kept) (1).']);
  });

  it('writes no comment parts for slides without threads, or with an empty list', async () => {
    const { zip } = await write([shape()], { threads: [] });
    expect(Object.keys(zip.files).some((n) => /comment/i.test(n))).toBe(false);
  });
});
