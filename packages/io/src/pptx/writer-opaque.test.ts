import JSZip from 'jszip';
import type { OpaquePart, OpaqueRef } from '@b32nio/spindle-shared/opaque';
import type { PptxElement, PptxOpaque, PptxSlide, PresentationImportResult } from '@b32nio/spindle-shared/export/pptx';
import { DEFAULT_COLOR_MAP } from './color';
import { importPptx } from './reader';
import { FALLBACK_THEME } from './theme';
import { exportPptx } from './writer';
import { lintPptxPackage } from './writer-lint';
import { relativeTarget, slideFragmentsXml } from './writer-opaque';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNG_BYTES = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0));
const enc = new TextEncoder();
const dec = new TextDecoder();
const layout = { partName: 'ppt/slideLayouts/slideLayout1.xml', name: 'Blank', layoutType: 'blank', masterPart: 'ppt/slideMasters/slideMaster1.xml', placeholders: [], elements: [], staticShapes: 0 };
const rels = (list: Array<[string, string, string?]>): string =>
  `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join('')}</Relationships>`;

/** The parts of one element: `[role, name, content, contentType]`; the hash is the name (the writer only looks parts up by it). */
type Spec = [string, string, string | Uint8Array, string?];
function parts(specs: Spec[]): { refs: OpaqueRef[]; store: OpaquePart[] } {
  const refs: OpaqueRef[] = [];
  const store: OpaquePart[] = [];
  for (const [role, name, content, type] of specs) {
    const hash = `hash:${name}`;
    refs.push({ hash, partName: name, role });
    store.push({ hash, contentType: type ?? 'application/xml', bytes: typeof content === 'string' ? enc.encode(content) : content });
  }
  return { refs, store };
}
const opaque = (over: Partial<PptxOpaque> & Pick<PptxOpaque, 'opaqueKind'>, refs: OpaqueRef[]): PptxOpaque => ({
  kind: 'opaque',
  pptxId: '4',
  name: 'Opaque 3',
  frame: { x: 100, y: 50, w: 400, h: 300 },
  parts: refs,
  ...over,
});
const slide = (elements: PptxElement[], over: Partial<PptxSlide> = {}): PptxSlide => ({ pptxId: '256', partName: 'ppt/slides/slide1.xml', layoutPart: layout.partName, elements, ...over });
const deck = (slides: PptxSlide[], store: OpaquePart[]): PresentationImportResult => ({ slideSize: { w: 1280, h: 720 }, theme: FALLBACK_THEME, colorMap: DEFAULT_COLOR_MAP, slides, layouts: [layout], fonts: { used: [], unavailable: [] }, opaqueParts: store, warnings: [] });

async function write(slides: PptxSlide[], store: OpaquePart[]) {
  const { bytes, warnings } = await exportPptx(deck(slides, store));
  const zip = await JSZip.loadAsync(bytes);
  const text = async (n: string) => (await zip.file(n)?.async('string')) as string;
  return { bytes, zip, warnings, text, lint: await lintPptxPackage(zip), slideXml: await text('ppt/slides/slide1.xml'), slideRels: await text('ppt/slides/_rels/slide1.xml.rels'), types: await text('[Content_Types].xml') };
}
const bytesOf = async (zip: JSZip, name: string) => new Uint8Array((await zip.file(name)?.async('uint8array')) as Uint8Array);

const CHART = parts([
  ['chart', 'ppt/charts/chart1.xml', '<c:chartSpace xmlns:c="c" xmlns:r="r"><c:externalData r:id="rId1"/></c:chartSpace>', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'],
  ['rels', 'ppt/charts/_rels/chart1.xml.rels', rels([['rId1', `${R}/package`, '../embeddings/Microsoft_Excel_Sheet1.xlsx']])],
  ['package', 'ppt/embeddings/Microsoft_Excel_Sheet1.xlsx', new Uint8Array([80, 75, 3, 4, 0, 255, 128]), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
]);

describe('chart', () => {
  it('writes the frame, the chart and its workbook byte for byte under their names, and a package the lints pass', async () => {
    const out = await write([slide([opaque({ opaqueKind: 'chart' }, CHART.refs)])], CHART.store);
    expect(out.lint).toEqual([]);
    expect(out.warnings).toEqual([]);
    expect(out.slideXml).toContain('<p:cNvPr id="4" name="Opaque 3"/>');
    expect(out.slideXml).toContain('<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/>');
    expect(out.slideRels).toContain(`Id="rId2" Type="${R}/chart" Target="../charts/chart1.xml"`);
    for (const [, name, content] of [['', 'ppt/charts/chart1.xml', CHART.store[0].bytes], ['', 'ppt/charts/_rels/chart1.xml.rels', CHART.store[1].bytes], ['', 'ppt/embeddings/Microsoft_Excel_Sheet1.xlsx', CHART.store[2].bytes]] as const) {
      expect(await bytesOf(out.zip, name)).toEqual(content);
    }
    expect(out.types).toContain('<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>');
    expect(out.types).toContain('<Override PartName="/ppt/embeddings/Microsoft_Excel_Sheet1.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>');
    expect(out.types).not.toContain('chart1.xml.rels');
  });

  it('reads back as the same element with the same parts (the part names and hashes survive)', async () => {
    const out = await write([slide([opaque({ opaqueKind: 'chart' }, CHART.refs)])], CHART.store);
    const back = await importPptx(out.bytes);
    const e = back.slides[0].elements?.[0] as PptxOpaque;
    expect(e).toMatchObject({ kind: 'opaque', opaqueKind: 'chart', pptxId: '4', name: 'Opaque 3' });
    expect(e.parts.map((p) => [p.role, p.partName])).toEqual(CHART.refs.map((p) => [p.role, p.partName]));
    expect(e.frame).toEqual({ x: 100, y: 50, w: 400, h: 300 });
    expect(back.opaqueParts.map((p) => dec.decode(p.bytes).length)).toEqual(CHART.store.map((p) => p.bytes.length));
  });

  it('keeps one copy of a part two elements share and one workbook per chart', async () => {
    const two = parts([['chart', 'ppt/charts/chart1_2.xml', '<c:chartSpace xmlns:c="c"/>'], ['package', 'ppt/embeddings/Microsoft_Excel_Sheet1.xlsx', new Uint8Array([1])]]);
    const dupe = [{ ...two.refs[0], hash: CHART.refs[0].hash, partName: 'ppt/charts/chart1.xml' }];
    const out = await write([slide([opaque({ opaqueKind: 'chart' }, CHART.refs), opaque({ opaqueKind: 'chart', pptxId: '5' }, dupe)])], CHART.store);
    expect(out.lint).toEqual([]);
    expect(Object.keys(out.zip.files).filter((n) => n.startsWith('ppt/charts/') && !n.endsWith('/'))).toEqual(['ppt/charts/chart1.xml', 'ppt/charts/_rels/chart1.xml.rels']);
    expect(out.slideXml.match(/r:id="rId2"/g)).toHaveLength(2);
  });
});

describe('part names', () => {
  it('moves a part named like one of the writer\'s own under ppt/opaque, and numbers a taken name', async () => {
    const p = parts([
      ['chart', 'ppt/charts/chart1.xml', '<c:chartSpace xmlns:c="c" xmlns:r="r"><c:x r:id="rId1"/><c:y r:id="rId2"/></c:chartSpace>'],
      ['rels', 'ppt/charts/_rels/chart1.xml.rels', rels([['rId1', `${R}/image`, '../slides/slide1.xml'], ['rId2', `${R}/image`, '../media/image1.png']])],
      ['image', 'ppt/slides/slide1.xml', '<img/>'],
      ['image', 'ppt/media/image1.png', new Uint8Array([9, 9]), 'image/png'],
    ]);
    // a picture of the slide has taken ppt/media/image1.png, a second chart claims the same name with other bytes
    const second = parts([['chart', 'ppt/charts/chart1.xml', '<c:chartSpace xmlns:c="c"/>']]);
    second.refs[0].hash = 'other';
    second.store[0].hash = 'other';
    const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
    const pic: PptxElement = { kind: 'picture', pptxId: '9', name: 'p', frame: { x: 0, y: 0, w: 10, h: 10 }, part: 'ppt/media/image1.png', mime: 'image/png', dataUrl: png };
    const out = await write([slide([pic, opaque({ opaqueKind: 'chart' }, p.refs), opaque({ opaqueKind: 'chart', pptxId: '5' }, second.refs)])], [...p.store, ...second.store]);
    expect(out.lint).toEqual([]);
    expect(out.zip.file('ppt/opaque/slide1.xml')).not.toBeNull();
    expect(out.zip.file('ppt/charts/chart1.xml')).not.toBeNull();
    expect(out.zip.file('ppt/charts/chart1_2.xml')).not.toBeNull();
    // the picture kept its media name, so the opaque image of the same name was numbered
    expect(out.zip.file('ppt/media/image1.png')).not.toBeNull();
    expect(out.zip.file('ppt/media/image1_2.png')).not.toBeNull();
    const chartRels = await out.text('ppt/charts/_rels/chart1.xml.rels');
    expect(chartRels).toContain('Target="../opaque/slide1.xml"');
    expect(chartRels).toContain('Target="../media/image1_2.png"');
    expect(chartRels).not.toContain('../slides/');
  });

  it('gives a part with an unusable name (a space, climbing out, none) a generic one', async () => {
    const p = parts([['chart', 'ppt/charts/chart1.xml', '<c:chartSpace xmlns:c="c" xmlns:r="r"/>'], ['rels', 'ppt/charts/_rels/chart1.xml.rels', rels([['rId1', `${R}/package`, '../embeddings/a b.xlsx'], ['rId2', `${R}/package`, '../../../x.bin'], ['rId3', `${R}/package`, '../embeddings/c.bin']])], ['package', 'ppt/embeddings/a b.xlsx', new Uint8Array([1])], ['package', '../../../x.bin', new Uint8Array([2])], ['package', 'ppt/embeddings/c.bin', new Uint8Array([3])]]);
    const nameless = { hash: 'nameless', role: 'package' };
    const store = [...p.store, { hash: 'nameless', contentType: 'application/octet-stream', bytes: new Uint8Array([4]) }];
    const out = await write([slide([opaque({ opaqueKind: 'chart' }, [...p.refs, nameless])])], store);
    // the climbing name and the nameless part are not the target of any relationship (the name cannot be resolved), so the lint says so
    expect(out.lint).toEqual(['ppt/opaque/part2.bin: no relationship targets this part', 'ppt/opaque/part3: no relationship targets this part']);
    expect(await out.text('ppt/charts/_rels/chart1.xml.rels')).toContain('Target="../opaque/part1.xlsx"');
    expect(out.zip.file('ppt/embeddings/c.bin')).not.toBeNull();
  });
});

describe('relationship rewriting', () => {
  it('leaves the relationship part as it was when no target changed name', async () => {
    const out = await write([slide([opaque({ opaqueKind: 'chart' }, CHART.refs)])], CHART.store);
    expect(await out.text('ppt/charts/_rels/chart1.xml.rels')).toBe(dec.decode(CHART.store[1].bytes));
  });

  it('leaves out a relationship to a part that was not kept, and an external one stays', async () => {
    const p = parts([
      ['chart', 'ppt/charts/chart1.xml', '<c:chartSpace xmlns:c="c"/>'],
      ['rels', 'ppt/charts/_rels/chart1.xml.rels', `${rels([['rId1', `${R}/package`, '../embeddings/gone.xlsx']]).replace('</Relationships>', '')}<Relationship Id="rId9" Type="${R}/hyperlink" Target="https://example.com/?a=1&amp;b=2" TargetMode="External"/></Relationships>`],
    ]);
    const out = await write([slide([opaque({ opaqueKind: 'chart' }, p.refs)])], p.store);
    expect(out.lint).toEqual([]);
    const text = await out.text('ppt/charts/_rels/chart1.xml.rels');
    expect(text).not.toContain('gone.xlsx');
    expect(text).toContain('TargetMode="External"');
    expect(out.warnings).toEqual(['Changed on export: relationships of kept content to parts that are missing left out (1).']);
  });

  it('reports a part the store no longer holds and one whose relationship part has no usable name', async () => {
    const lost = parts([['chart', 'ppt/charts/chart1.xml', '<c:chartSpace xmlns:c="c"/>'], ['package', 'ppt/embeddings/w.xlsx', 'x']]);
    const orphanRels = { hash: 'orphan', role: 'rels' };
    const orphanOwner = { hash: 'o2', role: 'rels', partName: 'ppt/charts/_rels/other.xml.rels' };
    const store = [...lost.store.slice(0, 1), { hash: 'orphan', contentType: 'x', bytes: enc.encode('<Relationships/>') }, { hash: 'o2', contentType: 'x', bytes: enc.encode('<Relationships/>') }];
    const out = await write([slide([opaque({ opaqueKind: 'chart' }, [...lost.refs, orphanRels, orphanOwner])])], store);
    expect(out.warnings).toEqual(['Changed on export: relationships of kept content to parts that are missing left out (1).']);
    expect(out.zip.file('ppt/embeddings/w.xlsx')).toBeNull();
  });

  it('reads a relationship written with single quotes and keeps a part whose relationships point back (no part is then a root, the first is)', async () => {
    const cycle = parts([
      ['ink', 'ppt/ink/a.xml', '<a/>', ''],
      ['rels', 'ppt/ink/_rels/a.xml.rels', `<Relationships xmlns='x'><Relationship Id='rId1' Type='${R}/image' Target='b.xml' TargetMode='Internal'/></Relationships>`],
      ['image', 'ppt/ink/b.xml', '<b/>'],
      ['rels', 'ppt/ink/_rels/b.xml.rels', rels([['rId1', `${R}/ink`, 'a.xml']])],
    ]);
    const out = await write([slide([opaque({ opaqueKind: 'ink', frame: undefined }, cycle.refs)])], cycle.store);
    expect(out.lint).toEqual([]);
    expect(out.slideRels).toContain('Target="../ink/a.xml"');
    expect(out.slideRels).not.toContain('b.xml');
    expect(out.slideXml).toContain('<p14:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></p14:xfrm>');
    expect(out.types).toContain('<Override PartName="/ppt/ink/a.xml" ContentType="application/octet-stream"/>');
  });

  it('writes the target relative to the new place of the owner when a part moved', () => {
    expect(relativeTarget('ppt/slides/slide1.xml', 'ppt/charts/chart1.xml')).toBe('../charts/chart1.xml');
    expect(relativeTarget('ppt/charts/chart1.xml', 'ppt/charts/style1.xml')).toBe('style1.xml');
    expect(relativeTarget('a.xml', 'ppt/b.xml')).toBe('ppt/b.xml');
    expect(relativeTarget('ppt/a/b/c.xml', 'ppt/d.xml')).toBe('../../d.xml');
  });
});

describe('elements that cannot be written', () => {
  it('leaves out an element whose root part the store does not hold, and says so', async () => {
    const out = await write([slide([opaque({ opaqueKind: 'chart' }, CHART.refs)])], []);
    expect(out.lint).toEqual([]);
    expect(out.slideXml).not.toContain('graphicFrame');
    expect(out.warnings).toEqual(['Changed on export: charts, diagrams, objects, media and ink left out because the parts they are made of are missing (1).']);
  });

  it('leaves out linked media and objects and ActiveX controls, never writing a link', async () => {
    const clip = parts([['video', 'ppt/media/clip.mp4', 'mp4', 'video/mp4']]);
    const linked = opaque({ opaqueKind: 'media', props: { clip: 'videoFile', linkedTarget: 'file:///C:/clip.mp4' } }, clip.refs);
    const control = opaque({ opaqueKind: 'ole', pptxId: '6', props: { control: 'true' } }, clip.refs);
    const out = await write([slide([linked, control])], clip.store);
    expect(out.lint).toEqual([]);
    expect(out.slideRels).not.toContain('file:');
    expect(out.warnings).toEqual([
      'Changed on export: ActiveX controls left out (their binary parts are not written) (1).',
      'Changed on export: media and objects linked from outside the file left out (the link is never written) (1).',
    ]);
  });

  it('writes nothing for an element when the writer has no emitter for it (a context without opaque parts)', async () => {
    // layouts have no emitter: a layout's elements are shapes, so this is reached through the dispatcher only by a hand-made context
    const { elementsXml } = await import('./writer-elements');
    const lossy: string[] = [];
    const ctx = { rels: undefined, slideTarget: () => undefined, lossy: (k: string) => lossy.push(k), image: () => undefined, tableStyle: () => undefined, nextFieldId: () => '' };
    const e = opaque({ opaqueKind: 'chart' }, []);
    expect(elementsXml([e], ctx as never, { ids: new Map([[e, 4]]), bySource: new Map() })).toBe('');
    expect(lossy).toEqual(['opaque']);
  });

  it('writes an element inside a group in the group\'s space, once', async () => {
    const group: PptxElement = { kind: 'group', pptxId: '7', name: 'g', frame: { x: 50, y: 50, w: 600, h: 400 }, children: [opaque({ opaqueKind: 'chart', frame: { x: 100, y: 100, w: 200, h: 100 } }, CHART.refs)] };
    const out = await write([slide([group])], CHART.store);
    expect(out.lint).toEqual([]);
    expect(out.warnings).toEqual([]);
    const back = await importPptx(out.bytes);
    const inner = (back.slides[0].elements?.[0] as { children: PptxOpaque[] }).children[0];
    expect(inner.opaqueKind).toBe('chart');
    expect(inner.frame).toEqual({ x: 100, y: 100, w: 200, h: 100 });
  });

  it('leaves a connector end free when its shape was an element that could not be written, and glues it to one that was', async () => {
    const kept = opaque({ opaqueKind: 'chart', pptxId: '4' }, CHART.refs);
    const lost = opaque({ opaqueKind: 'chart', pptxId: '5', parts: [] }, []);
    const cxn: PptxElement = { kind: 'connector', pptxId: '6', name: 'c', frame: { x: 0, y: 0, w: 5, h: 5 }, start: { id: '4', idx: 1 }, end: { id: '5', idx: 2 } };
    const out = await write([slide([kept, lost, cxn])], CHART.store);
    expect(out.slideXml).toContain('<a:stCxn id="4" idx="1"/>');
    expect(out.slideXml).not.toContain('endCxn');
  });
});

const DGM = parts([
  ['diagramData', 'ppt/diagrams/data1.xml', '<dgm:dataModel xmlns:dgm="d"><dgm:extLst><a:ext xmlns:a="a"><dsp:dataModelExt xmlns:dsp="s" relId="rId1" minVer="m"/></a:ext></dgm:extLst></dgm:dataModel>'],
  ['diagramLayout', 'ppt/diagrams/layout1.xml', '<dgm:layoutDef xmlns:dgm="d"/>'],
  ['diagramQuickStyle', 'ppt/diagrams/quickStyle1.xml', '<dgm:styleDef xmlns:dgm="d"/>'],
  ['diagramColors', 'ppt/diagrams/colors1.xml', '<dgm:colorsDef xmlns:dgm="d"/>'],
  ['diagramDrawing', 'ppt/diagrams/drawing1.xml', '<dsp:drawing xmlns:dsp="s"/>'],
]);

describe('SmartArt', () => {
  it('writes the four relationship ids and the cached drawing under the id the data names, even when that is the layout\'s', async () => {
    const text: PptxElement = { kind: 'shape', pptxId: '2', name: 't', frame: { x: 0, y: 0, w: 10, h: 10 }, link: { kind: 'url', href: 'https://example.com/' } };
    const out = await write([slide([text, opaque({ opaqueKind: 'smartArt' }, DGM.refs)])], DGM.store);
    expect(out.lint).toEqual([]);
    expect(out.slideRels).toContain(`Id="rId1" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"`);
    expect(out.slideXml).toMatch(/<dgm:relIds [^>]*r:dm="rId\d" r:lo="rId\d" r:qs="rId\d" r:cs="rId\d"\/>/);
    expect(out.slideRels).toContain(`Type="${R}/slideLayout"`);
    const back = await importPptx(out.bytes);
    const e = back.slides[0].elements?.find((x) => x.kind === 'opaque') as PptxOpaque;
    expect(e.parts.map((p) => p.role)).toEqual(['diagramData', 'diagramLayout', 'diagramQuickStyle', 'diagramColors', 'diagramDrawing']);
  });

  it('takes the next free id for the drawing when the data names none or one that is taken', async () => {
    const noExt = parts([['diagramData', 'ppt/diagrams/data1.xml', '<dgm:dataModel xmlns:dgm="d"/>'], ['diagramDrawing', 'ppt/diagrams/drawing1.xml', '<d/>']]);
    const out = await write([slide([opaque({ opaqueKind: 'smartArt' }, noExt.refs)])], noExt.store);
    expect(out.lint).toEqual([]);
    expect(out.slideXml).toContain('r:dm="rId');
    expect(out.slideXml).not.toContain('r:lo=');
  });
});

describe('OLE objects', () => {
  const OLE = parts([
    ['oleObject', 'ppt/embeddings/oleObject1.bin', new Uint8Array([208, 207, 17]), 'application/vnd.openxmlformats-officedocument.oleObject'],
    ['image', 'ppt/media/image1.png', PNG_BYTES, 'image/png'],
  ]);
  it('writes the object with its preview and the progId and name', async () => {
    const e = opaque({ opaqueKind: 'ole', props: { progId: 'Excel.Sheet.12', oleName: 'Worksheet' } }, OLE.refs);
    const out = await write([slide([e, { ...e, pptxId: '5' }])], OLE.store);
    expect(out.lint).toEqual([]);
    expect(out.slideXml).toContain('<p:oleObj name="Worksheet" progId="Excel.Sheet.12" r:id="rId3"><p:embed/><p:pic>');
    expect(out.slideRels).toContain(`Type="${R}/oleObject" Target="../embeddings/oleObject1.bin"`);
    expect(out.slideRels).toContain(`Type="${R}/image" Target="../media/image1.png"`);
    const back = await importPptx(out.bytes);
    const read = back.slides[0].elements?.[0] as PptxOpaque;
    expect(read.opaqueKind).toBe('ole');
    expect(read.props).toEqual({ progId: 'Excel.Sheet.12', oleName: 'Worksheet' });
    expect(read.parts.map((p) => p.role)).toEqual(['oleObject', 'image']);
  });
  it('writes an object with no preview and no name', async () => {
    const e = opaque({ opaqueKind: 'ole' }, [OLE.refs[0]]);
    const out = await write([slide([e])], OLE.store);
    expect(out.lint).toEqual([]);
    expect(out.slideXml).toContain('<p:oleObj r:id="rId2"><p:embed/></p:oleObj>');
  });
  it('leaves out an object that has only a preview', async () => {
    const out = await write([slide([opaque({ opaqueKind: 'ole' }, [OLE.refs[1]])])], OLE.store);
    expect(out.warnings).toHaveLength(1);
  });
});

describe('audio and video', () => {
  const VIDEO = parts([['video', 'ppt/media/clip.mp4', 'mp4', 'video/mp4'], ['image', 'ppt/media/poster.png', PNG_BYTES, 'image/png']]);
  it('writes the clip under both relationships (the clip\'s and the media extension\'s) and the poster, and reads back equal', async () => {
    const out = await write([slide([opaque({ opaqueKind: 'media', props: { clip: 'videoFile' }, altText: 'a clip' }, VIDEO.refs)])], VIDEO.store);
    expect(out.lint).toEqual([]);
    expect(out.slideXml).toContain('<a:videoFile r:link="rId2"/>');
    expect(out.slideXml).toContain('<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId3"/>');
    expect(out.slideRels).toContain(`Id="rId2" Type="${R}/video" Target="../media/clip.mp4"`);
    expect(out.slideRels).toContain('Id="rId3" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/clip.mp4"');
    const back = await importPptx(out.bytes);
    const e = back.slides[0].elements?.[0] as PptxOpaque;
    expect(e).toMatchObject({ opaqueKind: 'media', altText: 'a clip', props: { clip: 'videoFile' } });
    expect(e.parts.map((p) => [p.role, p.partName])).toEqual([['video', 'ppt/media/clip.mp4'], ['image', 'ppt/media/poster.png']]);
    expect(e.poster).toMatch(/^data:image\/png/);
  });
  it('writes a video clip with no stated clip kind as a video, with a frame of zero when it has none', async () => {
    const video = parts([['video', 'ppt/media/v.mp4', 'v', 'video/mp4']]);
    const out = await write([slide([opaque({ opaqueKind: 'media', frame: undefined }, video.refs)])], video.store);
    expect(out.slideXml).toContain('<a:videoFile r:link="rId2"/>');
    expect(out.slideXml).toContain('<a:ext cx="0" cy="0"/>');
    const chart = await write([slide([opaque({ opaqueKind: 'chart', frame: undefined }, CHART.refs)])], CHART.store);
    expect(chart.slideXml).toContain('<p:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></p:xfrm>');
  });
  it('chooses the clip element from the role when the element does not say, and writes no poster when there is none', async () => {
    const audio = parts([['audio', 'ppt/media/s.wav', 'wav', 'audio/wav']]);
    const out = await write([slide([opaque({ opaqueKind: 'media' }, audio.refs), opaque({ opaqueKind: 'media', pptxId: '5', props: { clip: 'bogus' } }, audio.refs)])], audio.store);
    expect(out.lint).toEqual([]);
    expect(out.slideXml).toContain('<a:audioFile r:link="rId2"/>');
    expect(out.slideXml.match(/<a:audioFile/g)).toHaveLength(2);
    expect(out.slideXml).not.toContain('videoFile');
    expect(out.slideXml).toContain('<a:blip/>');
    expect(out.slideRels).toContain(`Type="${R}/audio"`);
  });
});

describe('ink', () => {
  const INK = parts([['ink', 'ppt/ink/ink1.xml', '<inkml:ink xmlns:inkml="i"/>', 'application/inkml+xml']]);
  it('writes the content part inside the alternate content, and reads back equal', async () => {
    const out = await write([slide([opaque({ opaqueKind: 'ink', hidden: true }, INK.refs)])], INK.store);
    expect(out.lint).toEqual([]);
    expect(out.slideXml).toContain('<mc:AlternateContent');
    expect(out.slideXml).toContain('Requires="p14"');
    expect(out.slideRels).toContain('Type="http://schemas.microsoft.com/office/2007/relationships/ink" Target="../ink/ink1.xml"');
    const back = await importPptx(out.bytes);
    const e = back.slides[0].elements?.[0] as PptxOpaque;
    expect(e).toMatchObject({ opaqueKind: 'ink', pptxId: '4', hidden: true, frame: { x: 100, y: 50, w: 400, h: 300 } });
    expect(e.parts.map((p) => p.partName)).toEqual(['ppt/ink/ink1.xml']);
  });
});

describe('transitions and timing', () => {
  const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  const ids = new Map([['4', 4], ['5', 9]]);
  const run = (o: { transition?: string; timing?: string }, m = ids) => {
    const lossy: string[] = [];
    return { xml: slideFragmentsXml(o, m, (k) => lossy.push(k)), lossy };
  };
  const timing = (spid: string) => `<p:timing ${P}><p:tnLst><p:par><p:cTn id="1"><p:childTnLst><p:set><p:cBhvr><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:set></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
  const fade = `<p:transition ${P} spd="med"><p:fade/></p:transition>`;

  it('writes a transition and a timing as they are, the transition first', () => {
    expect(run({ timing: timing('4'), transition: fade })).toEqual({ xml: fade + timing('4'), lossy: [] });
    expect(run({}).xml).toBe('');
    expect(slideFragmentsXml(undefined, ids, () => undefined)).toBe('');
  });
  it('writes the alternate content of a newer transition as it is', () => {
    const alt = `<mc:AlternateContent xmlns:mc="m" ${P}><mc:Choice xmlns:p14="q" Requires="p14"><p:transition><p14:vortex/></p:transition></mc:Choice><mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`;
    expect(run({ transition: alt })).toEqual({ xml: alt, lossy: [] });
  });
  it('cuts the sound out of a transition and leaves out a transition that is malformed, of another kind, or that holds an entity', () => {
    const sound = `<p:transition ${P} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:fade/><p:sndAc><p:stSnd><p:snd r:embed="rId2" name="a"/></p:stSnd></p:sndAc></p:transition>`;
    expect(run({ transition: sound })).toEqual({ xml: `<p:transition ${P} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:fade/></p:transition>`, lossy: ['transition'] });
    const loose = `<p:transition ${P} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:fade r:id="rId5"/></p:transition>`;
    for (const bad of ['<p:transition', '<p:timing/>', '<!DOCTYPE x [<!ENTITY e "x">]><p:transition/>', '<?pi?><p:transition/>', 'text', loose]) expect(run({ transition: bad })).toEqual({ xml: '', lossy: ['transition'] });
  });
  it('reads a shape id written with single quotes', () => {
    expect(run({ timing: timing('4').replace('spid="4"', "spid='4'") }).lossy).toEqual([]);
    expect(run({ timing: timing('5').replace('spid="5"', "spid='5'") }).lossy).toEqual(['timing']);
  });
  it('leaves out a timing that names a shape under another id or none, holds a sound, or is malformed', () => {
    expect(run({ timing: timing('5') })).toEqual({ xml: '', lossy: ['timing'] });
    expect(run({ timing: timing('77') })).toEqual({ xml: '', lossy: ['timing'] });
    expect(run({ timing: `<p:timing ${P} xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:snd rel:embed="rId2"/></p:timing>` })).toEqual({ xml: '', lossy: ['timing'] });
    expect(run({ timing: '<p:timing>' })).toEqual({ xml: '', lossy: ['timing'] });
    expect(run({ timing: fade })).toEqual({ xml: '', lossy: ['timing'] });
  });
  it('is written into the slide after the colour map override, in a package the lints pass', async () => {
    const e = opaque({ opaqueKind: 'chart' }, CHART.refs);
    const out = await write([slide([e], { opaque: { transition: fade, timing: timing('4') } })], CHART.store);
    expect(out.lint).toEqual([]);
    expect(out.slideXml).toMatch(/<\/p:clrMapOvr><p:transition [^>]*><p:fade\/><\/p:transition><p:timing /);
    const back = await importPptx(out.bytes);
    expect(back.slides[0].opaque?.transition).toContain('<p:fade/>');
    expect(back.slides[0].opaque?.timing).toContain('spid="4"');
  });
  it('leaves out a timing whose target shape is an element that could not be written', async () => {
    const out = await write([slide([opaque({ opaqueKind: 'chart' }, CHART.refs)], { opaque: { timing: timing('4') } })], []);
    expect(out.slideXml).not.toContain('p:timing');
    expect(out.warnings).toEqual(expect.arrayContaining([expect.stringContaining('animation timelines left out')]));
  });
});
