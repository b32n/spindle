import JSZip from 'jszip';
import type { PptxElement, PptxOpaque, PresentationImportResult } from '@b32nio/spindle-shared/export/pptx';
import { importPptx } from './reader';
import { detectPptxLoss } from './loss';
import { slideFragments } from './fragments';
import { readContentTypes } from './opaque';
import { parseXml } from '../xlsx/xml';

// Opaque content (plan K11): charts, SmartArt, OLE objects and controls, audio and video, ink, and the slide's transition and timing, read from synthetic decks that are real zips.

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PRES_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
const CLR_MAP = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
const THEME = `<a:theme ${NS} name="T"><a:themeElements><a:clrScheme name="c">${['dk1:000000', 'lt1:FFFFFF', 'dk2:44546A', 'lt2:E7E6E6', 'accent1:4472C4', 'accent2:ED7D31', 'accent3:A5A5A5', 'accent4:FFC000', 'accent5:5B9BD5', 'accent6:70AD47', 'hlink:0563C1', 'folHlink:954F72']
  .map((c) => `<a:${c.split(':')[0]}><a:srgbClr val="${c.split(':')[1]}"/></a:${c.split(':')[0]}>`)
  .join('')}</a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;

const rel = (id: string, type: string, target: string, external = false): string => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;
const rels = (list: string[]): string => `<Relationships xmlns="${PKG_REL}">${list.join('')}</Relationships>`;

interface Deck {
  tree?: string;
  /** Children of `p:sld` after `p:cSld`. */
  after?: string;
  controls?: string;
  slideRels?: string[];
  extra?: Record<string, string | Uint8Array>;
  types?: string;
}

async function build(d: Deck): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="${CT}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="${PRES_TYPE}"/>${d.types ?? ''}</Types>`);
  zip.file('_rels/.rels', rels([rel('rId1', 'officeDocument', 'ppt/presentation.xml')]));
  zip.file('ppt/_rels/presentation.xml.rels', rels([rel('rId1', 'slide', 'slides/slide1.xml'), rel('rIdM', 'slideMaster', 'slideMasters/slideMaster1.xml')]));
  zip.file('ppt/presentation.xml', `<p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${NS}><p:cSld><p:spTree>${d.tree ?? ''}</p:spTree>${d.controls ?? ''}</p:cSld>${d.after ?? ''}</p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', rels([rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'), ...(d.slideRels ?? [])]));
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout ${NS}><p:cSld name="L"><p:spTree/></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels([rel('rId1', 'slideMaster', '../slideMasters/slideMaster1.xml')]));
  zip.file('ppt/slideMasters/slideMaster1.xml', `<p:sldMaster ${NS}><p:cSld><p:spTree/></p:cSld>${CLR_MAP}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`);
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', rels([rel('rIdT', 'theme', '../theme/theme1.xml'), rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml')]));
  zip.file('ppt/theme/theme1.xml', THEME);
  for (const [name, content] of Object.entries(d.extra ?? {})) zip.file(name, content);
  return zip.generateAsync({ type: 'uint8array' });
}

const read = async (d: Deck, options?: Parameters<typeof importPptx>[1]): Promise<PresentationImportResult> => importPptx(await build(d), options);
const opaqueOf = (r: PresentationImportResult): PptxOpaque[] => (r.slides[0].elements as PptxElement[]).filter((e): e is PptxOpaque => e.kind === 'opaque');
const roles = (e: PptxOpaque): string[] => e.parts.map((p) => `${p.role}:${p.partName}`);
const PNG = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3]);
const frame = (id: string, name: string, uri: string, data: string, extra = ''): string =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${name}"${extra}/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="952500" y="1905000"/><a:ext cx="4762500" cy="2857500"/></p:xfrm><a:graphic><a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic></p:graphicFrame>`;
const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const chartFrame = (id: string, rid: string, extra = ' descr="Sales by region"'): string => frame(id, `Chart ${id}`, CHART_URI, `<c:chart xmlns:c="${CHART_URI}" r:id="${rid}"/>`, extra);
const CHART_TITLED = `<c:chartSpace xmlns:c="${CHART_URI}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Sales </a:t></a:r><a:r><a:t>2024</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>`;
const CHART_PLAIN = `<c:chartSpace xmlns:c="${CHART_URI}"><c:chart><c:plotArea/></c:chart></c:chartSpace>`;
const WORKBOOK = Uint8Array.from([0x50, 0x4b, 3, 4, 9, 9, 9]);
const chartParts = (n: number, xml = CHART_TITLED): Record<string, string | Uint8Array> => ({
  [`ppt/charts/chart${n}.xml`]: xml,
  [`ppt/charts/_rels/chart${n}.xml.rels`]: rels([rel('rId1', 'package', `../embeddings/Microsoft_Excel_Sheet${n}.xlsx`), rel('rId2', 'chartStyle', `style${n}.xml`), rel('rId3', 'themeOverride', '../theme/themeOverride1.xml')]),
  [`ppt/embeddings/Microsoft_Excel_Sheet${n}.xlsx`]: WORKBOOK,
  [`ppt/charts/style${n}.xml`]: '<cs:chartStyle xmlns:cs="x"/>',
  'ppt/theme/themeOverride1.xml': '<a:themeOverride/>',
});

describe('charts', () => {
  const types = '<Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>';

  it('keeps the chart, its workbook, style parts and relationship parts, with the title, frame and alt text', async () => {
    const r = await read({ tree: chartFrame('3', 'rId2'), slideRels: [rel('rId2', 'chart', '../charts/chart1.xml')], extra: chartParts(1), types });
    const [chart] = opaqueOf(r);
    expect(chart).toMatchObject({ kind: 'opaque', opaqueKind: 'chart', pptxId: '3', name: 'Chart 3', altText: 'Sales by region', title: 'Sales 2024' });
    expect(chart.frame).toEqual({ x: 100, y: 200, w: 500, h: 300 });
    expect(roles(chart)).toEqual([
      'chart:ppt/charts/chart1.xml',
      'rels:ppt/charts/_rels/chart1.xml.rels',
      'package:ppt/embeddings/Microsoft_Excel_Sheet1.xlsx',
      'chartStyle:ppt/charts/style1.xml',
      'themeOverride:ppt/theme/themeOverride1.xml',
    ]);
    const store = new Map(r.opaqueParts.map((p) => [p.hash, p]));
    expect(store.get(chart.parts[0].hash)?.contentType).toBe('application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
    expect(store.get(chart.parts[1].hash)?.contentType).toBe('application/vnd.openxmlformats-package.relationships+xml');
    expect(store.get(chart.parts[2].hash)).toMatchObject({ contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', bytes: WORKBOOK });
    expect(r.warnings).toEqual([]);
  });

  it('stores equal bytes once however many charts hold them, and each chart still lists its parts', async () => {
    const extra = { ...chartParts(1), ...chartParts(2, CHART_PLAIN), 'ppt/embeddings/Microsoft_Excel_Sheet2.xlsx': WORKBOOK };
    const r = await read({
      tree: chartFrame('3', 'rId2') + chartFrame('4', 'rId3'),
      slideRels: [rel('rId2', 'chart', '../charts/chart1.xml'), rel('rId3', 'chart', '../charts/chart2.xml')],
      extra,
      types,
    });
    const [a, b] = opaqueOf(r);
    expect(a.title).toBe('Sales 2024');
    expect(b.title).toBeUndefined();
    const workbook = r.opaqueParts.filter((p) => p.contentType.endsWith('.sheet'));
    expect(workbook).toHaveLength(1);
    expect(a.parts[2].hash).toBe(b.parts[2].hash);
    // the shared theme override part is read once (one entry for both)
    expect(r.opaqueParts.filter((p) => p.bytes.length === '<a:themeOverride/>'.length)).toHaveLength(1);
  });

  it('keeps a chart whose relationship is missing as a placeholder with no parts, and warns about a part that is missing', async () => {
    const missing = await read({ tree: chartFrame('3', 'rId9') });
    expect(opaqueOf(missing)[0].parts).toEqual([]);
    const dangling = await read({ tree: chartFrame('3', 'rId2'), slideRels: [rel('rId2', 'chart', '../charts/chart1.xml')] });
    expect(opaqueOf(dangling)[0].parts).toEqual([]);
    expect(dangling.warnings).toContain('ppt/charts/chart1.xml could not be read (it is missing or too large to read) and is not kept.');
  });

  it('keeps a hidden chart hidden and one without a description plain', async () => {
    const [c] = opaqueOf(await read({ tree: chartFrame('3', 'rId9', ' hidden="1"') }));
    expect(c.hidden).toBe(true);
    expect(c.altText).toBeUndefined();
  });

  it('does not take a document part for a chart', async () => {
    const r = await read({ tree: chartFrame('3', 'rId2'), slideRels: [rel('rId2', 'slide', 'slide1.xml')] });
    expect(opaqueOf(r)[0].parts).toEqual([]);
  });

  it('keeps a modern chart (chartex) under a mc:AlternateContent choice', async () => {
    const cx = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
    const tree = `<mc:AlternateContent><mc:Choice Requires="cx1">${frame('5', 'Funnel', cx, `<cx:chart xmlns:cx="${cx}" r:id="rId2"/>`)}</mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="5" name="Fallback"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></mc:Fallback></mc:AlternateContent>`;
    const r = await read({ tree, slideRels: [rel('rId2', 'chartEx', '../charts/chartEx1.xml')], extra: { 'ppt/charts/chartEx1.xml': '<cx:chartSpace/>' } });
    expect(opaqueOf(r)[0]).toMatchObject({ opaqueKind: 'chart', pptxId: '5', parts: [{ role: 'chartEx', partName: 'ppt/charts/chartEx1.xml' }] });
  });

  it('draws an element inside a group where the group puts it', async () => {
    const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="1905000" y="0"/><a:ext cx="4762500" cy="2857500"/><a:chOff x="0" y="0"/><a:chExt cx="4762500" cy="2857500"/></a:xfrm></p:grpSpPr>${chartFrame('3', 'rId2')}</p:grpSp>`;
    const r = await read({ tree: group, slideRels: [rel('rId2', 'chart', '../charts/chart1.xml')], extra: chartParts(1), types });
    const g = (r.slides[0].elements as PptxElement[])[0];
    expect(g.kind === 'group' && g.children[0]).toMatchObject({ kind: 'opaque', frame: { x: 300, y: 200 } });
  });
});

describe('SmartArt', () => {
  const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
  const data = (ext: string): string => `<dgm:dataModel xmlns:dgm="${DGM}"><dgm:extLst>${ext}</dgm:extLst></dgm:dataModel>`;
  const ids = (dm: string): string => frame('3', 'Diagram', DGM, `<dgm:relIds xmlns:dgm="${DGM}" r:dm="${dm}" r:lo="rId3" r:qs="rId4" r:cs="rId5"/>`);
  const slideRels = [rel('rId1', 'diagramDrawing', '../diagrams/drawing1.xml'), rel('rId2', 'diagramData', '../diagrams/data1.xml'), rel('rId3', 'diagramLayout', '../diagrams/layout1.xml'), rel('rId4', 'diagramQuickStyle', '../diagrams/quickStyle1.xml'), rel('rId5', 'diagramColors', '../diagrams/colors1.xml')];
  const others = { 'ppt/diagrams/layout1.xml': '<l/>', 'ppt/diagrams/quickStyle1.xml': '<q/>', 'ppt/diagrams/colors1.xml': '<c/>', 'ppt/diagrams/drawing1.xml': '<d/>' };

  it('keeps the four diagram parts and the cached drawing the data part names', async () => {
    const r = await read({ tree: ids('rId2'), slideRels, extra: { ...others, 'ppt/diagrams/data1.xml': data('<a:ext xmlns:a="x"><dsp:dataModelExt xmlns:dsp="y" relId="rId1" minVer="z"/></a:ext>') } });
    const [d] = opaqueOf(r);
    expect(d.opaqueKind).toBe('smartArt');
    expect(roles(d).filter((x) => !x.startsWith('rels'))).toEqual(['diagramData:ppt/diagrams/data1.xml', 'diagramLayout:ppt/diagrams/layout1.xml', 'diagramQuickStyle:ppt/diagrams/quickStyle1.xml', 'diagramColors:ppt/diagrams/colors1.xml', 'diagramDrawing:ppt/diagrams/drawing1.xml']);
  });

  it('keeps a diagram frame with no relationship ids as an empty placeholder', async () => {
    const [d] = opaqueOf(await read({ tree: frame('3', 'Diagram', DGM, '') }));
    expect(d).toMatchObject({ opaqueKind: 'smartArt', parts: [] });
  });

  it('keeps what it can find when the data part names no drawing or a relationship is missing', async () => {
    const r = await read({ tree: ids('rId8'), slideRels, extra: { ...others, 'ppt/diagrams/data1.xml': data('') } });
    expect(roles(opaqueOf(r)[0])).toEqual(['diagramLayout:ppt/diagrams/layout1.xml', 'diagramQuickStyle:ppt/diagrams/quickStyle1.xml', 'diagramColors:ppt/diagrams/colors1.xml']);
    const noData = await read({ tree: ids('rId2'), slideRels, extra: { ...others, 'ppt/diagrams/data1.xml': data('') } });
    expect(roles(opaqueOf(noData)[0])[0]).toBe('diagramData:ppt/diagrams/data1.xml');
  });
});

describe('OLE objects and controls', () => {
  const OLE = 'http://schemas.openxmlformats.org/presentationml/2006/ole';
  const ole = (rid: string, preview = 'rId1'): string => frame('3', 'Embedded', OLE, `<p:oleObj name="Worksheet" r:id="${rid}" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${preview}"/></p:blipFill><p:spPr/></p:pic></p:oleObj>`);

  it('keeps the object and its preview and shows the preview', async () => {
    const r = await read({ tree: ole('rId2'), slideRels: [rel('rId1', 'image', '../media/preview.png'), rel('rId2', 'oleObject', '../embeddings/oleObject1.bin')], extra: { 'ppt/media/preview.png': PNG, 'ppt/embeddings/oleObject1.bin': Uint8Array.from([1, 2, 3]) } });
    const [o] = opaqueOf(r);
    expect(o).toMatchObject({ opaqueKind: 'ole', props: { progId: 'Excel.Sheet.12', oleName: 'Worksheet' } });
    expect(o.poster).toMatch(/^data:image\/png;base64,/);
    expect(roles(o)).toEqual(['oleObject:ppt/embeddings/oleObject1.bin', 'image:ppt/media/preview.png']);
  });

  it('keeps a linked object as an address that is never fetched, and has no poster when the preview is not a picture a browser shows', async () => {
    const r = await read({ tree: ole('rId2'), slideRels: [rel('rId1', 'image', '../media/preview.emf'), rel('rId2', 'oleObject', 'file:///C:/sheet.xlsx', true)], extra: { 'ppt/media/preview.emf': Uint8Array.from([1]) } });
    const [o] = opaqueOf(r);
    expect(o.props).toMatchObject({ linkedTarget: 'file:///C:/sheet.xlsx' });
    expect(o.poster).toBeUndefined();
    expect(roles(o)).toEqual(['image:ppt/media/preview.emf']);
  });

  it('keeps an object with no relationship, no preview and no name', async () => {
    const bare = frame('3', 'Bare', OLE, '<p:oleObj/>');
    const [o] = opaqueOf(await read({ tree: bare }));
    expect(o).toMatchObject({ opaqueKind: 'ole', parts: [] });
    expect(o.props).toBeUndefined();
  });

  it('keeps an ActiveX control from p:controls with its part and picture', async () => {
    const controls = '<p:controls><p:control spid="7" name="Button" r:id="rId3"><p:pic><p:nvPicPr><p:cNvPr id="0" name="ctl" descr="A button" hidden="1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/></p:blipFill><p:spPr><a:xfrm><a:off x="952500" y="952500"/><a:ext cx="952500" cy="476250"/></a:xfrm></p:spPr></p:pic></p:control><p:control r:id="rId9"><p:pic><p:nvPicPr><p:cNvPr id="8" name="named by pic"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic></p:control></p:controls>';
    const r = await read({
      controls,
      slideRels: [rel('rId1', 'image', '../media/preview.png'), rel('rId3', 'control', '../activeX/activeX1.xml')],
      extra: { 'ppt/media/preview.png': PNG, 'ppt/activeX/activeX1.xml': '<ax/>', 'ppt/activeX/_rels/activeX1.xml.rels': rels([rel('rId1', 'activeXControlBinary', 'activeX1.bin')]), 'ppt/activeX/activeX1.bin': Uint8Array.from([5]) },
    });
    const [a, b] = opaqueOf(r);
    expect(a).toMatchObject({ opaqueKind: 'ole', pptxId: '7', name: 'Button', altText: 'A button', hidden: true, props: { control: 'true' }, frame: { x: 100, y: 100, w: 100, h: 50 } });
    expect(roles(a)).toEqual(['control:ppt/activeX/activeX1.xml', 'rels:ppt/activeX/_rels/activeX1.xml.rels', 'image:ppt/media/preview.png', 'activeXControlBinary:ppt/activeX/activeX1.bin']);
    expect(a.poster).toBeDefined();
    expect(b).toMatchObject({ pptxId: '8', name: 'named by pic', parts: [] });
    expect(b.frame).toBeUndefined();
  });
});

describe('audio and video', () => {
  const clip = (kind: string, link: string, embed?: string, attrs = ' descr="A clip" hidden="1"'): string =>
    `<p:pic><p:nvPicPr><p:cNvPr id="3" name="Clip"${attrs}/><p:cNvPicPr/><p:nvPr><a:${kind}${link ? ` r:link="${link}"` : ''}/>${embed ? `<p:extLst><p:ext uri="{x}"><p14:media r:embed="${embed}"/></p:ext></p:extLst>` : ''}</p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm></p:spPr></p:pic>`;
  const media = { 'ppt/media/poster.png': PNG, 'ppt/media/clip.mp4': Uint8Array.from([0, 0, 0, 24]) };

  it('keeps the clip under both relationships, once, with the poster frame as the picture', async () => {
    const r = await read({ tree: clip('videoFile', 'rId2', 'rId3'), slideRels: [rel('rId1', 'image', '../media/poster.png'), rel('rId2', 'video', '../media/clip.mp4'), rel('rId3', 'media', '../media/clip.mp4')], extra: media });
    const [m] = opaqueOf(r);
    expect(m).toMatchObject({ opaqueKind: 'media', pptxId: '3', name: 'Clip', altText: 'A clip', hidden: true, props: { clip: 'videoFile' } });
    expect(roles(m)).toEqual(['video:ppt/media/clip.mp4', 'image:ppt/media/poster.png']);
    expect(m.poster).toMatch(/^data:image\/png/);
    expect(r.opaqueParts).toHaveLength(2);
  });

  it('keeps a linked clip as an address that is never fetched', async () => {
    const r = await read({ tree: clip('audioFile', 'rId2'), slideRels: [rel('rId1', 'image', '../media/poster.png'), rel('rId2', 'audio', 'https://example.com/a.mp3', true)], extra: media });
    const [m] = opaqueOf(r);
    expect(m.props).toEqual({ clip: 'audioFile', linkedTarget: 'https://example.com/a.mp3' });
    expect(roles(m)).toEqual(['image:ppt/media/poster.png']);
  });

  it('keeps a clip given only by the newer media relationship, plain and visible', async () => {
    const r = await read({ tree: clip('videoFile', '', 'rId3', ''), slideRels: [rel('rId1', 'image', '../media/poster.png'), rel('rId3', 'media', '../media/clip.mp4')], extra: media });
    const [m] = opaqueOf(r);
    expect(m.altText).toBeUndefined();
    expect(m.hidden).toBeUndefined();
    expect(roles(m)).toEqual(['media:ppt/media/clip.mp4', 'image:ppt/media/poster.png']);
  });

  it('is still a clip when its poster is missing', async () => {
    const [m] = opaqueOf(await read({ tree: clip('quickTimeFile', 'rId2'), slideRels: [rel('rId2', 'video', '../media/clip.mp4')], extra: media }));
    expect(m).toMatchObject({ opaqueKind: 'media', props: { clip: 'quickTimeFile' } });
    expect(m.poster).toBeUndefined();
    expect(roles(m)).toEqual(['video:ppt/media/clip.mp4']);
  });
});

describe('ink and equations', () => {
  const ink = (rid: string, nv = '<p14:nvContentPartPr><p14:cNvPr id="3" name="Ink 2"/></p14:nvContentPartPr>'): string =>
    `<mc:AlternateContent><mc:Choice Requires="p14"><p:contentPart r:id="${rid}">${nv}<p14:xfrm><a:off x="952500" y="0"/><a:ext cx="952500" cy="476250"/></p14:xfrm></p:contentPart></mc:Choice></mc:AlternateContent>`;

  it('keeps ink with the part its relationship names', async () => {
    const r = await read({ tree: ink('rId2'), slideRels: [rel('rId2', 'customXml', '../ink/ink1.xml')], extra: { 'ppt/ink/ink1.xml': '<ink/>' } });
    const [i] = opaqueOf(r);
    expect(i).toMatchObject({ opaqueKind: 'ink', pptxId: '3', name: 'Ink 2', frame: { x: 100, y: 0, w: 100, h: 50 } });
    expect(roles(i)).toEqual(['customXml:ppt/ink/ink1.xml']);
  });

  it('keeps ink with no part and no name', async () => {
    const [i] = opaqueOf(await read({ tree: ink('rId9', '') }));
    expect(i).toMatchObject({ opaqueKind: 'ink', pptxId: '', name: '', parts: [] });
  });

  it('reads the fallback text of an equation, whose a14 choice carries only math', async () => {
    const sp = (text: string): string => `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Equation"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
    const tree = `<mc:AlternateContent><mc:Choice Requires="a14">${sp('')}</mc:Choice><mc:Fallback>${sp('x = 1')}</mc:Fallback></mc:AlternateContent><mc:AlternateContent><mc:Choice Requires="a14">${sp('only choice')}</mc:Choice></mc:AlternateContent><mc:AlternateContent><mc:Choice>${sp('plain choice')}</mc:Choice><mc:Fallback>${sp('not this')}</mc:Fallback></mc:AlternateContent>`;
    const els = (await read({ tree })).slides[0].elements as PptxElement[];
    const text = els.map((e) => (e.kind === 'shape' ? e.text?.paragraphs[0]?.runs[0]?.text : undefined));
    expect(text).toEqual(['x = 1', 'only choice', 'plain choice']);
  });
});

describe('limits', () => {
  const chart = (extra: Record<string, string | Uint8Array>): Deck => ({ tree: chartFrame('3', 'rId2'), slideRels: [rel('rId2', 'chart', '../charts/chart1.xml')], extra });

  it('drops a part over the per-part cap with a warning and keeps the rest', async () => {
    const r = await read(chart({ ...chartParts(1, CHART_PLAIN), 'ppt/embeddings/Microsoft_Excel_Sheet1.xlsx': new Uint8Array(1200) }), { opaqueLimits: { maxPartBytes: 1000 } });
    const [c] = opaqueOf(r);
    expect(roles(c)).toEqual(['chart:ppt/charts/chart1.xml', 'rels:ppt/charts/_rels/chart1.xml.rels', 'chartStyle:ppt/charts/style1.xml', 'themeOverride:ppt/theme/themeOverride1.xml']);
    expect(r.warnings).toContain('ppt/embeddings/Microsoft_Excel_Sheet1.xlsx is over the size limit for a kept part and is dropped.');
  });

  it('drops a chart whose own part is over the cap, with what hangs on it', async () => {
    const r = await read(chart(chartParts(1)), { opaqueLimits: { maxPartBytes: 20 } });
    expect(opaqueOf(r)[0].parts).toEqual([]);
    expect(r.warnings).toContain('ppt/charts/chart1.xml is over the size limit for a kept part and is dropped.');
  });

  it('drops a part the store has no room for', async () => {
    const r = await read(chart(chartParts(1)), { opaqueLimits: { maxTotalBytes: 1000 } });
    expect(r.warnings.filter((w) => w.includes('over the size limit'))).toHaveLength(0);
    const tight = await read(chart(chartParts(1)), { opaqueLimits: { maxTotalBytes: 100 } });
    expect(tight.warnings.some((w) => w.includes('ppt/charts/chart1.xml is not kept: the deck’s kept parts are over the size limit and is dropped.'))).toBe(true);
  });

  it('follows relationships a bounded number of steps and parts, and not in circles', async () => {
    const chain: Record<string, string> = { 'ppt/charts/chart1.xml': '<c/>' };
    for (let i = 0; i < 12; i++) {
      const from = i === 0 ? 'chart1' : `p${i}`;
      chain[`ppt/charts/_rels/${from}.xml.rels`] = rels([rel('rId1', 'x', `p${i + 1}.xml`), rel('rId2', 'x', `${from}.xml`)]);
      chain[`ppt/charts/p${i + 1}.xml`] = `<p n="${i}"/>`;
    }
    const deep = opaqueOf(await read(chart(chain)))[0];
    expect(deep.parts.filter((p) => p.role === 'x')).toHaveLength(8);

    const wide: Record<string, string> = { 'ppt/charts/chart1.xml': '<c/>', 'ppt/charts/_rels/chart1.xml.rels': rels(Array.from({ length: 300 }, (_, i) => rel(`r${i}`, 'x', `w${i}.xml`))) };
    for (let i = 0; i < 300; i++) wide[`ppt/charts/w${i}.xml`] = `<w n="${i}"/>`;
    const many = opaqueOf(await read(chart(wide)))[0];
    expect(many.parts.filter((p) => p.role === 'x')).toHaveLength(255);
  });

  it('does not follow relationships to the document\'s own structure, external targets or missing parts', async () => {
    const extra = { 'ppt/charts/chart1.xml': '<c/>', 'ppt/charts/_rels/chart1.xml.rels': rels([rel('r1', 'slide', '../slides/slide1.xml'), rel('r2', 'hyperlink', 'https://example.com', true), rel('r3', 'package', '../embeddings/nothing.xlsx')]) };
    expect(roles(opaqueOf(await read(chart(extra)))[0])).toEqual(['chart:ppt/charts/chart1.xml', 'rels:ppt/charts/_rels/chart1.xml.rels']);
  });
});

describe('transition and timing', () => {
  const timing = '<p:timing><p:tnLst><p:par><p:cTn id="1"><p:childTnLst><p:set><p:cBhvr><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cBhvr></p:set></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>';
  const trans = '<p:transition spd="med"><p:fade/></p:transition>';

  it('keeps the transition and the timing tree as written, with the namespaces they use', async () => {
    const r = await read({ after: `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${trans}${timing}` });
    const { transition, timing: t } = r.slides[0].opaque ?? {};
    expect(transition).toBe(trans.replace('<p:transition', `<p:transition ${NS.replace(/ xmlns/g, ' xmlns').trim().split(' ').map((d) => ` ${d}`).join('').trim()}`).replace(/<p:transition +/, '<p:transition '));
    expect(t?.startsWith('<p:timing xmlns:a=')).toBe(true);
    expect(t).toContain('<p:spTgt spid="3"/>');
    expect(parseXml(t as string)?.name).toBe('timing');
  });

  it('keeps a newer transition together with its fallback', async () => {
    const wrapped = '<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="2000"><p14:vortex dir="r"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>';
    const r = await read({ after: wrapped });
    expect(r.slides[0].opaque?.transition).toContain('<p14:vortex dir="r"/>');
    expect(r.slides[0].opaque?.transition).toContain('<p:fade/>');
    expect(r.slides[0].opaque?.timing).toBeUndefined();
  });

  it('has no opaque record for a slide with neither, or with other alternate content', async () => {
    expect((await read({})).slides[0].opaque).toBeUndefined();
    expect((await read({ after: '<mc:AlternateContent><mc:Choice Requires="p14"><p:extLst/></mc:Choice></mc:AlternateContent>' })).slides[0].opaque).toBeUndefined();
  });

  it('degrades a slide whose content cannot be read, without an opaque record', async () => {
    const r = await read({ tree: '<p:sp/>', after: trans });
    expect(r.slides[0].opaque).toBeDefined();
  });
});

describe('slideFragments', () => {
  it('finds the fragments among comments, CDATA, processing instructions, text and quoted > characters', () => {
    const xml = `<?xml version="1.0"?><!DOCTYPE x><!-- <p:transition/> --><p:sld xmlns:p='urn:p' xmlns:a="urn:a" note="a>b"><p:cSld><![CDATA[<p:timing/>]]></p:cSld><?pi <p:timing/> ?>text<p:transition  spd="fast" note='>' xmlns:a="urn:own"><p:fade/></p:transition><p:timing/><p:timing x="second"/><p:transition x="second"/></p:sld>`;
    const f = slideFragments(xml);
    expect(f.transition).toBe(`<p:transition xmlns:p='urn:p'  spd="fast" note='>' xmlns:a="urn:own"><p:fade/></p:transition>`.replace(`xmlns:p='urn:p'  spd`, `xmlns:p='urn:p'  spd`));
    expect(f.timing).toBe(`<p:timing xmlns:p='urn:p' xmlns:a="urn:a"/>`);
  });

  it('does not repeat a declaration the fragment already has, whatever the prefix (dots included)', () => {
    const f = slideFragments('<s xmlns="urn:d" xmlns:p.q="urn:q"><p.q:timing xmlns:p.q="urn:q" xmlns="urn:own"/></s>');
    expect(f.timing).toBe('<p.q:timing xmlns:p.q="urn:q" xmlns="urn:own"/>');
  });

  it('finds a fragment in a slide whose root declares no namespace', () => {
    expect(slideFragments('<sld><timing/></sld>')).toEqual({ timing: '<timing/>' });
  });

  it('finds nothing in a slide with no children', () => {
    expect(slideFragments('<p:sld xmlns:p="urn:p"/>')).toEqual({});
  });
});

describe('content types', () => {
  it('prefers an override, then the extension default, then a generic type', () => {
    const t = readContentTypes(parseXml(`<Types xmlns="${CT}"><Default Extension="PNG" ContentType="image/png"/><Override PartName="/ppt/Charts/A.png" ContentType="chart/x"/><Default Extension="" ContentType=""/><Default/><Override ContentType="y"/></Types>`) ?? undefined);
    expect(t.of('ppt/charts/a.png')).toBe('chart/x');
    expect(t.of('x/b.png')).toBe('image/png');
    expect(t.of('x/c.bin')).toBe('application/octet-stream');
    expect(readContentTypes(undefined).of('a.xml')).toBe('application/octet-stream');
  });
});

describe('loss report', () => {
  const chartDeck = (): Deck => ({ tree: chartFrame('3', 'rId2'), slideRels: [rel('rId2', 'chart', '../charts/chart1.xml')], extra: chartParts(1) });

  it('reports what is kept as kept, and a chart no slide shows as not imported', async () => {
    const report = await detectPptxLoss(await build({ ...chartDeck(), after: '<p:transition><p:fade/></p:transition>' }), { detail: 'full' });
    const byId = new Map(report.items.map((i) => [i.id, i]));
    expect(byId.get('pptx.charts')).toMatchObject({ state: 'kept', count: 1, where: ['Slide 1'] });
    expect(byId.get('pptx.transitions')).toMatchObject({ state: 'kept' });
    const orphan = await detectPptxLoss(await build({ extra: chartParts(1) }));
    expect(orphan.items.find((i) => i.id === 'pptx.charts')).toMatchObject({ count: 1, effect: expect.stringContaining('not imported') });
    expect(orphan.items.find((i) => i.id === 'pptx.charts')?.state).toBeUndefined();
  });

  it('reports ink as kept when a slide shows it and as not imported when none does', async () => {
    const tree = '<mc:AlternateContent><mc:Choice Requires="p14"><p:contentPart r:id="rId2"/></mc:Choice></mc:AlternateContent>';
    const kept = await detectPptxLoss(await build({ tree, slideRels: [rel('rId2', 'customXml', '../ink/ink1.xml')], extra: { 'ppt/ink/ink1.xml': '<ink/>' } }), { detail: 'full' });
    expect(kept.items.find((i) => i.id === 'pptx.ink')?.state).toBe('kept');
    // ink is a finer id: the dialog's summary level leaves it out
    expect((await detectPptxLoss(await build({ tree, slideRels: [rel('rId2', 'customXml', '../ink/ink1.xml')], extra: { 'ppt/ink/ink1.xml': '<ink/>' } }))).items.find((i) => i.id === 'pptx.ink')).toBeUndefined();
    const orphan = await detectPptxLoss(await build({ extra: { 'ppt/ink/ink1.xml': '<ink/>' } }), { detail: 'full' });
    expect(orphan.items.find((i) => i.id === 'pptx.ink')?.state).toBeUndefined();
  });

  it('reports a part over the size cap as lost', async () => {
    const big = new Uint8Array(26 * 1024 * 1024);
    const report = await detectPptxLoss(await build({ extra: { 'ppt/media/big.mp4': big } }));
    expect(report.items.find((i) => i.id === 'pptx.opaque.tooLarge')).toMatchObject({ count: 1, where: ['ppt/media/big.mp4'], state: 'lost' });
  });
});
