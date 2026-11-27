import { readFileSync } from 'fs';
import { join } from 'path';
import JSZip from 'jszip';
import type { ImportLossItem, ImportLossReport } from '@b32nio/spindle-shared';
import { PACKAGE_LIMITS, detectPptxLoss } from './loss';

// ---------------------------------------------------------------------------
// Minimal synthetic decks, one construct at a time. Each package is a real zip built here, so the detector runs
// through the same code path as for a file from disk.
// ---------------------------------------------------------------------------



const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

const rel = (id: string, type: string, target: string, external = false): string =>
  `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;
const relsXml = (rels: string[]): string => `<Relationships xmlns="${RELS_NS}">${rels.join('')}</Relationships>`;

interface SlideSpec {
  /** Shape elements placed inside `p:spTree`. */
  shapes?: string;
  /** Elements after `p:cSld` (transition, timing, clrMapOvr). */
  tail?: string;
  /** Extra attributes on `p:sld` (`show="0"`). */
  attrs?: string;
  /** Relationship elements of this slide. */
  rels?: string[];
  /** The whole slide part, replacing the generated one. */
  raw?: string;
}

interface Spec {
  slides?: SlideSpec[];
  /** Children of `p:presentation` after `p:sldIdLst` (sldSz, extLst). */
  presentation?: string;
  /** The whole presentation part, replacing the generated one (the slide relationships are still written). */
  presentationRaw?: string;
  /** Extra files (name -> text or bytes). */
  files?: Record<string, string | Uint8Array>;
  /** Omit `ppt/presentation.xml` altogether. */
  noPresentation?: boolean;
}

const slideXml = (s: SlideSpec = {}): string =>
  s.raw ??
  `<p:sld ${NS}${s.attrs ? ` ${s.attrs}` : ''}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${s.shapes ?? ''}</p:spTree></p:cSld>${s.tail ?? ''}</p:sld>`;

async function build(spec: Spec = {}): Promise<Uint8Array> {
  const slides = spec.slides ?? [{}];
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  const presRels: string[] = [];
  const ids: string[] = [];
  slides.forEach((s, i) => {
    const n = i + 1;
    presRels.push(rel(`rId${n}`, 'slide', `slides/slide${n}.xml`));
    ids.push(`<p:sldId id="${256 + i}" r:id="rId${n}"/>`);
    zip.file(`ppt/slides/slide${n}.xml`, slideXml(s));
    if (s.rels) zip.file(`ppt/slides/_rels/slide${n}.xml.rels`, relsXml(s.rels));
  });
  zip.file('ppt/_rels/presentation.xml.rels', relsXml(presRels));
  if (!spec.noPresentation)
    zip.file(
      'ppt/presentation.xml',
      spec.presentationRaw ??
        `<p:presentation ${NS}><p:sldIdLst>${ids.join('')}</p:sldIdLst>${spec.presentation ?? '<p:sldSz cx="12192000" cy="6858000"/>'}</p:presentation>`
    );
  for (const [name, content] of Object.entries(spec.files ?? {})) zip.file(name, content);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const full = (bytes: Uint8Array) => detectPptxLoss(bytes, { detail: 'full' });
const ids = (r: ImportLossReport): string[] => r.items.map((i) => i.id).sort();
const item = (r: ImportLossReport, id: string): ImportLossItem | undefined => r.items.find((i) => i.id === id);

// ---- shape builders ----------------------------------------------------------

const frame = '<a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>';

/** A `p:sp`. `spPr` is the content of `p:spPr`; `body` the content of `p:txBody` (after bodyPr); `bodyPr` its attributes/children. */
function sp(o: { spPr?: string; body?: string; bodyPr?: string; ph?: string; id?: number } = {}): string {
  const text = o.body !== undefined || o.bodyPr !== undefined;
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${o.id ?? 2}" name="s"/><p:cNvSpPr/><p:nvPr>${o.ph ?? ''}</p:nvPr></p:nvSpPr>` +
    `<p:spPr>${o.spPr ?? frame}</p:spPr>` +
    (text ? `<p:txBody>${o.bodyPr ?? '<a:bodyPr/>'}<a:lstStyle/>${o.body ?? '<a:p/>'}</p:txBody>` : '') +
    '</p:sp>'
  );
}

const para = (pPr = '', runs = '<a:r><a:t>x</a:t></a:r>'): string => `<a:p>${pPr}${runs}</a:p>`;
const geom = (prst: string): string => `${frame}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>`;
const run = (rPr: string, text = 'x'): string => `<a:r><a:rPr ${rPr}/><a:t>${text}</a:t></a:r>`;

const gf = (uri: string, inner = ''): string =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="g"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm><a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic></p:graphicFrame>`;
const CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const DIAGRAM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const pic = (blipFill = '<a:blip r:embed="rId1"/>', nvPr = ''): string =>
  `<p:pic><p:nvPicPr><p:cNvPr id="4" name="p"/><p:cNvPicPr/><p:nvPr>${nvPr}</p:nvPr></p:nvPicPr><p:blipFill>${blipFill}<a:stretch/></p:blipFill><p:spPr>${frame}</p:spPr></p:pic>`;
const one = (shapes: string, extra: Partial<SlideSpec> = {}): Promise<ImportLossReport> => build({ slides: [{ shapes, ...extra }] }).then(full);
const text = (inner: string, bodyPr = '<a:bodyPr/>'): string => sp({ bodyPr, body: inner });

describe('plain content', () => {
  it('reports nothing for a deck of plain text', async () => {
    const bytes = await build({
      slides: [{ shapes: text(para('', run('lang="en-US" b="1"'))) }, { shapes: sp({ spPr: geom('rect'), body: para() }) }],
    });
    expect(await detectPptxLoss(bytes)).toEqual({ format: 'pptx', items: [] });
    expect(await full(bytes)).toEqual({ format: 'pptx', items: [] });
  });

  it('reports nothing for bytes that are not a package, and never throws', async () => {
    for (const bytes of [new Uint8Array(0), new Uint8Array([1, 2, 3]), new TextEncoder().encode('not a zip at all'), new Uint8Array(4096).fill(0x50)])
      expect(await full(bytes)).toEqual({ format: 'pptx', items: [] });
  });

  it('reports nothing for a zip that is not a presentation, and does not call its parts unaccounted', async () => {
    const zip = new JSZip();
    zip.file('hello.txt', 'hi');
    expect(await full(await zip.generateAsync({ type: 'uint8array' }))).toEqual({ format: 'pptx', items: [] });
  });
});

describe('summary ids', () => {
  it('counts charts as objects, not parts, and names the slides', async () => {
    const r = await build({
      slides: [{ shapes: gf(CHART) + gf(CHART) }, { shapes: gf(CHART) }, { shapes: gf('http://schemas.microsoft.com/office/drawing/2014/chartex') }],
      files: { 'ppt/charts/chart1.xml': '<x/>', 'ppt/charts/style1.xml': '<x/>', 'ppt/charts/colors1.xml': '<x/>', 'ppt/embeddings/Microsoft_Excel_Sheet1.xlsx': 'x' },
    }).then(full);
    expect(item(r, 'pptx.charts')).toMatchObject({ count: 4, where: ['Slide 1', 'Slide 2', 'Slide 3'] });
  });

  it('falls back to the chart parts when no slide shows a chart frame', async () => {
    const r = await build({ files: { 'ppt/charts/chart1.xml': '<x/>', 'ppt/charts/chart2.xml': '<x/>', 'ppt/charts/style1.xml': '<x/>' } }).then(full);
    expect(item(r, 'pptx.charts')).toEqual(expect.objectContaining({ count: 2 }));
    expect(item(r, 'pptx.charts')?.where).toBeUndefined();
  });

  it('counts SmartArt diagrams, falling back to the data parts', async () => {
    expect(item(await one(gf(DIAGRAM) + gf(DIAGRAM)), 'pptx.smartArt')).toMatchObject({ count: 2, where: ['Slide 1'] });
    const parts = await build({ files: { 'ppt/diagrams/data1.xml': '<x/>', 'ppt/diagrams/layout1.xml': '<x/>', 'ppt/diagrams/drawing1.xml': '<x/>' } }).then(full);
    expect(item(parts, 'pptx.smartArt')?.count).toBe(1);
  });

  it('ignores graphic frames that are neither charts nor diagrams (tables, a missing uri)', async () => {
    const r = await one(gf('http://schemas.openxmlformats.org/drawingml/2006/table') + `<p:graphicFrame><a:graphic><a:graphicData/></a:graphic></p:graphicFrame>`);
    expect(ids(r)).toEqual([]);
  });

  it('counts audio and video clips (a:videoFile, a:audioFile, a:quickTimeFile), one per picture', async () => {
    const r = await one(
      pic('<a:blip r:embed="rId1"/>', '<a:videoFile r:link="rId2"/><p:extLst><p:ext uri="x"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId3"/></p:ext></p:extLst>') +
        pic('<a:blip r:embed="rId1"/>', '<a:audioFile r:link="rId4"/>') +
        pic('<a:blip r:embed="rId1"/>', '<a:quickTimeFile r:link="rId5"/>') +
        pic()
    );
    expect(item(r, 'pptx.media')).toMatchObject({ count: 3, where: ['Slide 1'] });
  });

  it('counts embedded OLE objects and controls once, not again for the fallback picture', async () => {
    const ole = `<mc:AlternateContent><mc:Choice Requires="v"><p:oleObj name="x" r:id="rId2" progId="Excel.Sheet.12"/></mc:Choice><mc:Fallback><p:oleObj name="x" r:id="rId2"/></mc:Fallback></mc:AlternateContent>`;
    const r = await one(gf('http://schemas.openxmlformats.org/presentationml/2006/ole', ole) + `<p:control r:id="rId9" name="Button"/>`);
    expect(item(r, 'pptx.embeddedObjects')).toMatchObject({ count: 2, where: ['Slide 1'] });
  });

  it('counts slides with a transition once, although a PowerPoint 2010 transition is written twice', async () => {
    const t10 = '<mc:AlternateContent><mc:Choice Requires="p14"><p:transition spd="slow"><p14:vortex xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>';
    const r = await build({ slides: [{ tail: t10 }, {}, { tail: '<p:transition advClick="0" advTm="3000"/>' }] }).then(full);
    expect(item(r, 'pptx.transitions')).toMatchObject({ count: 2, where: ['Slide 1', 'Slide 3'] });
  });

  it('counts animation effects, or one tree for a slide whose timing has no preset effect', async () => {
    const cTn = (cls: string, id: number) => `<p:par><p:cTn id="${id}" presetID="1" presetClass="${cls}" fill="hold"/></p:par>`;
    const timing = (inner: string) => `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">${inner}</p:cTn></p:par></p:tnLst></p:timing>`;
    const r = await build({ slides: [{ tail: timing(cTn('entr', 2) + cTn('exit', 3) + cTn('emph', 4)) }, { tail: timing('<p:set/>') }, { tail: timing('<p:video><p:cMediaNode vol="80000"/></p:video>') }, {}] }).then(full);
    expect(item(r, 'pptx.animations')).toMatchObject({ count: 4, where: ['Slide 1', 'Slide 2'] });
  });

  it('does not count a timing tree that only plays media', async () => {
    const media = '<p:timing><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:video><p:cMediaNode/></p:video></p:cTn></p:par></p:tnLst></p:timing>';
    expect(item(await build({ slides: [{ tail: media }] }).then(full), 'pptx.animations')).toBeUndefined();
  });

  it('does not report comments: the reader keeps them as threads (K10c)', async () => {
    const legacy = `<p:cmLst ${NS}><p:cm authorId="0" dt="2024-01-01T00:00:00.000" idx="1"><p:pos x="1" y="1"/><p:text>a</p:text></p:cm></p:cmLst>`;
    const modern = `<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:cm id="{1}" authorId="{2}"><p188:reply id="{3}"/></p188:cm></p188:cmLst>`;
    const r = await build({
      slides: [{ rels: [rel('rId1', 'comments', '../comments/comment1.xml')] }, { rels: [rel('rId1', 'comments', '../comments/modernComment_1.xml')] }],
      files: { 'ppt/comments/comment1.xml': legacy, 'ppt/comments/modernComment_1.xml': modern, 'ppt/commentAuthors.xml': '<x/>', 'ppt/comments/comment7.xml': legacy },
    }).then(full);
    expect(item(r, 'pptx.comments')).toBeUndefined();
    expect(item(r, 'pptx.package.unaccountedParts')).toBeUndefined();
  });

  it('counts the macro project once', async () => {
    const r = await build({ files: { 'ppt/vbaProject.bin': 'x', 'ppt/vbaProjectSignature.bin': 'y' } }).then(full);
    expect(item(r, 'pptx.macros')?.count).toBe(1);
    expect(item(r, 'pptx.macros')?.effect).toMatch(/^lost: macros are never run or kept/);
  });

  it('counts hidden slides (show="0" or "false"), not slides with show="1"', async () => {
    const r = await build({ slides: [{ attrs: 'show="0"' }, { attrs: 'show="1"' }, { attrs: 'show="false"' }, {}] }).then(full);
    expect(item(r, 'pptx.hiddenSlides')).toMatchObject({ count: 2, where: ['Slide 1', 'Slide 3'] });
  });

  it('counts the slides whose notes carry formatting (bold, italic, underline, bullets, numbering), not plain notes', async () => {
    const notes = (para: string) => `<p:notes ${NS}><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${sp({ ph: '<p:ph type="body"/>', body: para })}</p:spTree></p:cSld></p:notes>`;
    const run = (rPr: string) => `<a:p><a:r>${rPr}<a:t>x</a:t></a:r></a:p>`;
    const cases = [
      run('<a:rPr b="1"/>'), // 1: bold
      run('<a:rPr i="true"/>'), // 2: italic
      run('<a:rPr u="sng"/>'), // 3: underline
      `<a:p><a:pPr><a:buChar char="-"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>`, // 4: bullet
      `<a:p><a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>x</a:t></a:r></a:p>`, // 5: numbering
      run('<a:rPr b="0" i="0" u="none" lang="en-US"/>'), // 6: explicitly plain
      run(''), // 7: no properties
    ];
    const r = await build({
      slides: cases.map((_, i) => ({ rels: [rel('rId1', 'notesSlide', `../notesSlides/notesSlide${i + 1}.xml`)] })),
      files: Object.fromEntries(cases.map((c, i) => [`ppt/notesSlides/notesSlide${i + 1}.xml`, notes(c)])),
    }).then(full);
    expect(item(r, 'pptx.notes')).toMatchObject({ count: 5, where: ['Slide 1', 'Slide 2', 'Slide 3', 'Slide 4', 'Slide 5'] });
    expect(item(r, 'pptx.notes')?.effect).toMatch(/^approximated:/);
  });

  it('finds formatted notes pages no slide points at, by file name', async () => {
    const r = await build({
      files: { 'ppt/notesSlides/notesSlide3.xml': `<p:notes ${NS}><p:cSld><p:spTree>${sp({ ph: '<p:ph type="body"/>', body: '<a:p><a:r><a:rPr b="1"/><a:t>x</a:t></a:r></a:p>' })}</p:spTree></p:cSld></p:notes>`, 'ppt/notesSlides/notesSlide4.xml': `<p:notes ${NS}/>` } }).then(full);
    expect(item(r, 'pptx.notes')).toMatchObject({ count: 1, where: ['notesSlide3.xml'] });
  });

  it('counts shapes whose geometry the editor lacks: custom geometry and presets outside its library', async () => {
    const supported = ['rect', 'roundRect', 'ellipse', 'triangle', 'rightTriangle', 'diamond', 'pentagon', 'hexagon', 'octagon', 'star5', 'arrowRight', 'arrowLeft', 'chevron', 'parallelogram', 'trapezoid', 'plus', 'heart', 'cloud', 'rtTriangle', 'rightArrow', 'flowChartProcess', 'wedgeRectCallout', 'lineInv', 'line', 'straightConnector1'];
    const r = await one(
      supported.map((p) => sp({ spPr: geom(p) })).join('') +
        sp({ spPr: geom('ribbon') }) +
        sp({ spPr: geom('ribbon') }) +
        sp({ spPr: geom('flowChartDocument') }) +
        sp({ spPr: `${frame}<a:custGeom><a:pathLst/></a:custGeom>` }) +
        sp({ spPr: frame })
    );
    expect(item(r, 'pptx.shapes.unsupportedGeometry')).toMatchObject({ count: 4, where: ['ribbon', 'flowChartDocument', 'custGeom'] });
    expect(item(r, 'pptx.shapes.custGeom')).toMatchObject({ count: 1, where: ['Slide 1'] });
  });

  it('counts a picture with custom geometry, not one cropped to a preset shape', async () => {
    const withGeom = (g: string) => pic().replace(`<p:spPr>${frame}</p:spPr>`, `<p:spPr>${frame}${g}</p:spPr>`);
    const r = await one(withGeom('<a:custGeom><a:pathLst/></a:custGeom>') + withGeom('<a:prstGeom prst="ellipse"/>') + withGeom('<a:prstGeom prst="ribbon"/>'));
    expect(item(r, 'pptx.shapes.unsupportedGeometry')).toMatchObject({ count: 1, where: ['custGeom'] });
    expect(item(r, 'pptx.shapes.custGeom')?.count).toBe(1);
  });

  it('does not count a connector as unsupported geometry (it is reported as elbow/curved at the full level)', async () => {
    const cxn = (prst: string) => `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="5" name="c"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${geom(prst)}</p:spPr></p:cxnSp>`;
    const r = await one(cxn('straightConnector1') + cxn('bentConnector3') + cxn('curvedConnector3') + cxn('bentConnector') + sp({ spPr: geom('bentConnector2') }));
    expect(ids(r)).toEqual(['pptx.shapes.connectors.elbowCurved']);
    expect(item(r, 'pptx.shapes.connectors.elbowCurved')?.count).toBe(4);
  });

  it('counts pictures in EMF and WMF formats (by the file the picture points at), not PNG or an unknown relationship', async () => {
    const r = await build({
      slides: [
        {
          shapes: pic('<a:blip r:embed="rId1"/>') + pic('<a:blip r:embed="rId2"/>') + pic('<a:blip r:embed="rId3"/>') + pic('<a:blip r:embed="rId9"/>') + pic('<a:blip r:link="rId4"/>') + pic('<a:blip/>'),
          rels: [rel('rId1', 'image', '../media/a.emf'), rel('rId2', 'image', '../media/b.WMF'), rel('rId3', 'image', '../media/c.png'), rel('rId4', 'image', 'file:///x/d.emz', true)],
        },
      ],
    }).then(full);
    expect(item(r, 'pptx.images.unsupportedFormat')).toMatchObject({ count: 3, where: ['Slide 1'] });
  });

  it('reports EMF in a slide background too', async () => {
    const bg = `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId1"/></a:blipFill></p:bgPr></p:bg>`;
    const r = await build({
      slides: [{ raw: `<p:sld ${NS}><p:cSld>${bg}<p:spTree/></p:cSld></p:sld>`, rels: [rel('rId1', 'image', '../media/a.emf')] }],
    }).then(full);
    expect(item(r, 'pptx.images.unsupportedFormat')?.count).toBe(1);
  });
});

describe('fonts', () => {
  const latin = (face: string) => `<a:latin typeface="${face}"/>`;
  const runWith = (rPr: string) => `<a:r><a:rPr lang="en-US">${rPr}</a:rPr><a:t>x</a:t></a:r>`;
  const theme = (major: string, minor: string, extra = '') =>
    `<a:theme ${NS}><a:themeElements><a:fontScheme name="t"><a:majorFont><a:latin typeface="${major}"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/></a:majorFont><a:minorFont><a:latin typeface="${minor}"/></a:minorFont></a:fontScheme>${extra}</a:themeElements></a:theme>`;

  it('lists fonts outside the web-safe and metric-compatible list, once each, sorted, from runs, layouts, masters and themes', async () => {
    const r = await build({
      slides: [{ shapes: text(para('', runWith(latin('Zapfino') + '<a:ea typeface="MS Gothic"/><a:cs typeface="Arial"/>') + runWith(latin('zapfino')) + runWith(latin('ARIAL') + latin('+mn-lt') + latin('')))) }],
      files: {
        'ppt/slideLayouts/slideLayout1.xml': `<p:sldLayout ${NS}><a:defRPr>${latin('Papyrus')}</a:defRPr></p:sldLayout>`,
        'ppt/slideMasters/slideMaster1.xml': `<p:sldMaster ${NS}><a:defRPr>${latin('Calibri')}${latin('Impact')}</a:defRPr></p:sldMaster>`,
        'ppt/slideMasters/_rels/slideMaster1.xml.rels': relsXml([rel('rId1', 'theme', '../theme/theme1.xml'), rel('rId2', 'theme', 'http://x/theme', true)]),
        'ppt/theme/theme1.xml': theme('Comic Sans MS', 'Georgia'),
        'ppt/theme/theme2.xml': theme('Unlinked Font', 'Arial'),
      },
    }).then(full);
    const fonts = item(r, 'pptx.fonts.unavailable');
    expect(fonts?.count).toBe(5);
    expect(fonts?.where).toEqual(['Comic Sans MS', 'Impact', 'MS Gothic', 'Papyrus', 'Zapfino']);
    expect(fonts?.effect).toMatch(/^approximated: font names are kept/);
  });

  it('ignores a font element with no typeface', async () => {
    const r = await one(text(para('', '<a:r><a:rPr><a:latin/><a:ea typeface="  "/></a:rPr><a:t>x</a:t></a:r>')));
    expect(item(r, 'pptx.fonts.unavailable')).toBeUndefined();
  });

  it('knows the list: web-safe families and the Office defaults with open substitutes are available', async () => {
    const { AVAILABLE_FONTS, isAvailableFont } = await import('./fonts');
    for (const f of ['Arial', 'Times New Roman', 'Courier New', 'Georgia', 'Verdana', 'Calibri', 'Cambria', 'Carlito', 'sans-serif'])
      expect(isAvailableFont(f)).toBe(true);
    expect(isAvailableFont('  calibri LIGHT ')).toBe(true);
    for (const f of ['Comic Sans MS', 'Segoe UI', 'MS Gothic', 'Papyrus', '']) expect(isAvailableFont(f)).toBe(false);
    expect([...AVAILABLE_FONTS].every((f) => f === f.toLowerCase())).toBe(true);
  });
});

describe('listing the slides', () => {
  it('names slides by their place in the presentation, not their file name', async () => {
    const zipped = await build({
      slides: [{ attrs: 'show="0"' }, { attrs: 'show="0"' }],
      // swap the order: the second part comes first in the list
      presentationRaw: `<p:presentation ${NS}><p:sldIdLst><p:sldId id="257" r:id="rId2"/><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`,
    });
    const r = await full(zipped);
    expect(item(r, 'pptx.hiddenSlides')).toMatchObject({ count: 2, where: ['Slide 1', 'Slide 2'] });
  });

  it('skips a slide whose relationship is missing, external or points at a missing part, and lists a slide only once', async () => {
    const r = await build({
      slides: [{ attrs: 'show="0"' }],
      presentationRaw:
        `<p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId1"/><p:sldId id="258" r:id="rId8"/><p:sldId id="259"/>` +
        `<p:sldId id="260" r:id="rId9"/><p:sldId id="261" r:id="rId10"/></p:sldIdLst></p:presentation>`,
      files: {
        'ppt/_rels/presentation.xml.rels': relsXml([
          rel('rId1', 'slide', 'slides/slide1.xml'),
          rel('rId9', 'slide', 'http://example.com/slide.xml', true),
          rel('rId10', 'slide', 'slides/missing.xml'),
        ]),
      },
    }).then(full);
    expect(item(r, 'pptx.hiddenSlides')).toMatchObject({ count: 1, where: ['Slide 1'] });
  });

  it('still reads slide parts the list does not mention, by file name, and works with no presentation part at all', async () => {
    const hidden = slideXml({ attrs: 'show="0"' });
    const r = await build({ noPresentation: true, files: { 'ppt/slides/slide10.xml': hidden, 'ppt/slides/slide2.xml': hidden, 'ppt/slides/sub/slide3.xml': hidden } }).then(full);
    // slide1.xml from the builder is listed in no presentation either
    expect(item(r, 'pptx.hiddenSlides')).toMatchObject({ count: 2, where: ['slide10.xml', 'slide2.xml'] });
    expect(item(r, 'pptx.package.unaccountedParts')).toBeUndefined();
  });

  it('skips a slide that is not well-formed XML and reports the others', async () => {
    const r = await build({ slides: [{ raw: '<p:sld><oops' }, { attrs: 'show="0"' }] }).then(full);
    expect(item(r, 'pptx.hiddenSlides')).toMatchObject({ count: 1, where: ['Slide 2'] });
  });

  it('survives a presentation part that is not XML', async () => {
    const r = await build({ presentationRaw: 'garbage', slides: [{ attrs: 'show="0"' }] }).then(full);
    expect(item(r, 'pptx.hiddenSlides')?.where).toEqual(['slide1.xml']);
  });
});

describe('summary versus full', () => {
  it('reports only the summary ids by default and the finer ones at the full level', async () => {
    const shapes = gf(CHART) + sp({ spPr: `${frame}<a:gradFill/>`, body: para('<a:pPr rtl="1"/>') });
    const bytes = await build({ slides: [{ shapes }], presentation: '<p:sldSz cx="9144000" cy="6858000"/>' });
    expect(ids(await detectPptxLoss(bytes))).toEqual(['pptx.charts']);
    expect(ids(await detectPptxLoss(bytes, { detail: 'summary' }))).toEqual(['pptx.charts']);
    expect(ids(await full(bytes))).toEqual(['pptx.charts', 'pptx.shapes.gradientFill', 'pptx.slideSize', 'pptx.text.rtl']);
  });

  it('starts every effect with a state word and says which ids the reader does not read yet', async () => {
    const shapes = gf(CHART) + sp({ spPr: geom('ribbon') });
    const r = await build({ slides: [{ shapes, attrs: 'show="0"', rels: [rel('rId1', 'hyperlink', 'https://example.com/a', true)] }], files: { 'ppt/vbaProject.bin': 'x', 'ppt/odd/part.bin': 'x' } }).then(full);
    expect(r.items.length).toBeGreaterThan(5);
    for (const i of r.items) expect(i.effect).toMatch(/^(lost|approximated|kept): /);
    expect(item(r, 'pptx.charts')?.effect).toMatch(/the PPTX reader does not read this yet/);
    expect(item(r, 'pptx.hiddenSlides')?.effect).toMatch(/^kept: /);
  });

  it('caps where at five entries but counts them all', async () => {
    const r = await build({ slides: Array.from({ length: 8 }, () => ({ attrs: 'show="0"' })) }).then(full);
    expect(item(r, 'pptx.hiddenSlides')?.count).toBe(8);
    expect(item(r, 'pptx.hiddenSlides')?.where).toEqual(['Slide 1', 'Slide 2', 'Slide 3', 'Slide 4', 'Slide 5']);
  });
});

describe('package level', () => {
  const OLE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  const utf16 = (s: string) => [...s].flatMap((c) => [c.charCodeAt(0), 0]);

  it('names an encrypted package, already in the summary', async () => {
    const bytes = new Uint8Array([...OLE, ...new Array(64).fill(0), ...utf16('EncryptionInfo'), 0, 0, ...utf16('EncryptedPackage')]);
    for (const detail of ['summary', 'full'] as const) {
      const r = await detectPptxLoss(bytes, { detail });
      expect(ids(r)).toEqual(['pptx.encrypted']);
      expect(r.items[0]).toMatchObject({ count: 1, effect: expect.stringMatching(/^lost: cannot be opened/) });
    }
  });

  it('names a legacy binary presentation, already in the summary', async () => {
    const bytes = new Uint8Array([...OLE, ...new Array(64).fill(0), ...utf16('PowerPoint Document')]);
    const r = await detectPptxLoss(bytes);
    expect(ids(r)).toEqual(['pptx.legacyBinary']);
    expect(r.items[0].effect).toMatch(/^lost: not supported/);
  });

  it('names the parts nothing reads or reports, once, and not the known ones', async () => {
    const files = Object.fromEntries(
      [
        'docProps/core.xml', 'docProps/thumbnail.jpeg', 'ppt/presProps.xml', 'ppt/viewProps.xml', 'ppt/tableStyles.xml', 'ppt/theme/theme1.xml',
        'ppt/slideLayouts/slideLayout1.xml', 'ppt/slideMasters/slideMaster1.xml', 'ppt/printerSettings/printerSettings1.bin',
        'ppt/notesSlides/notesSlide1.xml', 'ppt/notesMasters/notesMaster1.xml', 'ppt/handoutMasters/handoutMaster1.xml', 'ppt/charts/chart1.xml',
        'ppt/diagrams/data1.xml', 'ppt/media/image1.png', 'ppt/embeddings/oleObject1.bin', 'ppt/comments/comment1.xml', 'ppt/commentAuthors.xml',
        'ppt/authors.xml', 'ppt/ink/ink1.xml', 'ppt/activeX/activeX1.xml', 'ppt/vbaProject.bin', 'ppt/vbaProjectSignature.bin', '_xmlsignatures/sig1.xml',
      ].map((n) => [n, '<x/>'])
    );
    const unknown = { 'customXml/item1.xml': '<x/>', 'ppt/tags/tag1.xml': '<x/>', 'ppt/fonts/font1.fntdata': 'x', 'ppt/slides/sub/slide1.xml': '<x/>' };
    const r = await build({ files: { ...files, ...unknown } }).then(full);
    expect(item(r, 'pptx.package.unaccountedParts')).toMatchObject({
      count: 4,
      where: ['customXml/item1.xml', 'ppt/tags/tag1.xml', 'ppt/fonts/font1.fntdata', 'ppt/slides/sub/slide1.xml'],
    });
  });

  it('caps the part names at five but counts them all, and leaves them out of the summary', async () => {
    const files = Object.fromEntries(Array.from({ length: 8 }, (_, i) => [`extra/p${i}.xml`, '<x/>']));
    const bytes = await build({ files });
    expect(item(await full(bytes), 'pptx.package.unaccountedParts')?.count).toBe(8);
    expect(item(await full(bytes), 'pptx.package.unaccountedParts')?.where).toHaveLength(5);
    expect(ids(await detectPptxLoss(bytes))).toEqual([]);
  });

  it('refuses entity expansion and external entities: the markup is read, the entities are not', async () => {
    const doctype = '<?xml version="1.0"?><!DOCTYPE p:sld [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;"><!ENTITY ext SYSTEM "file:///etc/passwd">]>';
    const slide = `${doctype}<p:sld ${NS} show="0"><p:cSld><p:spTree>${sp({ body: para('', run('lang="&lol2;"')) })}</p:spTree></p:cSld></p:sld>`;
    const r = await build({ slides: [{ raw: slide }] }).then(full);
    expect(item(r, 'pptx.hiddenSlides')?.count).toBe(1);
    // the language comes back exactly as written, not expanded to 'lollollol...'
    expect(item(r, 'pptx.text.lang')).toMatchObject({ count: 1, where: ['&lol2;'] });
  });

  it('ignores a DOCTYPE it cannot delimit rather than failing', async () => {
    const r = await build({ slides: [{ raw: `<!DOCTYPE p:sld [<!ENTITY a "b"` }, { attrs: 'show="0"' }] }).then(full);
    expect(item(r, 'pptx.hiddenSlides')).toMatchObject({ count: 1, where: ['Slide 2'] });
  });

  it('does not stack-overflow or hang on deeply nested XML', async () => {
    const deep = `<p:sld ${NS} show="0"><p:cSld>${'<a:g>'.repeat(5000)}${'</a:g>'.repeat(5000)}</p:cSld></p:sld>`;
    const r = await build({ slides: [{ raw: deep }, { attrs: 'show="0"' }] }).then(full);
    expect(item(r, 'pptx.hiddenSlides')).toMatchObject({ count: 1, where: ['Slide 2'] });
  });

  it('handles groups nested to the depth the parser allows', async () => {
    const depth = 60;
    const nest = (n: number): string => (n === 0 ? sp() : `<p:grpSp><p:nvGrpSpPr/><p:grpSpPr/>${nest(n - 1)}</p:grpSp>`);
    const r = await one(nest(depth));
    expect(item(r, 'pptx.shapes.groups.nested')?.count).toBe(depth - 1);
  });

  it('counts a very large number of objects quickly', async () => {
    const started = Date.now();
    const r = await one(text(para('', '<a:fld id="{1}" type="slidenum"><a:t>1</a:t></a:fld>'.repeat(20000))));
    expect(item(r, 'pptx.text.fields')?.count).toBe(20000);
    expect(Date.now() - started).toBeLessThan(10000);
  });

  it('does not scan a package with more entries than the limit', async () => {
    const saved = PACKAGE_LIMITS.entries;
    PACKAGE_LIMITS.entries = 5;
    try {
      const files = Object.fromEntries(Array.from({ length: 6 }, (_, i) => [`extra/p${i}.xml`, '<x/>']));
      const bytes = await build({ slides: [{ attrs: 'show="0"' }], files });
      for (const detail of ['summary', 'full'] as const) {
        const r = await detectPptxLoss(bytes, { detail });
        expect(ids(r)).toEqual(['pptx.sizeLimit']);
        expect(r.items[0]).toMatchObject({ count: 1, where: ['more than 5 parts'] });
      }
    } finally {
      PACKAGE_LIMITS.entries = saved;
    }
  });

  it('skips a part that is over the size limit and names it, in the summary too', async () => {
    const saved = { ...PACKAGE_LIMITS };
    PACKAGE_LIMITS.partBytes = 600;
    try {
      const big = slideXml({ attrs: 'show="0"' }).replace('</p:sld>', `<!-- ${'x'.repeat(2000)} --></p:sld>`);
      const bytes = await build({ slides: [{ raw: big }, { attrs: 'show="0"' }] });
      const r = await detectPptxLoss(bytes);
      expect(item(r, 'pptx.hiddenSlides')).toMatchObject({ count: 1, where: ['Slide 2'] });
      expect(item(r, 'pptx.sizeLimit')).toMatchObject({ count: 1, where: ['ppt/slides/slide1.xml'] });
    } finally {
      Object.assign(PACKAGE_LIMITS, saved);
    }
  });

  it('stops reading once the total size budget is spent', async () => {
    const saved = { ...PACKAGE_LIMITS };
    const filler = `<!-- ${'x'.repeat(700)} -->`;
    const slide = slideXml({ attrs: 'show="0"' }).replace('</p:sld>', `${filler}</p:sld>`);
    const bytes = await build({ slides: [{ raw: slide }, { raw: slide }, { raw: slide }] });
    PACKAGE_LIMITS.totalBytes = 2600;
    try {
      const r = await full(bytes);
      expect(item(r, 'pptx.hiddenSlides')?.count).toBeLessThan(3);
      expect(item(r, 'pptx.sizeLimit')?.count).toBeGreaterThanOrEqual(1);
    } finally {
      Object.assign(PACKAGE_LIMITS, saved);
    }
  });

  it('stops a forged-header bomb: the inflated size is counted as the part comes out', async () => {
    const saved = PACKAGE_LIMITS.partBytes;
    PACKAGE_LIMITS.partBytes = 100_000;
    try {
      const slide = slideXml({ attrs: 'show="0"' }).replace('</p:sld>', `<!--${' '.repeat(8_000_000)}--></p:sld>`);
      const bytes = await build({ slides: [{ raw: slide }] });
      // forge the uncompressed size (central directory and local header) of the slide part down to 1000 bytes
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      let forged = 0;
      for (let i = 0; i + 30 < bytes.length; i++) {
        const sig = view.getUint32(i, true);
        if (sig === 0x02014b50 && view.getUint32(i + 24, true) > 1_000_000) {
          view.setUint32(i + 24, 1000, true);
          forged++;
        }
        if (sig === 0x04034b50 && view.getUint32(i + 22, true) > 1_000_000) {
          view.setUint32(i + 22, 1000, true);
          forged++;
        }
      }
      expect(forged).toBeGreaterThan(0);
      const started = Date.now();
      const r = await full(bytes);
      expect(item(r, 'pptx.sizeLimit')).toMatchObject({ count: 1, where: ['ppt/slides/slide1.xml'] });
      expect(item(r, 'pptx.hiddenSlides')).toBeUndefined();
      expect(Date.now() - started).toBeLessThan(5000);
    } finally {
      PACKAGE_LIMITS.partBytes = saved;
    }
  });

  it('charges the bytes of parts that hit the limit to the total budget, so forged sizes on many parts stop the scan', async () => {
    const saved = { ...PACKAGE_LIMITS };
    PACKAGE_LIMITS.partBytes = 50_000;
    PACKAGE_LIMITS.totalBytes = 160_000;
    try {
      const files = Object.fromEntries(
        Array.from({ length: 12 }, (_, i) => [`ppt/slides/slide${i + 1}.xml`, `<p:sld ${NS}><!--${' '.repeat(2_000_000)}--></p:sld>`])
      );
      const bytes = await build({ slides: [], files });
      const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      for (let i = 0; i + 30 < bytes.length; i++) {
        const sig = view.getUint32(i, true);
        if (sig === 0x02014b50 && view.getUint32(i + 24, true) > 1_000_000) view.setUint32(i + 24, 1000, true);
        if (sig === 0x04034b50 && view.getUint32(i + 22, true) > 1_000_000) view.setUint32(i + 22, 1000, true);
      }
      const { OoxmlPackage } = await import('../common/ooxml-package');
      const pkg = new OoxmlPackage(await JSZip.loadAsync(bytes));
      const proto = Object.getPrototypeOf(await JSZip.loadAsync(bytes).then((z) => z.file('ppt/slides/slide1.xml')));
      const real = proto.internalStream;
      let inflated = 0;
      proto.internalStream = function (this: unknown, ...args: unknown[]) {
        inflated++;
        return real.apply(this, args);
      };
      try {
        for (let i = 1; i <= 12; i++) await pkg.text(`ppt/slides/slide${i}.xml`);
      } finally {
        proto.internalStream = real;
      }
      // the parts after the budget is spent are refused on their declared size, without being inflated
      expect(inflated).toBeLessThan(8);
      // without the charge each of the 12 parts would inflate 50,000 bytes (600,000); with it the budget runs out after
      // about 160,000 (plus at most one stream chunk of overshoot) and the rest are refused without inflating
      expect(pkg.read).toBeGreaterThanOrEqual(160_000);
      expect(pkg.read).toBeLessThan(160_000 + 70_000);
      expect(pkg.skipped).toHaveLength(12);
    } finally {
      Object.assign(PACKAGE_LIMITS, saved);
    }
  });

  it('inflates nothing from a zip bomb: the declared size is checked before the part is read', async () => {
    const slide = slideXml().replace('</p:sld>', `<!--${' '.repeat(PACKAGE_LIMITS.partBytes + 1024)}--></p:sld>`);
    const bytes = await build({ slides: [{ raw: slide }] });
    expect(bytes.length).toBeLessThan(200_000);
    const started = Date.now();
    const r = await full(bytes);
    expect(item(r, 'pptx.sizeLimit')).toMatchObject({ count: 1, where: ['ppt/slides/slide1.xml'] });
    expect(Date.now() - started).toBeLessThan(5000);
  });

  it('treats a part that cannot be inflated as unreadable, not as a failure', async () => {
    const bytes = await build({ slides: [{ attrs: 'show="0"' }, { attrs: 'show="0"' }] });
    // corrupt the compressed data of slide 1 (find its local header, flip bytes inside the deflate stream)
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const name = new TextEncoder().encode('ppt/slides/slide1.xml');
    for (let i = 0; i + 30 < bytes.length; i++) {
      if (view.getUint32(i, true) === 0x04034b50 && view.getUint16(i + 26, true) === name.length && bytes.subarray(i + 30, i + 30 + name.length).every((b, k) => b === name[k])) {
        const start = i + 30 + name.length + view.getUint16(i + 28, true);
        for (let k = 0; k < 24; k++) bytes[start + k] = 0xff;
      }
    }
    const r = await full(bytes);
    expect(item(r, 'pptx.hiddenSlides')).toMatchObject({ count: 1, where: ['Slide 2'] });
  });

  it('checks the size after reading when the archive gives no declared size up front', async () => {
    const saved = PACKAGE_LIMITS.partBytes;
    const real = JSZip.loadAsync.bind(JSZip);
    const spy = jest.spyOn(JSZip, 'loadAsync').mockImplementation(async (data: Parameters<typeof JSZip.loadAsync>[0]) => {
      const zip = await real(data);
      for (const file of Object.values(zip.files)) {
        const record = (file as unknown as { _data?: { uncompressedSize?: number } })._data;
        if (!record) continue;
        // the first look (the detector's) finds nothing; JSZip's own consistency check afterwards sees the real size
        const actual = record.uncompressedSize;
        let looks = 0;
        Object.defineProperty(record, 'uncompressedSize', { get: () => (looks++ === 0 ? undefined : actual) });
      }
      return zip;
    });
    try {
      const big = slideXml({ attrs: 'show="0"' }).replace('</p:sld>', `<!-- ${'x'.repeat(4000)} --></p:sld>`);
      const bytes = await build({ slides: [{ raw: big }, { attrs: 'show="0"' }] });
      PACKAGE_LIMITS.partBytes = 1500;
      const r = await full(bytes);
      expect(item(r, 'pptx.hiddenSlides')).toMatchObject({ count: 1, where: ['Slide 2'] });
      expect(item(r, 'pptx.sizeLimit')).toMatchObject({ count: 1, where: ['ppt/slides/slide1.xml'] });
    } finally {
      PACKAGE_LIMITS.partBytes = saved;
      spy.mockRestore();
    }
  });
});

// ---------------------------------------------------------------------------
// Full-level ids: each id has a table row here (what is counted, what is not).
// ---------------------------------------------------------------------------

const countOf = (r: ImportLossReport, id: string): number => item(r, id)?.count ?? 0;
const cell = (tcPr = '<a:tcPr/>', attrs = ''): string => `<a:tc${attrs ? ` ${attrs}` : ''}><a:txBody><a:bodyPr/><a:p/></a:txBody>${tcPr}</a:tc>`;
const table = (rows: string, tblPr = '<a:tblPr/>'): string =>
  gf('http://schemas.openxmlformats.org/drawingml/2006/table', `<a:tbl>${tblPr}<a:tblGrid/><a:tr h="1">${rows}</a:tr></a:tbl>`);
const body = (pPr: string): string => text(para(pPr));
const withRun = (rPr: string): string => text(para('', run(rPr)));
const bodyPrOf = (attrs: string, children = ''): string => sp({ bodyPr: `<a:bodyPr ${attrs}>${children}</a:bodyPr>`, body: para() });

describe('masters and theme', () => {
  const master = (clrMap: string, n = 1) => ({
    [`ppt/slideMasters/slideMaster${n}.xml`]: `<p:sldMaster ${NS}><p:cSld/>${clrMap}</p:sldMaster>`,
  });
  const defaultMap = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"/>';

  it('counts the slide masters of a deck with several, and the handout and notes masters', async () => {
    const several = await build({ files: { ...master(defaultMap, 1), ...master(defaultMap, 2), 'ppt/handoutMasters/handoutMaster1.xml': '<x/>', 'ppt/notesMasters/notesMaster1.xml': '<x/>' } }).then(full);
    expect(countOf(several, 'pptx.masters.multiple')).toBe(2);
    expect(countOf(several, 'pptx.masters.handoutMaster')).toBe(1);
    expect(countOf(several, 'pptx.masters.notesMaster')).toBe(1);
    const single = await build({ files: master(defaultMap) }).then(full);
    expect(ids(single)).toEqual([]);
  });

  it('counts placeholders that take their position from the layout, in shapes, pictures and graphic frames', async () => {
    const ph = '<p:ph type="body" idx="1"/>';
    const frameWithPh = (xfrm: string) =>
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="g"/><p:cNvGraphicFramePr/><p:nvPr>${ph}</p:nvPr></p:nvGraphicFramePr>${xfrm}<a:graphic><a:graphicData uri="u"/></a:graphic></p:graphicFrame>`;
    const r = await one(
      sp({ ph, spPr: '' }) + // inherits
        sp({ ph }) + // has its own frame
        `<p:pic><p:nvPicPr><p:cNvPr id="4" name="p"/><p:cNvPicPr/><p:nvPr>${ph}</p:nvPr></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>` + // inherits
        frameWithPh('') + // inherits
        frameWithPh('<p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm>') +
        sp({ spPr: '' }) + // no placeholder
        '<p:sp><p:spPr/></p:sp>' // no nv element at all
    );
    expect(item(r, 'pptx.masters.layoutInheritance')).toMatchObject({ count: 3, where: ['Slide 1'] });
  });

  it('counts a colour map that is not the default: on a master (any of the four roles), a layout or a slide', async () => {
    const r = await build({
      slides: [{ tail: '<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"/></p:clrMapOvr>' }, { tail: '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' }],
      files: {
        ...master('<p:clrMap bg1="dk1" tx1="lt1" bg2="lt2" tx2="dk2"/>', 1),
        ...master('<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"/>', 2),
        ...master('<p:clrMap/>', 3),
        ...master('<p:clrMap bg1="lt1" tx1="dk1" bg2="dk2" tx2="dk2"/>', 4),
        'ppt/slideLayouts/slideLayout1.xml': `<p:sldLayout ${NS}><p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"/></p:clrMapOvr></p:sldLayout>`,
        'ppt/slideLayouts/slideLayout2.xml': `<p:sldLayout ${NS}><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
      },
    }).then(full);
    expect(item(r, 'pptx.theme.colorMap')).toMatchObject({ count: 4, where: ['Slide 1', 'slideMaster1.xml', 'slideMaster4.xml', 'slideLayout1.xml'] });
  });

  it('counts the font and effect schemes of the themes the masters use, not other themes', async () => {
    const theme = (major: string, minor: string, effects: string) =>
      `<a:theme ${NS}><a:themeElements><a:fontScheme name="t"><a:majorFont><a:latin typeface="${major}"/></a:majorFont><a:minorFont><a:latin typeface="${minor}"/></a:minorFont></a:fontScheme><a:fmtScheme><a:effectStyleLst>${effects}</a:effectStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
    const style = (inner: string) => `<a:effectStyle>${inner}</a:effectStyle>`;
    const link = (n: number) => ({
      [`ppt/slideMasters/slideMaster${n}.xml`]: `<p:sldMaster ${NS}/>`,
      [`ppt/slideMasters/_rels/slideMaster${n}.xml.rels`]: relsXml([rel('rId1', 'theme', `../theme/theme${n}.xml`)]),
    });
    const r = await build({
      files: {
        ...link(1), ...link(2), ...link(3), ...link(4), ...link(5),
        'ppt/theme/theme1.xml': theme('Calibri Light', 'Calibri', style('<a:effectLst/>')), // Office default fonts, no real effect
        'ppt/theme/theme2.xml': theme('Georgia', 'Calibri', style('<a:effectLst><a:outerShdw/></a:effectLst>')),
        'ppt/theme/theme3.xml': theme('Verdana', 'Georgia', style('<a:effectDag/>')),
        'ppt/theme/theme4.xml': theme('', 'calibri', style('<a:effectLst><a:other/></a:effectLst>')),
        // theme5.xml is missing; theme6.xml belongs to no master
        'ppt/theme/theme6.xml': theme('Papyrus', 'Papyrus', style('<a:effectLst><a:glow/></a:effectLst>')),
      },
    }).then(full);
    expect(item(r, 'pptx.theme.fontScheme')).toMatchObject({ count: 2, where: ['Georgia', 'Verdana / Georgia'] });
    expect(item(r, 'pptx.theme.effectScheme')).toMatchObject({ count: 2, where: ['theme2.xml', 'theme3.xml'] });
  });
});

describe('text', () => {
  it.each([
    ['normAutofit with a font scale', bodyPrOf('', '<a:normAutofit fontScale="92500" lnSpcReduction="10000"/>'), 1],
    ['normAutofit with only a line spacing reduction', bodyPrOf('', '<a:normAutofit lnSpcReduction="10000"/>'), 1],
    ['shape that resizes to fit its text', bodyPrOf('', '<a:spAutoFit/>'), 1],
    ['normAutofit that shrinks nothing', bodyPrOf('', '<a:normAutofit/>'), 0],
    ['no autofit', bodyPrOf('', '<a:noAutofit/>'), 0],
  ])('autofit: %s', async (_name, shapes, expected) => {
    expect(countOf(await one(shapes), 'pptx.text.autofit')).toBe(expected);
  });

  it.each([
    ['bullet colour', '<a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buChar char="&#8226;"/>', 1],
    ['bullet size (percent)', '<a:buSzPct val="80000"/>', 1],
    ['bullet size (points)', '<a:buSzPts val="1200"/>', 1],
    ['a bullet character other than the round bullet', '<a:buChar char="&#8211;"/>', 1],
    ['a bullet character with no char at all', '<a:buChar/>', 1],
    ['numbering with a start value', '<a:buAutoNum type="arabicPeriod" startAt="3"/>', 1],
    ['numbering in another scheme', '<a:buAutoNum type="alphaLcPeriod"/>', 1],
    ['the plain round bullet', '<a:buFont typeface="Arial"/><a:buChar char="&#8226;"/>', 0],
    ['plain arabic numbering', '<a:buAutoNum type="arabicPeriod"/>', 0],
    ['no bullet', '<a:buNone/>', 0],
  ])('custom bullets: %s', async (_name, inner, expected) => {
    expect(countOf(await one(body(`<a:pPr>${inner}</a:pPr>`)), 'pptx.text.bullets.custom')).toBe(expected);
  });

  it('picture bullets, tab stops, right-to-left paragraphs', async () => {
    const r = await one(
      body('<a:pPr><a:buBlip><a:blip r:embed="rId1"/></a:buBlip></a:pPr>') +
        body('<a:pPr><a:tabLst><a:tab pos="100" algn="l"/><a:tab pos="200"/></a:tabLst></a:pPr>') +
        body('<a:pPr><a:tabLst/></a:pPr>') +
        body('<a:pPr rtl="1"/>') +
        body('<a:pPr rtl="true"/>') +
        body('<a:pPr rtl="0"/>')
    );
    expect(countOf(r, 'pptx.text.bullets.picture')).toBe(1);
    expect(countOf(r, 'pptx.text.tabStops')).toBe(1);
    expect(countOf(r, 'pptx.text.rtl')).toBe(2);
  });

  it('columns, vertical text (text boxes and table cells), WordArt warps', async () => {
    const r = await one(
      bodyPrOf('numCol="2"') + bodyPrOf('numCol="1"') + bodyPrOf('') +
        bodyPrOf('vert="vert"') + bodyPrOf('vert="horz"') +
        bodyPrOf('', '<a:prstTxWarp prst="textArchUp"/>') + bodyPrOf('', '<a:prstTxWarp prst="textNoShape"/>') +
        table(cell('<a:tcPr vert="vert270"/>') + cell('<a:tcPr vert="horz"/>') + cell())
    );
    expect(countOf(r, 'pptx.text.columns')).toBe(1);
    expect(countOf(r, 'pptx.text.vertical')).toBe(2);
    expect(countOf(r, 'pptx.text.wordArt')).toBe(1);
  });

  it('fields', async () => {
    const r = await one(text(para('', '<a:fld id="{1}" type="slidenum"><a:t>1</a:t></a:fld><a:fld id="{2}" type="datetime1"><a:t>1/1</a:t></a:fld>')));
    expect(item(r, 'pptx.text.fields')).toMatchObject({ count: 2, where: ['Slide 1'] });
  });

  it.each([
    ['character spacing', 'spc="100"', 1],
    ['negative character spacing', 'spc="-50"', 1],
    ['zero character spacing', 'spc="0"', 0],
    ['all caps', 'cap="all"', 1],
    ['small caps', 'cap="small"', 1],
    ['no caps', 'cap="none"', 0],
    ['spacing and caps on one run', 'spc="10" cap="all"', 1],
  ])('spacing: %s', async (_n, rPr, expected) => {
    expect(countOf(await one(withRun(rPr)), 'pptx.text.spacing')).toBe(expected);
  });

  it.each([
    ['superscript', 'baseline="30000"', 1],
    ['subscript', 'baseline="-25000"', 1],
    ['baseline 0', 'baseline="0"', 0],
    ['no baseline', 'b="1"', 0],
  ])('baseline: %s', async (_n, rPr, expected) => {
    expect(countOf(await one(withRun(rPr)), 'pptx.text.baseline')).toBe(expected);
  });

  it('counts runs tagged with a language other than en-US, naming the languages', async () => {
    const r = await one(withRun('lang="de-DE"') + withRun('lang="de-DE"') + withRun('lang="ja-JP"') + withRun('lang="en-US"') + withRun('lang="EN-us"') + withRun('lang=""') + withRun('b="1"'));
    expect(item(r, 'pptx.text.lang')).toMatchObject({ count: 3, where: ['de-DE', 'ja-JP'] });
  });
});

describe('shapes', () => {
  it('counts effects on shapes, groups and runs; an empty effect list is none', async () => {
    const r = await one(
      sp({ spPr: `${frame}<a:effectLst><a:outerShdw blurRad="1"/></a:effectLst>` }) +
        sp({ spPr: `${frame}<a:effectLst/>` }) +
        sp({ spPr: `${frame}<a:effectDag/>` }) +
        withRun('lang="en-US"').replace('<a:rPr lang="en-US"/>', '<a:rPr lang="en-US"><a:effectLst><a:glow rad="1"/></a:effectLst></a:rPr>') +
        '<p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:effectLst><a:softEdge rad="1"/></a:effectLst></p:grpSpPr></p:grpSp>' +
        sp({ spPr: `${frame}<a:ln><a:effectLst><a:glow/></a:effectLst></a:ln>` }) // an effect list in another place is not a shape effect
    );
    expect(countOf(r, 'pptx.shapes.effects')).toBe(4);
  });

  it('counts gradient and pattern fills on shapes, cells and backgrounds, not on outlines', async () => {
    const r = await build({
      slides: [
        {
          raw: slideXml({
            shapes:
              sp({ spPr: `${frame}<a:gradFill/>` }) +
              sp({ spPr: `${frame}<a:pattFill prst="pct5"/>` }) +
              sp({ spPr: `${frame}<a:ln><a:gradFill/><a:pattFill/></a:ln>` }) +
              table(cell('<a:tcPr><a:gradFill/></a:tcPr>') + cell('<a:tcPr><a:pattFill/></a:tcPr>')),
          }).replace('<p:cSld>', '<p:cSld><p:bg><p:bgPr><a:gradFill/></p:bgPr></p:bg>'),
        },
      ],
    }).then(full);
    expect(countOf(r, 'pptx.shapes.gradientFill')).toBe(3);
    expect(countOf(r, 'pptx.shapes.patternFill')).toBe(2);
  });

  it('counts groups inside groups (the inner ones), not a flat group', async () => {
    const grp = (inner: string) => `<p:grpSp><p:nvGrpSpPr/><p:grpSpPr/>${inner}</p:grpSp>`;
    expect(countOf(await one(grp(sp()) + grp(sp())), 'pptx.shapes.groups.nested')).toBe(0);
    expect(countOf(await one(grp(grp(sp()) + grp(grp(sp())))), 'pptx.shapes.groups.nested')).toBe(3);
  });

  it('counts 3-D scene and shape properties once per holder', async () => {
    const r = await one(
      sp({ spPr: `${frame}<a:scene3d><a:camera prst="orthographicFront"/></a:scene3d><a:sp3d/>` }) +
        sp({ spPr: `${frame}<a:sp3d extrusionH="1"/>` }) +
        sp({ bodyPr: '<a:bodyPr><a:scene3d/></a:bodyPr>' })
    );
    expect(countOf(r, 'pptx.shapes.scene3d')).toBe(3);
  });

  it('counts elbow and curved connectors (connector shapes and shapes with a connector geometry)', async () => {
    const cxn = (prst: string) => `<p:cxnSp><p:nvCxnSpPr/><p:spPr>${geom(prst)}</p:spPr></p:cxnSp>`;
    const r = await one(cxn('bentConnector3') + cxn('curvedConnector2') + cxn('straightConnector1'));
    expect(countOf(r, 'pptx.shapes.connectors.elbowCurved')).toBe(2);
  });
});

describe('images and tables', () => {
  it('counts cropped pictures: a non-zero crop edge, not a zero or empty crop', async () => {
    const r = await one(
      pic('<a:blip r:embed="rId1"/><a:srcRect l="10000"/>') +
        pic('<a:blip r:embed="rId1"/><a:srcRect b="-5000" r="0"/>') +
        pic('<a:blip r:embed="rId1"/><a:srcRect l="0" t="0"/>') +
        pic('<a:blip r:embed="rId1"/><a:srcRect/>') +
        pic('<a:blip r:embed="rId1"/>') +
        pic('<a:srcRect t="1"/>')
    );
    expect(countOf(r, 'pptx.images.crop')).toBe(3);
  });

  it('counts recoloured, linked and SVG pictures', async () => {
    const r = await one(
      pic('<a:blip r:embed="rId1"><a:duotone/></a:blip>') +
        pic('<a:blip r:embed="rId1"><a:lum bright="10000"/></a:blip>') +
        pic('<a:blip r:embed="rId1"><a:alphaModFix amt="50000"/></a:blip>') +
        pic('<a:blip r:link="rId2"/>') +
        pic('<a:blip r:embed="rId1"/>') +
        pic('<a:blip r:embed="rId1"><a:extLst><a:ext uri="{x}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId3"/></a:ext></a:extLst></a:blip>')
    );
    expect(countOf(r, 'pptx.images.recolor')).toBe(2);
    expect(countOf(r, 'pptx.images.linked')).toBe(1);
    expect(countOf(r, 'pptx.images.svg')).toBe(1);
  });

  it('counts merged cell regions (the anchor cell), not the cells they cover', async () => {
    const r = await one(table(cell(undefined, 'gridSpan="2"') + cell(undefined, 'hMerge="1"') + cell(undefined, 'rowSpan="3"') + cell(undefined, 'gridSpan="1"') + cell(undefined, 'vMerge="1"') + cell()));
    expect(countOf(r, 'pptx.tables.merge')).toBe(2);
  });

  it('counts tables with a table style (a reference or header/total flags) and banded tables', async () => {
    const r = await one(
      table(cell(), '<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>') +
        table(cell(), '<a:tblPr lastCol="true" bandCol="true"/>') +
        table(cell(), '<a:tblPr bandRow="0" firstRow="0"/>') +
        table(cell()) +
        table(cell(), '') // no table properties at all
    );
    expect(countOf(r, 'pptx.tables.styles')).toBe(2);
    expect(countOf(r, 'pptx.tables.banding')).toBe(2);
  });
});

describe('presentation level', () => {
  const ext = (inner: string) => `<p:extLst><p:ext uri="x">${inner}</p:ext></p:extLst>`;

  it('counts sections and custom shows', async () => {
    const sections = ext('<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p14:section name="A" id="{1}"><p14:sldIdLst/></p14:section><p14:section name="B" id="{2}"/></p14:sectionLst>');
    const shows = '<p:custShowLst><p:custShow name="a" id="0"><p:sldLst/></p:custShow><p:custShow name="b" id="1"/></p:custShowLst>';
    const r = await build({ presentation: shows + sections }).then(full);
    expect(countOf(r, 'pptx.sections')).toBe(2);
    expect(countOf(r, 'pptx.customShows')).toBe(2);
  });

  it.each([
    ['12192000', '6858000', 0, undefined],
    ['9144000', '5143500', 0, undefined],
    ['9144000', '6858000', 1, '720 x 540 pt'],
    ['10692000', '7560000', 1, '842 x 595 pt'],
    ['6858000', '9144000', 1, '540 x 720 pt'],
    ['0', '6858000', 0, undefined],
    ['wide', 'tall', 0, undefined],
  ])('slide size %s x %s EMU: %i other than 16:9', async (cx, cy, expected, where) => {
    const r = await build({ presentation: `<p:sldSz cx="${cx}" cy="${cy}"/>` }).then(full);
    expect(countOf(r, 'pptx.slideSize')).toBe(expected);
    expect(item(r, 'pptx.slideSize')?.where).toEqual(where ? [where] : undefined);
  });

  it('is silent about a presentation with no slide size', async () => {
    expect(countOf(await build({ presentation: '' }).then(full), 'pptx.slideSize')).toBe(0);
  });

  it('counts slide and section zoom objects', async () => {
    const zoom = (name: string) => `<mc:AlternateContent><mc:Choice Requires="p166"><p:graphicFrame><a:graphic><a:graphicData uri="u"><p166:${name} xmlns:p166="http://schemas.microsoft.com/office/powerpoint/2016/6/main"/></a:graphicData></a:graphic></p:graphicFrame></mc:Choice></mc:AlternateContent>`;
    const r = await one(zoom('sldZmObj') + zoom('secZmObj') + zoom('sldZm') + zoom('secZm') + zoom('other'));
    expect(countOf(r, 'pptx.zoom')).toBe(4);
  });

  it('counts ink drawings by their content parts, falling back to the ink parts', async () => {
    const r = await one(`<mc:AlternateContent><mc:Choice Requires="p14"><p:contentPart r:id="rId1"/></mc:Choice></mc:AlternateContent><p:contentPart r:id="rId2"/>`);
    expect(item(r, 'pptx.ink')).toMatchObject({ count: 2, where: ['Slide 1'] });
    const parts = await build({ files: { 'ppt/ink/ink1.xml': '<x/>', 'ppt/ink/ink2.xml': '<x/>' } }).then(full);
    expect(countOf(parts, 'pptx.ink')).toBe(2);
  });

  it('counts equations (a14:m), not other elements called m', async () => {
    const r = await one(
      text(para('', '<a:r><a:t>x</a:t></a:r>')) +
        `<mc:AlternateContent><mc:Choice Requires="a14"><p:sp><p:txBody><a:p><a14:m><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"/></a14:m></a:p></p:txBody></p:sp></mc:Choice><mc:Fallback><p:sp/></mc:Fallback></mc:AlternateContent>` +
        text('<a:p><m xmlns="urn:other"/></a:p>')
    );
    expect(countOf(r, 'pptx.equations')).toBe(1);
  });

  it('counts digital signatures by their signature parts', async () => {
    const r = await build({ files: { '_xmlsignatures/sig1.xml': '<x/>', '_xmlsignatures/sig2.xml': '<x/>', '_xmlsignatures/origin.sigs': '', '_xmlsignatures/_rels/origin.sigs.rels': '<x/>' } }).then(full);
    expect(countOf(r, 'pptx.signatures')).toBe(2);
  });

  it('counts external link targets by scheme, not linked pictures and not internal relationships', async () => {
    const r = await build({
      slides: [
        {
          rels: [
            rel('rId1', 'hyperlink', 'https://example.com/a', true),
            rel('rId2', 'hyperlink', 'HTTPS://example.com/b', true),
            rel('rId3', 'hyperlink', 'mailto:someone@example.com', true),
            rel('rId4', 'oleObject', 'file:///C:/book.xlsx', true),
            rel('rId5', 'hyperlink', 'slides.pptx', true),
            rel('rId6', 'image', 'https://example.com/pic.png', true),
            rel('rId7', 'image', '../media/pic.png'),
            rel('rId8', 'hyperlink', 'slide2.xml'),
          ],
        },
      ],
    }).then(full);
    expect(item(r, 'pptx.externalLinks')).toMatchObject({ count: 5, where: ['https', 'mailto', 'file', 'relative'] });
    expect(countOf(r, 'pptx.images.linked')).toBe(0);
  });

  it('counts action hyperlinks by kind, never recording a macro name', async () => {
    const link = (tag: string, action?: string) => `<a:${tag} r:id=""${action === undefined ? '' : ` action="${action}"`}/>`;
    const shape = (inner: string) => `<p:sp><p:nvSpPr><p:cNvPr id="2" name="s">${inner}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`;
    const r = await one(
      shape(link('hlinkClick', 'ppaction://hlinkshowjump?jump=nextslide')) +
        shape(link('hlinkClick', 'ppaction://hlinkshowjump?jump=lastslide')) +
        shape(link('hlinkClick', 'ppaction://macro?name=Secret.Module')) +
        shape(link('hlinkHover', 'ppaction://hlinkfile')) +
        shape(link('hlinkClick', 'something-else')) +
        shape(link('hlinkClick', '')) +
        shape(link('hlinkClick'))
    );
    expect(item(r, 'pptx.actions')).toMatchObject({ count: 5, where: ['hlinkshowjump', 'macro', 'hlinkfile', 'other'] });
    expect(JSON.stringify(r)).not.toContain('Secret');
  });
});

describe('detectPptxLoss: the state of each id after the reader (K9c, K10a)', () => {
  // the ids `importPptx` reads say how (kept or approximated); the rest are still lost
  const READ: Record<string, 'kept' | 'approximated'> = {
    'pptx.fonts.unavailable': 'approximated',
    'pptx.shapes.unsupportedGeometry': 'approximated',
    'pptx.hiddenSlides': 'kept',
    'pptx.notes': 'approximated',
    'pptx.masters.notesMaster': 'approximated',
    'pptx.masters.layoutInheritance': 'approximated',
    'pptx.theme.colorMap': 'approximated',
    'pptx.theme.fontScheme': 'approximated',
    'pptx.text.autofit': 'approximated',
    'pptx.text.bullets.custom': 'approximated',
    'pptx.text.tabStops': 'approximated',
    'pptx.text.columns': 'approximated',
    'pptx.text.vertical': 'approximated',
    'pptx.text.fields': 'approximated',
    'pptx.text.rtl': 'approximated',
    'pptx.text.spacing': 'approximated',
    'pptx.text.baseline': 'approximated',
    'pptx.text.lang': 'kept',
    'pptx.shapes.custGeom': 'approximated',
    'pptx.shapes.connectors.elbowCurved': 'kept',
    'pptx.shapes.groups.nested': 'kept',
    'pptx.shapes.gradientFill': 'approximated',
    'pptx.shapes.patternFill': 'approximated',
    'pptx.slideSize': 'kept',
    'pptx.images.unsupportedFormat': 'approximated',
    'pptx.images.crop': 'kept',
    'pptx.images.svg': 'approximated',
    'pptx.tables.merge': 'kept',
    'pptx.tables.styles': 'approximated',
    'pptx.tables.banding': 'approximated',
    'pptx.actions': 'approximated',
    'pptx.externalLinks': 'approximated',
  };

  it('reports every id the reader reads as kept or approximated and every other id as lost', async () => {
    const bytes = readFileSync(join(__dirname, '../../../../fixtures/pptx/kitchen-sink.pptx'));
    const report = await detectPptxLoss(new Uint8Array(bytes), { detail: 'full' });
    const seen = new Set<string>();
    for (const it of report.items) {
      const state = it.effect.split(':')[0];
      expect([it.id, state]).toEqual([it.id, READ[it.id] ?? 'lost']);
      seen.add(it.id);
    }
    // the fixture exercises most of the read ids, so the table above is not vacuous
    expect([...seen].filter((id) => READ[id]).length).toBeGreaterThan(8);
  });
});
