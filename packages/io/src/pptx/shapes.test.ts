import JSZip from 'jszip';
import type { PptxShape, PptxSlide, PresentationImportResult } from '@b32nio/spindle-shared/export/pptx';
import { importPptx } from './reader';
import * as shapesModule from './shapes';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PRES_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
const CLR_MAP = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
const THEME = `<a:theme ${NS} name="T"><a:themeElements><a:clrScheme name="c">${['dk1:000000', 'lt1:FFFFFF', 'dk2:44546A', 'lt2:E7E6E6', 'accent1:4472C4', 'accent2:ED7D31', 'accent3:A5A5A5', 'accent4:FFC000', 'accent5:5B9BD5', 'accent6:70AD47', 'hlink:0563C1', 'folHlink:954F72']
  .map((c) => `<a:${c.split(':')[0]}><a:srgbClr val="${c.split(':')[1]}"/></a:${c.split(':')[0]}>`)
  .join('')}</a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;

const rel = (id: string, type: string, target: string): string => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`;
const rels = (list: string[]): string => `<Relationships xmlns="${PKG_REL}">${list.join('')}</Relationships>`;
const E = 9525;

interface Deck {
  /** Children of the slide's `p:spTree`. */
  tree?: string;
  /** Attributes and children of `p:sld` besides `p:cSld`. */
  sldTail?: string;
  bg?: string;
  /** Children of the layout's / master's `p:spTree`. */
  layoutTree?: string;
  masterTree?: string;
  layoutAttrs?: string;
  layoutTail?: string;
  masterTail?: string;
  defaultStyle?: string;
  extra?: Record<string, string>;
  slideRels?: string[];
  noLayoutRel?: boolean;
  layoutWithoutMaster?: boolean;
}

async function build(d: Deck = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="${CT}"><Override PartName="/ppt/presentation.xml" ContentType="${PRES_TYPE}"/></Types>`);
  zip.file('_rels/.rels', rels([rel('rId1', 'officeDocument', 'ppt/presentation.xml')]));
  zip.file('ppt/_rels/presentation.xml.rels', rels([rel('rId1', 'slide', 'slides/slide1.xml'), rel('rIdM', 'slideMaster', 'slideMasters/slideMaster1.xml')]));
  zip.file('ppt/presentation.xml', `<p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/>${d.defaultStyle ?? ''}</p:presentation>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${NS}><p:cSld>${d.bg ?? ''}<p:spTree>${d.tree ?? ''}</p:spTree></p:cSld>${d.sldTail ?? ''}</p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', rels([...(d.noLayoutRel ? [] : [rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml')]), ...(d.slideRels ?? [])]));
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout ${NS}${d.layoutAttrs ?? ''}><p:cSld name="L"><p:spTree>${d.layoutTree ?? ''}</p:spTree></p:cSld>${d.layoutTail ?? ''}</p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels(d.layoutWithoutMaster ? [] : [rel('rId1', 'slideMaster', '../slideMasters/slideMaster1.xml')]));
  zip.file('ppt/slideMasters/slideMaster1.xml', `<p:sldMaster ${NS}><p:cSld><p:spTree>${d.masterTree ?? ''}</p:spTree></p:cSld>${CLR_MAP}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>${d.masterTail ?? ''}</p:sldMaster>`);
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', rels([rel('rIdT', 'theme', '../theme/theme1.xml'), rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml')]));
  zip.file('ppt/theme/theme1.xml', THEME);
  for (const [name, content] of Object.entries(d.extra ?? {})) zip.file(name, content);
  return zip.generateAsync({ type: 'uint8array' });
}

const read = async (d: Deck = {}): Promise<PresentationImportResult> => importPptx(await build(d));
const shapesOf = async (tree: string, d: Deck = {}): Promise<PptxShape[]> => (await read({ ...d, tree })).slides[0].elements as PptxShape[];
const one = async (tree: string, d: Deck = {}): Promise<PptxShape> => (await shapesOf(tree, d))[0];

const xfrm = (x: number, y: number, w: number, h: number, attrs = ''): string => `<a:xfrm${attrs}><a:off x="${x * E}" y="${y * E}"/><a:ext cx="${w * E}" cy="${h * E}"/></a:xfrm>`;
const sp = (opts: { id?: number; name?: string; nvSp?: string; nvPr?: string; spPr?: string; body?: string; style?: string; cNvPr?: string } = {}): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${opts.id ?? 2}" name="${opts.name ?? 'S'}"${opts.cNvPr ?? ''}/><p:cNvSpPr${opts.nvSp ?? ''}/><p:nvPr>${opts.nvPr ?? ''}</p:nvPr></p:nvSpPr><p:spPr>${opts.spPr ?? ''}</p:spPr>${opts.style ?? ''}${opts.body ?? ''}</p:sp>`;
const txBody = (paras: string, bodyPr = '', lst = ''): string => `<p:txBody><a:bodyPr${bodyPr}/>${lst || '<a:lstStyle/>'}${paras}</p:txBody>`;
const para = (text: string, rPr = '', pPr = ''): string => `<a:p>${pPr}<a:r>${rPr}<a:t>${text}</a:t></a:r></a:p>`;
const srgb = (hex: string, inner = ''): string => `<a:srgbClr val="${hex}">${inner}</a:srgbClr>`;

describe('p:sp: identity and frame', () => {
  it('reads id, name, alt text, hidden, text box flag and the frame in px with rotation and flips', async () => {
    const s = await one(sp({ id: 7, name: ' Box ', cNvPr: ' descr=" a box " hidden="1"', nvSp: ' txBox="1"', spPr: xfrm(10, 20, 300, 40, ' rot="5400000" flipH="1" flipV="true"') }));
    expect(s).toEqual({
      kind: 'shape', pptxId: '7', name: 'Box', altText: 'a box', hidden: true, textBox: true,
      frame: { x: 10, y: 20, w: 300, h: 40, rot: 90, flipH: true, flipV: true },
    });
  });

  it('leaves out a frame that is incomplete or negative and an id that is missing', async () => {
    const s = await one('<p:sp><p:spPr><a:xfrm><a:off x="1" y="1"/></a:xfrm></p:spPr></p:sp>');
    expect(s).toEqual({ kind: 'shape', pptxId: '', name: '' });
    expect((await one(sp({ spPr: xfrm(0, 0, -5, 5) }))).frame).toBeUndefined();
  });

  it('reads p:sp among the shapes of a slide: groups and connectors are left to the later steps, and mc:AlternateContent is entered (Choice, else Fallback)', async () => {
    const tree =
      sp({ id: 2 }) +
      '<p:grpSp/><p:cxnSp/>' +
      `<mc:AlternateContent><mc:Choice Requires="x">${sp({ id: 3 })}</mc:Choice><mc:Fallback>${sp({ id: 4 })}</mc:Fallback></mc:AlternateContent>` +
      `<mc:AlternateContent><mc:Fallback>${sp({ id: 5 })}</mc:Fallback></mc:AlternateContent>`;
    expect((await shapesOf(tree)).map((s) => s.pptxId)).toEqual(['2', '3', '5']);
  });

  it('reads the static shapes of a layout the same way, alternate content included', async () => {
    const layoutTree = `<mc:AlternateContent><mc:Choice Requires="x">${sp({ id: 3 })}</mc:Choice><mc:Fallback>${sp({ id: 4 })}</mc:Fallback></mc:AlternateContent><p:pic/>${sp({ id: 6 })}`;
    expect((await read({ layoutTree })).layouts[0].elements.map((s) => s.pptxId)).toEqual(['3', '6']);
  });
});

describe('p:sp: geometry (K5)', () => {
  it('reads a preset with its adjustments and says whether the library has it', async () => {
    const g = async (xml: string) => (await one(sp({ spPr: xml }))).geometry;
    expect(await g('<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/><a:gd name="adj2" fmla=" val -5.5 "/><a:gd name="x" fmla="*/ w 1 2"/><a:gd fmla="val 1"/><a:gd name="y"/></a:avLst></a:prstGeom>')).toEqual({
      kind: 'preset', prst: 'roundRect', adjustments: { adj: 16667, adj2: -5.5 }, known: true,
    });
    expect(await g('<a:prstGeom prst="mathDivide"/>')).toEqual({ kind: 'preset', prst: 'mathDivide', known: false });
    expect(await g('<a:custGeom/>')).toEqual({ kind: 'custom' });
    expect(await g('<a:prstGeom/>')).toBeUndefined();
    expect(await g('')).toBeUndefined();
  });
});

describe('p:sp: fills (K1)', () => {
  const fill = async (xml: string, d: Deck = {}) => (await one(sp({ spPr: xml }), d)).fill;

  it('reads none and solid colours with transforms through the colour map of the slide', async () => {
    expect(await fill('<a:noFill/>')).toEqual({ kind: 'none' });
    expect(await fill(`<a:solidFill>${srgb('FF0000', '<a:alpha val="50000"/>')}</a:solidFill>`)).toEqual({ kind: 'solid', color: { kind: 'rgb', hex: '#FF0000', alpha: 0.5 } });
    expect(await fill('<a:solidFill><a:schemeClr val="bg1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>')).toEqual({ kind: 'solid', color: { kind: 'theme', slot: 'lt1', lumMod: 0.75 } });
    expect(await fill('<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>')).toBeUndefined();
    expect(await fill('<a:grpFill/>')).toBeUndefined();
  });

  it('applies the layout colour-map override and then the slide one', async () => {
    const bgFill = '<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>';
    const layoutOvr = '<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>';
    expect(await fill(bgFill, { layoutTail: layoutOvr })).toEqual({ kind: 'solid', color: { kind: 'theme', slot: 'dk1' } });
    const slideOvr = '<p:clrMapOvr><a:overrideClrMapping bg1="accent2"/></p:clrMapOvr>';
    expect(await fill(bgFill, { layoutTail: layoutOvr, sldTail: slideOvr })).toEqual({ kind: 'solid', color: { kind: 'theme', slot: 'accent2' } });
    expect(await fill(bgFill, { sldTail: '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' })).toEqual({ kind: 'solid', color: { kind: 'theme', slot: 'lt1' } });
  });

  it('reads linear, radial and path gradients with their stops', async () => {
    const stops = `<a:gsLst><a:gs pos="100000">${srgb('0000FF')}</a:gs><a:gs pos="0">${srgb('FF0000', '<a:alpha val="40000"/>')}</a:gs><a:gs pos="x">${srgb('00FF00')}</a:gs><a:gs pos="50000"><a:schemeClr val="phClr"/></a:gs></a:gsLst>`;
    expect(await fill(`<a:gradFill rotWithShape="1">${stops}<a:lin ang="5400000" scaled="1"/></a:gradFill>`)).toEqual({
      kind: 'gradient', shape: 'linear', angle: 90, scaled: true, rotWithShape: true,
      stops: [{ pos: 1, color: { kind: 'rgb', hex: '#0000FF' } }, { pos: 0, color: { kind: 'rgb', hex: '#FF0000', alpha: 0.4 } }],
    });
    expect(await fill(`<a:gradFill>${stops}<a:path path="circle"/></a:gradFill>`)).toMatchObject({ shape: 'radial' });
    expect(await fill(`<a:gradFill>${stops}<a:path path="rect"/></a:gradFill>`)).toMatchObject({ shape: 'rect' });
    expect(await fill(`<a:gradFill>${stops}<a:path path="shape"/></a:gradFill>`)).toMatchObject({ shape: 'shape' });
    expect(await fill(`<a:gradFill>${stops}</a:gradFill>`)).toEqual({ kind: 'gradient', shape: 'linear', angle: 0, stops: expect.any(Array) });
    expect(await fill('<a:gradFill><a:gsLst/></a:gradFill>')).toBeUndefined();
  });

  it('reads patterns and picture fills (the image part, tiling, transparency)', async () => {
    expect(await fill(`<a:pattFill prst="dkUpDiag"><a:fgClr>${srgb('112233')}</a:fgClr><a:bgClr>${srgb('FFFFFF')}</a:bgClr></a:pattFill>`)).toEqual({
      kind: 'pattern', prst: 'dkUpDiag', fg: { kind: 'rgb', hex: '#112233' }, bg: { kind: 'rgb', hex: '#FFFFFF' },
    });
    expect(await fill('<a:pattFill/>')).toEqual({ kind: 'pattern', prst: 'pct50' });
    const d = { slideRels: [rel('rId9', 'image', '../media/image1.png'), '<Relationship Id="rId8" Type="x/image" Target="http://example.com/a.png" TargetMode="External"/>'] };
    expect(await fill('<a:blipFill><a:blip r:embed="rId9"><a:alphaModFix amt="25000"/></a:blip><a:tile/></a:blipFill>', d)).toEqual({ kind: 'picture', part: 'ppt/media/image1.png', tile: true, alpha: 0.25 });
    expect(await fill('<a:blipFill><a:blip r:embed="rId8"><a:alphaModFix amt="100000"/></a:blip></a:blipFill>', d)).toEqual({ kind: 'picture' });
    expect(await fill('<a:blipFill><a:blip/></a:blipFill>')).toEqual({ kind: 'picture' });
  });
});

describe('p:sp: outlines (K1)', () => {
  const line = async (xml: string, d: Deck = {}) => (await one(sp({ spPr: xml }), d)).line;

  it('reads width, fill, dash, cap, join and the line ends', async () => {
    expect(
      await line(`<a:ln w="19050" cap="rnd"><a:solidFill>${srgb('00FF00')}</a:solidFill><a:prstDash val="dash"/><a:bevel/><a:headEnd type="triangle" w="lg" len="sm"/><a:tailEnd type="arrow"/></a:ln>`)
    ).toEqual({ width: 2, fill: { kind: 'solid', color: { kind: 'rgb', hex: '#00FF00' } }, dash: 'dash', cap: 'round', join: 'bevel', head: { type: 'triangle', w: 'lg', len: 'sm' }, tail: { type: 'arrow' } });
    expect(await line('<a:ln cap="sq"><a:noFill/><a:custDash/><a:round/><a:headEnd type="none"/><a:miter/></a:ln>')).toEqual({ fill: { kind: 'none' }, dash: 'custom', cap: 'square', join: 'miter' });
    expect(await line('<a:ln w="-1"/>')).toEqual({});
    expect(await line('')).toBeUndefined();
  });
});

describe('p:sp: shape style references', () => {
  const style = (fill: string, ln: string, font = '') => `<p:style><a:lnRef idx="${ln}"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="${fill}"><a:schemeClr val="accent2"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>${font || '<a:fontRef idx="minor"/>'}</p:style>`;

  it('takes a solid fill and an outline from the style when spPr states none, and warns for a gradient style', async () => {
    const r = await read({ tree: sp({ style: style('3', '2') }) });
    const s = (r.slides[0].elements as PptxShape[])[0];
    expect(s.fill).toEqual({ kind: 'solid', color: { kind: 'theme', slot: 'accent2' } });
    expect(s.line).toEqual({ width: 12700 / E, fill: { kind: 'solid', color: { kind: 'theme', slot: 'accent1' } } });
    expect(r.warnings).toContain("A shape takes a gradient fill from the theme's style list; it is read as a solid colour.");
    const plain = await read({ tree: sp({ style: style('1', '7') }) });
    expect((plain.slides[0].elements as PptxShape[])[0].line?.width).toBe(19050 / E);
    expect(plain.warnings.filter((w) => /gradient fill/.test(w))).toEqual([]);
  });

  it('maps the three style line widths to 6350, 12700 and 19050 EMU in px', async () => {
    const widths = await Promise.all(['0', '1', '2', '3'].map(async (ln) => (await one(sp({ style: style('0', ln) }))).line?.width));
    expect(widths).toEqual([undefined, 6350 / E, 12700 / E, 19050 / E]);
  });

  it('reads a grpFill as no fill, so the style fillRef applies (a parent group\'s fill is resolved by the group reader)', async () => {
    const s = await one(sp({ spPr: '<a:grpFill/>', style: style('1', '0') }));
    expect(s.fill).toEqual({ kind: 'solid', color: { kind: 'theme', slot: 'accent2' } });
    expect((await one(sp({ spPr: '<a:grpFill/>' }))).fill).toBeUndefined();
  });

  it('lets spPr win and treats index 0 and a missing or colourless reference as nothing', async () => {
    const own = await one(sp({ spPr: `<a:solidFill>${srgb('112233')}</a:solidFill><a:ln w="9525"><a:noFill/></a:ln>`, style: style('1', '1') }));
    expect(own.fill).toEqual({ kind: 'solid', color: { kind: 'rgb', hex: '#112233' } });
    expect(own.line).toEqual({ width: 1, fill: { kind: 'none' } });
    const keepWidth = await one(sp({ spPr: '<a:ln w="38100"/>', style: style('1', '1') }));
    expect(keepWidth.line).toEqual({ width: 4, fill: { kind: 'solid', color: { kind: 'theme', slot: 'accent1' } } });
    const zero = await one(sp({ style: style('0', '0') }));
    expect(zero.fill).toBeUndefined();
    expect(zero.line).toBeUndefined();
    const bare = await one(sp({ style: '<p:style><a:lnRef idx="1"/><a:fillRef idx="1"><a:schemeClr val="phClr"/></a:fillRef><a:fontRef idx="minor"/><a:effectRef/></p:style>' }));
    expect(bare.fill).toBeUndefined();
    expect(bare.line).toBeUndefined();
    const noIdx = await one(sp({ style: '<p:style><a:lnRef><a:schemeClr val="accent1"/></a:lnRef></p:style>' }));
    expect(noIdx.line).toBeUndefined();
  });

  it('uses the fontRef colour as the text colour below the shape list style', async () => {
    const font = '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>';
    const s = await one(sp({ style: style('1', '1', font), body: txBody(para('x') + para('y'), '', '<a:lstStyle><a:lvl1pPr><a:defRPr sz="1000"/></a:lvl1pPr></a:lstStyle>') }));
    expect(s.text?.paragraphs[0].style.color).toEqual({ kind: 'theme', slot: 'lt1' });
    expect(s.text?.paragraphs[0].style.fontSize).toBe(10);
    const own = await one(sp({ style: style('1', '1', font), body: txBody(para('x'), '', `<a:lstStyle><a:lvl1pPr><a:defRPr>${'<a:solidFill>' + srgb('FF0000') + '</a:solidFill>'}</a:defRPr></a:lvl1pPr></a:lstStyle>`) }));
    expect(own.text?.paragraphs[0].style.color).toEqual({ kind: 'rgb', hex: '#FF0000' });
  });
});

describe('p:sp: text bodies (K3)', () => {
  it('reads the body properties: insets, anchor, vertical text, wrap, columns, rotation and the stored autofit', async () => {
    const body = (attrs: string, inner = '') => `<p:txBody><a:bodyPr${attrs}>${inner}</a:bodyPr><a:lstStyle/><a:p/></p:txBody>`;
    const props = async (b: string) => (await one(sp({ body: b }))).text?.props;
    expect(await props(body(` lIns="${E}" tIns="${2 * E}" rIns="0" bIns="${4 * E}" anchor="ctr" anchorCtr="1" vert="eaVert" wrap="none" numCol="2" spcCol="${10 * E}" rot="1800000"`, '<a:normAutofit fontScale="62500" lnSpcReduction="20000"/>'))).toEqual({
      insets: { l: 1, t: 2, r: 0, b: 4 }, anchor: 'middle', anchorCtr: true, vert: 'eaVert', wrap: 'none', columns: 2, columnSpacing: 10, rot: 30,
      autofit: { kind: 'shrink', fontScale: 0.625, lnSpcReduction: 0.2 },
    });
    expect(await props(body(' anchor="b" anchorCtr="0" wrap="square" numCol="0" spcCol="-1" rot="0" lIns="x"', '<a:noAutofit/>'))).toEqual({ anchor: 'bottom', anchorCtr: false, wrap: 'square', autofit: { kind: 'none' } });
    expect(await props(body('', '<a:spAutoFit/>'))).toEqual({ autofit: { kind: 'resize' } });
    expect(await props(body('', '<a:normAutofit/>'))).toEqual({ autofit: { kind: 'shrink' } });
    expect(await props(body(' wrap="weird"'))).toEqual({});
    expect(await props('<p:txBody><a:lstStyle/></p:txBody>')).toEqual({});
  });

  it('reads runs with every mark, fields and line breaks, in order', async () => {
    const rPr = `<a:rPr lang="de-DE" sz="1800" b="1" i="0" u="sng" strike="dblStrike" baseline="30000" cap="small" spc="150"><a:solidFill>${srgb('FF0000')}</a:solidFill><a:highlight>${srgb('FFFF00')}</a:highlight><a:latin typeface="Georgia"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mn-cs"/></a:rPr>`;
    const s = await one(sp({
      body: txBody(`<a:p>${'<a:r>' + rPr + '<a:t>Hi </a:t></a:r>'}<a:br><a:rPr sz="900"/></a:br><a:fld id="{1}" type="slidenum"><a:rPr b="1"/><a:t>3</a:t></a:fld><a:fld><a:t>d</a:t></a:fld><a:fld/><a:r><a:rPr u="none" strike="sngStrike" baseline="-25000" cap="all" spc="0"><a:latin typeface="+mn-lt"/><a:ea typeface="+x"/></a:rPr></a:r><a:endParaRPr sz="2000"/></a:p>`),
    }));
    const p = s.text?.paragraphs[0];
    expect(p?.runs).toEqual([
      { kind: 'text', text: 'Hi ', lang: 'de-DE', size: 18, bold: true, italic: false, underline: 'sng', strike: 'double', baseline: 0.3, caps: 'small', spacing: 1.5, color: { kind: 'rgb', hex: '#FF0000' }, highlight: { kind: 'rgb', hex: '#FFFF00' }, font: 'Georgia', fontEastAsian: 'major', fontComplex: 'minor' },
      { kind: 'break', text: '', size: 9 },
      { kind: 'field', text: '3', fieldType: 'slidenum', bold: true },
      { kind: 'field', text: 'd' },
      { kind: 'field', text: '' },
      { kind: 'text', text: '', strike: 'single', baseline: -0.25, caps: 'all', font: 'minor' },
    ]);
    expect(p?.endSize).toBe(20);
  });

  it('merges the paragraph style: presentation default, shape list style, the paragraph\'s own properties and its level', async () => {
    const defaultStyle = '<p:defaultTextStyle><a:lvl1pPr marL="0" algn="l"><a:defRPr sz="1800"/></a:lvl1pPr><a:lvl2pPr marL="457200"><a:defRPr sz="1600"/></a:lvl2pPr></p:defaultTextStyle>';
    const lst = '<a:lstStyle><a:lvl2pPr algn="r"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:lvl2pPr></a:lstStyle>';
    const s = await one(
      sp({
        body: txBody(
          para('a') + para('b', '', '<a:pPr lvl="1"><a:spcBef><a:spcPts val="600"/></a:spcBef><a:buChar char="-"/><a:buClr>' + srgb('00FF00') + '</a:buClr><a:buSzPct val="80000"/><a:buFont typeface="Wingdings"/><a:tabLst><a:tab pos="95250" algn="r"/><a:tab pos="190500"/><a:tab algn="l"/></a:tabLst></a:pPr>') +
            para('c', '', '<a:pPr lvl="9" rtl="1"/>') + para('d', '', '<a:pPr lvl="x"/>'),
          '',
          lst
        ),
      }),
      { defaultStyle }
    );
    const ps = s.text?.paragraphs ?? [];
    expect(ps[0]).toMatchObject({ level: 0, style: { fontSize: 18, align: 'left', marginLeft: 0 } });
    expect(ps[1]).toMatchObject({
      level: 1,
      style: { fontSize: 16, marginLeft: 48, align: 'right', lineHeight: 0.9, spaceBefore: 6, bullet: { kind: 'char', char: '-' }, bulletColor: { kind: 'rgb', hex: '#00FF00' }, bulletSizePct: 0.8, bulletFont: 'Wingdings', tabStops: [{ pos: 10, align: 'r' }, { pos: 20, align: 'l' }] },
    });
    expect(ps[2]).toMatchObject({ level: 0, rtl: true });
    expect(ps[3].level).toBe(0);
  });
});

describe('placeholders and layouts', () => {
  const layoutTree =
    sp({ id: 2, name: 'Title', nvPr: '<p:ph type="title"/>', spPr: xfrm(100, 50, 800, 90), body: txBody('<a:p/>', ' lIns="95250" anchor="b"', '<a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4000" b="1"/></a:lvl1pPr></a:lstStyle>') }) +
    sp({ id: 3, name: 'Body', nvPr: '<p:ph idx="1"/>', spPr: xfrm(100, 200, 800, 300), body: txBody('<a:p/>', '', '<a:lstStyle><a:lvl1pPr marL="342900"><a:buChar char="•"/></a:lvl1pPr></a:lstStyle>') });
  const masterTree = sp({ id: 2, name: 'MT', nvPr: '<p:ph type="title"/>', body: txBody('<a:p/>', ' tIns="0" anchor="ctr"') });

  it('takes position, text style and body properties from the layout placeholder, and keeps what the slide states', async () => {
    const tree =
      sp({ id: 4, nvPr: '<p:ph type="title"/>', body: txBody(para('Hello')) }) +
      sp({ id: 5, nvPr: '<p:ph idx="1"/>', spPr: xfrm(1, 2, 3, 4), body: txBody(para('Body', '', '<a:pPr algn="l"/>'), ' anchor="t"') }) +
      sp({ id: 6, nvPr: '<p:ph type="dt" idx="10"/>' });
    const els = await shapesOf(tree, { layoutTree, masterTree });
    expect(els[0].placeholder).toEqual({ type: 'title' });
    expect(els[0].frame).toEqual({ x: 100, y: 50, w: 800, h: 90 });
    expect(els[0].text?.props).toEqual({ insets: { l: 10, t: 0 }, anchor: 'bottom' });
    expect(els[0].text?.paragraphs[0].style).toMatchObject({ align: 'center', fontSize: 40, bold: true });
    expect(els[1].placeholder).toEqual({ type: 'obj', idx: 1 });
    expect(els[1].frame).toEqual({ x: 1, y: 2, w: 3, h: 4 });
    expect(els[1].text?.props).toEqual({ anchor: 'top' });
    expect(els[1].text?.paragraphs[0].style).toMatchObject({ align: 'left', marginLeft: 36, bullet: { kind: 'char', char: '•' } });
    expect(els[2].frame).toBeUndefined();
  });

  it('falls back to the presentation default for a placeholder the layout lacks, or a slide without a layout', async () => {
    const defaultStyle = '<p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr></p:defaultTextStyle>';
    const s = await one(sp({ nvPr: '<p:ph type="tbl" idx="9"/>', body: txBody(para('x')) }), { layoutTree, defaultStyle });
    expect(s.frame).toBeUndefined();
    expect(s.text?.paragraphs[0].style.fontSize).toBe(14);
  });

  it('exposes the static content of layout and master in drawing order, without placeholders, and hides the master\'s when the layout says so', async () => {
    const m = sp({ id: 10, name: 'Logo', spPr: xfrm(1, 1, 10, 10) + '<a:prstGeom prst="rect"/>' }) + masterTree + '<p:pic/>';
    const l = sp({ id: 11, name: 'Rule', spPr: xfrm(0, 0, 100, 2) }) + layoutTree + '<p:cxnSp/>';
    const r = await read({ masterTree: m, layoutTree: l });
    expect(r.layouts[0].elements.map((e) => e.name)).toEqual(['Logo', 'Rule']);
    expect(r.layouts[0].staticShapes).toBe(4);
    const hidden = await read({ masterTree: m, layoutTree: l, layoutAttrs: ' showMasterSp="0"' });
    expect(hidden.layouts[0].elements.map((e) => e.name)).toEqual(['Rule']);
    const field = await read({ layoutTree: sp({ id: 12, name: 'Num', body: txBody('<a:p><a:fld type="slidenum"><a:t>‹#›</a:t></a:fld></a:p>') }) });
    expect(field.layouts[0].elements[0].text?.paragraphs[0].runs[0]).toEqual({ kind: 'field', text: '‹#›', fieldType: 'slidenum' });
  });

  it('inherits the master default level for static layout text and the master colour map for master shapes', async () => {
    const defaultStyle = '<p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="2000"/></a:lvl1pPr></p:defaultTextStyle>';
    const r = await read({ defaultStyle, layoutTree: sp({ body: txBody(para('x')) }) });
    expect(r.layouts[0].elements[0].text?.paragraphs[0].style.fontSize).toBe(20);
  });
});

describe('slides without a usable layout', () => {
  const bgFill = '<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>';
  it('read their shapes with the default colour map', async () => {
    const s = await one(sp({ spPr: bgFill, nvPr: '<p:ph type="title"/>' }), { noLayoutRel: true });
    expect(s.fill).toEqual({ kind: 'solid', color: { kind: 'theme', slot: 'lt1' } });
    expect(s.frame).toBeUndefined();
  });
  it('read a layout whose master is unknown with the default colour map', async () => {
    expect((await one(sp({ spPr: bgFill }), { layoutWithoutMaster: true })).fill).toEqual({ kind: 'solid', color: { kind: 'theme', slot: 'lt1' } });
  });
});

describe('slide background', () => {
  it('reads a background fill, a style reference colour, and none', async () => {
    const bg = (inner: string) => `<p:bg>${inner}</p:bg>`;
    const of = async (b: string) => (await read({ bg: b })).slides[0].background;
    expect(await of(bg(`<p:bgPr><a:solidFill>${srgb('112233')}</a:solidFill></p:bgPr>`))).toEqual({ kind: 'solid', color: { kind: 'rgb', hex: '#112233' } });
    expect(await of(bg(`<p:bgPr><a:gradFill><a:gsLst><a:gs pos="0">${srgb('000000')}</a:gs></a:gsLst></a:gradFill></p:bgPr>`))).toMatchObject({ kind: 'gradient' });
    expect(await of(bg('<p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>'))).toEqual({ kind: 'solid', color: { kind: 'theme', slot: 'lt1' } });
    expect(await of(bg('<p:bgRef idx="1001"/>'))).toBeUndefined();
    expect(await of('')).toBeUndefined();
  });
});

describe('fonts', () => {
  it('lists the typefaces the deck names once per family and reports the ones outside the available list', async () => {
    const run = (face: string) => para('x', `<a:rPr><a:latin typeface="${face}"/></a:rPr>`);
    const r = await read({
      tree: sp({ body: txBody(run('Segoe UI') + run('segoe ui') + run('Arial') + run('+mn-lt')) }),
      layoutTree: sp({ body: txBody(run('Comic Sans MS')) }),
      masterTree: sp({ body: txBody(run('Calibri')) }),
    });
    expect(r.fonts.used).toEqual(['Arial', 'Calibri', 'Calibri Light', 'Comic Sans MS', 'Segoe UI']);
    expect(r.fonts.unavailable).toEqual(['Comic Sans MS', 'Segoe UI']);
    expect(r.warnings).toContain('2 fonts the editor may not have (Comic Sans MS, Segoe UI) are used; text in them shows in a substitute and may reflow.');
    const one = await read({ tree: sp({ body: txBody(run('Segoe UI')) }) });
    expect(one.warnings).toContain('A font the editor may not have (Segoe UI) is used; text in it shows in a substitute and may reflow.');
    const none = await read({ tree: sp({ body: txBody(run('Arial')) }) });
    expect(none.warnings).toEqual([]);
    expect(none.fonts).toEqual({ used: ['Arial', 'Calibri', 'Calibri Light'], unavailable: [] });
  });
});

describe('a slide that cannot be read', () => {
  it('degrades to an error placeholder with a warning and leaves the others alone', async () => {
    const spy = jest.spyOn(shapesModule, 'readShape');
    // the layouts' static content goes through `readShapes`: only the slide's own shapes fail
    const failing = (error: unknown) => () => {
      throw error;
    };
    spy.mockImplementation(failing(new Error('boom')));
    const r = await read({ tree: sp() });
    const s: PptxSlide = r.slides[0];
    expect(s.error).toBe('the content could not be read');
    expect(s.elements).toBeUndefined();
    expect(r.warnings).toContain('Slide 1 could not be read (boom); it is kept as an empty placeholder.');
    spy.mockImplementation(failing('not an error'));
    expect((await read({ tree: sp() })).warnings).toContain('Slide 1 could not be read (unknown error); it is kept as an empty placeholder.');
    spy.mockRestore();
  });
});
