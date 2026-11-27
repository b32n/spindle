import JSZip from 'jszip';
import type { PptxLayout } from '@b32nio/spindle-shared/export/pptx';
import { importPptx } from './reader';
import { matchLayoutPlaceholder, mergeLevels, parseLevel, parseLevels } from './layouts';
import { DEFAULT_COLOR_MAP, overrideColorMap } from './color';
import { parseXml } from '../xlsx/xml';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PRES_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
const CLR_MAP = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';

const rel = (id: string, type: string, target: string): string => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`;
const rels = (list: string[]): string => `<Relationships xmlns="${PKG_REL}">${list.join('')}</Relationships>`;
const theme = (accent1 = '4472C4'): string =>
  `<a:theme ${NS} name="T"><a:themeElements><a:clrScheme name="c"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="${accent1}"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Georgia"/></a:majorFont><a:minorFont><a:latin typeface="Verdana"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;

const xfrm = (x: number, y: number, w: number, h: number, attrs = ''): string =>
  `<a:xfrm${attrs}><a:off x="${x * 9525}" y="${y * 9525}"/><a:ext cx="${w * 9525}" cy="${h * 9525}"/></a:xfrm>`;
const sp = (ph: string, opts: { name?: string; xfrm?: string; body?: string; lst?: string; text?: string[] } = {}): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="${opts.name ?? ''}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr>${opts.xfrm ?? ''}</p:spPr><p:txBody><a:bodyPr${opts.body ?? ''}/>${opts.lst ?? '<a:lstStyle/>'}${(opts.text ?? []).map((t) => `<a:p><a:r><a:t>${t}</a:t></a:r></a:p>`).join('')}</p:txBody></p:sp>`;
const shape = '<p:sp><p:nvSpPr><p:cNvPr id="9" name="Bar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>';
const tree = (...shapes: string[]): string => `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${shapes.join('')}</p:spTree>`;

interface LayoutSpec {
  xml?: string;
  /** Raw part content instead of `xml`. */
  raw?: string;
}
interface MasterSpec {
  xml?: string;
  layouts?: LayoutSpec[];
  /** List layouts in `p:sldLayoutIdLst` in this order of indexes (default: relationship order). */
  order?: number[];
  raw?: string;
  accent?: string;
}

const master = (cSld: string, body = '', attrs = ''): string =>
  `<p:sldMaster ${NS}>${cSld}${CLR_MAP.replace('<p:clrMap', `<p:clrMap${attrs}`)}${body}</p:sldMaster>`;
const layout = (cSld: string, attrs = '', tail = ''): string => `<p:sldLayout ${NS}${attrs}>${cSld}${tail}</p:sldLayout>`;

async function build(masters: MasterSpec[], opts: { defaultStyle?: string; slideLayouts?: string[]; unlisted?: boolean; extraMasterIds?: string } = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="${CT}"><Override PartName="/ppt/presentation.xml" ContentType="${PRES_TYPE}"/></Types>`);
  zip.file('_rels/.rels', rels([rel('rId1', 'officeDocument', 'ppt/presentation.xml')]));
  const presRels: string[] = [];
  let layoutNo = 0;
  masters.forEach((m, i) => {
    const n = i + 1;
    if (!opts.unlisted) presRels.push(rel(`rIdM${n}`, 'slideMaster', `slideMasters/slideMaster${n}.xml`));
    const layouts = m.layouts ?? [{}];
    const parts: string[] = [];
    const mrels: string[] = [rel('rIdT', 'theme', `../theme/theme${n}.xml`)];
    layouts.forEach((l, j) => {
      layoutNo++;
      parts.push(`ppt/slideLayouts/slideLayout${layoutNo}.xml`);
      mrels.push(rel(`rIdL${j + 1}`, 'slideLayout', `../slideLayouts/slideLayout${layoutNo}.xml`));
      zip.file(`ppt/slideLayouts/slideLayout${layoutNo}.xml`, l.raw ?? l.xml ?? layout('<p:cSld/>'));
      zip.file(`ppt/slideLayouts/_rels/slideLayout${layoutNo}.xml.rels`, rels([rel('rId1', 'slideMaster', `../slideMasters/slideMaster${n}.xml`)]));
    });
    const list = m.order ? `<p:sldLayoutIdLst>${m.order.map((k) => `<p:sldLayoutId id="${2147483700 + k}" r:id="rIdL${k + 1}"/>`).join('')}</p:sldLayoutIdLst>` : '';
    const base = m.xml ?? master('<p:cSld/>');
    zip.file(`ppt/slideMasters/slideMaster${n}.xml`, m.raw ?? base.replace('</p:sldMaster>', `${list}</p:sldMaster>`));
    zip.file(`ppt/slideMasters/_rels/slideMaster${n}.xml.rels`, rels(mrels));
    zip.file(`ppt/theme/theme${n}.xml`, theme(m.accent));
  });
  const slideLayouts = opts.slideLayouts ?? [];
  const sldIds = slideLayouts.map((_, i) => `<p:sldId id="${256 + i}" r:id="rIdS${i + 1}"/>`).join('');
  slideLayouts.forEach((target, i) => {
    presRels.push(rel(`rIdS${i + 1}`, 'slide', `slides/slide${i + 1}.xml`));
    zip.file(`ppt/slides/slide${i + 1}.xml`, `<p:sld ${NS}><p:cSld><p:spTree/></p:cSld></p:sld>`);
    zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, rels([rel('rId1', 'slideLayout', target)]));
  });
  const masterIds = masters.map((_, i) => `<p:sldMasterId id="${2147483648 + i}" r:id="rIdM${i + 1}"/>`).join('');
  zip.file('ppt/_rels/presentation.xml.rels', rels(presRels));
  zip.file(
    'ppt/presentation.xml',
    `<p:presentation ${NS}>${opts.unlisted ? '' : `<p:sldMasterIdLst>${masterIds}${opts.extraMasterIds ?? ''}</p:sldMasterIdLst>`}<p:sldIdLst>${sldIds}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/>${opts.defaultStyle ?? ''}</p:presentation>`
  );
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const layoutsOf = async (masters: MasterSpec[], opts?: Parameters<typeof build>[1]): Promise<{ layouts: PptxLayout[]; warnings: string[] }> => {
  const r = await importPptx(await build(masters, opts));
  return { layouts: r.layouts, warnings: r.warnings };
};
const one = async (layoutXml: string, masterXml = master('<p:cSld/>'), defaultStyle?: string) => {
  const r = await layoutsOf([{ xml: masterXml, layouts: [{ xml: layoutXml }] }], { defaultStyle });
  return { layout: r.layouts[0], warnings: r.warnings };
};
const lvl = (n: number, attrs: string, inner = ''): string => `<a:lvl${n}pPr${attrs}>${inner}</a:lvl${n}pPr>`;
const defRPr = (attrs: string, inner = ''): string => `<a:defRPr${attrs}>${inner}</a:defRPr>`;
const parse = (xml: string) => parseXml(`<a:lstStyle ${NS}>${xml}</a:lstStyle>`) ?? undefined;

describe('level styles', () => {
  it('reads every property of a level and nothing it does not state', () => {
    const style = parseLevel(
      parse(lvl(1, ' marL="571500" indent="-285750" algn="ctr"', `<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft><a:buChar char="•"/>${defRPr(' sz="2800" b="1" i="0"', '<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:latin typeface="+mj-lt"/>')}`))?.children[0],
      DEFAULT_COLOR_MAP
    );
    expect(style).toEqual({
      fontSize: 28, bold: true, italic: false, color: { kind: 'theme', slot: 'dk1', lumMod: 0.75 }, font: 'major', align: 'center',
      marginLeft: 60, indent: -30, lineHeight: 0.9, spaceBefore: 10, spaceAfter: 6, bullet: { kind: 'char', char: '•' },
    });
    expect(parseLevel(undefined, DEFAULT_COLOR_MAP)).toEqual({});
    expect(parseLevel(parse(lvl(1, ''))?.children[0], DEFAULT_COLOR_MAP)).toEqual({});
  });

  it('reads percentage paragraph spacing and exact line spacing', () => {
    const at = (inner: string) => parseLevel(parse(lvl(1, '', inner))?.children[0], DEFAULT_COLOR_MAP);
    expect(at('<a:spcBef><a:spcPct val="20000"/></a:spcBef><a:spcAft><a:spcPct val="10%"/></a:spcAft>')).toEqual({ spaceBeforePct: 0.2, spaceAfterPct: 0.1 });
    expect(at('<a:lnSpc><a:spcPts val="1800"/></a:lnSpc>')).toEqual({ lineSpacingPts: 18 });
    expect(at('<a:lnSpc><a:spcPct val="80000"/><a:spcPts val="1800"/></a:lnSpc>')).toEqual({ lineHeight: 0.8 });
    expect(at('<a:lnSpc><a:spcPts val="0"/></a:lnSpc>')).toEqual({});
    expect(at('<a:spcBef><a:spcPts val="600"/><a:spcPct val="20000"/></a:spcBef>')).toEqual({ spaceBefore: 6 });
    expect(at('<a:spcBef><a:spcPct val="-1"/></a:spcBef><a:spcAft><a:spcPct val="x"/></a:spcAft>')).toEqual({});
  });

  it('lets a level that states spacing in one unit replace the other unit', () => {
    const lv = (o: object) => Array.from({ length: 9 }, (_, i) => (i === 0 ? o : {}));
    const base = lv({ spaceBefore: 6, spaceAfterPct: 0.1, lineHeight: 0.9, fontSize: 20 });
    expect(mergeLevels(base, lv({ spaceBeforePct: 0.2, spaceAfter: 3, lineSpacingPts: 24 }))[0]).toEqual({
      fontSize: 20, spaceBeforePct: 0.2, spaceAfter: 3, lineSpacingPts: 24,
    });
    expect(mergeLevels(base, lv({ fontSize: 10 }))[0]).toEqual({ ...base[0], fontSize: 10 });
  });

  it.each([
    ['<a:buNone/>', { kind: 'none' }],
    ['<a:buChar char=""/>', undefined],
    ['<a:buAutoNum type="romanUcPeriod"/>', { kind: 'auto', scheme: 'romanUcPeriod' }],
    ['<a:buAutoNum type="arabicPeriod" startAt="3"/>', { kind: 'auto', scheme: 'arabicPeriod', startAt: 3 }],
    ['<a:buAutoNum startAt="1"/>', { kind: 'auto', scheme: 'arabicPeriod' }],
  ])('reads the bullet %s', (xml, bullet) => {
    expect(parseLevel(parse(lvl(1, '', xml))?.children[0], DEFAULT_COLOR_MAP).bullet).toEqual(bullet);
  });

  it('maps alignments, theme fonts, literal fonts, and ignores unusable values', () => {
    const at = (attrs: string, inner = '') => parseLevel(parse(lvl(1, attrs, inner))?.children[0], DEFAULT_COLOR_MAP);
    expect(at(' algn="l"').align).toBe('left');
    expect(at(' algn="r"').align).toBe('right');
    expect(at(' algn="dist"').align).toBe('justify');
    expect(at(' algn="bogus"').align).toBeUndefined();
    expect(at('', defRPr('', '<a:latin typeface="+mn-lt"/>')).font).toBe('minor');
    expect(at('', defRPr('', '<a:latin typeface="Georgia"/>')).font).toBe('Georgia');
    expect(at('', defRPr('', '<a:latin typeface="+mn-ea"/>')).font).toBeUndefined();
    expect(at('', defRPr('', '<a:latin typeface=" "/>')).font).toBeUndefined();
    expect(at('', defRPr(' sz="0"')).fontSize).toBeUndefined();
    expect(at('', defRPr(' sz="x"')).fontSize).toBeUndefined();
    expect(at('', '<a:lnSpc><a:spcPct val="0"/></a:lnSpc>').lineHeight).toBeUndefined();
    expect(at('', '<a:spcBef><a:spcPts val="-5"/></a:spcBef><a:spcAft><a:spcPts val="-5"/></a:spcAft>')).toEqual({});
    expect(at(' marL="x"')).toEqual({});
    expect(at('', defRPr(' b="true"')).bold).toBe(true);
  });

  it('parses nine levels and merges level by level', () => {
    const base = parseLevels(parse(lvl(1, '', defRPr(' sz="1800"')) + lvl(2, '', defRPr(' sz="1600" b="1"'))), DEFAULT_COLOR_MAP);
    expect(base).toHaveLength(9);
    const over = parseLevels(parse(lvl(2, '', defRPr(' sz="1400"'))), DEFAULT_COLOR_MAP);
    const merged = mergeLevels(base, over);
    expect(merged[0]).toEqual({ fontSize: 18 });
    expect(merged[1]).toEqual({ fontSize: 14, bold: true });
    expect(merged[8]).toEqual({});
    expect(mergeLevels([], [])).toHaveLength(9);
    expect(parseLevels(undefined, DEFAULT_COLOR_MAP).every((l) => Object.keys(l).length === 0)).toBe(true);
  });
});

describe('colour map override', () => {
  it('replaces only the roles it names with valid slots', () => {
    const el = parseXml(`<a:overrideClrMapping ${NS} bg1="dk1" tx1="lt1" bg2="nonsense"/>`) ?? undefined;
    const map = overrideColorMap(DEFAULT_COLOR_MAP, el);
    expect(map.bg1).toBe('dk1');
    expect(map.tx1).toBe('lt1');
    expect(map.bg2).toBe('lt2');
    expect(overrideColorMap(DEFAULT_COLOR_MAP, undefined)).toEqual(DEFAULT_COLOR_MAP);
  });
});

const titleStyle = `<p:txStyles><p:titleStyle>${lvl(1, ' algn="l"', defRPr(' sz="4400"', '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/>'))}</p:titleStyle><p:bodyStyle>${lvl(1, ' marL="228600" indent="-228600"', defRPr(' sz="2800"', '<a:latin typeface="+mn-lt"/>') + '<a:buChar char="•"/>')}${lvl(2, ' marL="685800"', defRPr(' sz="2400"'))}</p:bodyStyle><p:otherStyle>${lvl(1, ' algn="r"', defRPr(' sz="1200"'))}</p:otherStyle></p:txStyles>`;
const masterTree = tree(
  sp('<p:ph type="title"/>', { name: 'Title', xfrm: xfrm(10, 20, 300, 80), body: ' anchor="ctr"' }),
  sp('<p:ph type="body" idx="1"/>', { name: 'Text', xfrm: xfrm(10, 120, 300, 200), lst: `<a:lstStyle>${lvl(1, '', defRPr(' b="1"'))}</a:lstStyle>` }),
  sp('<p:ph type="dt" idx="2"/>', { xfrm: xfrm(10, 500, 100, 30) }),
  sp('<p:ph type="sldNum" idx="4"/>', { xfrm: xfrm(200, 500, 100, 30) }),
  shape
);
const fullMaster = master(`<p:cSld name="Corporate"><p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></p:bgPr></p:bg>${masterTree}</p:cSld>`, titleStyle);

describe('layout flattening', () => {
  it('inherits frame, anchor, text style levels and the master name from the master', async () => {
    const { layout: l, warnings } = await one(
      layout(`<p:cSld name="Title and Content">${tree(sp('<p:ph type="title"/>'), sp('<p:ph idx="1"/>', { name: 'Content', text: ['Click', 'to edit'] }))}</p:cSld>`, ' type="obj"'),
      fullMaster
    );
    expect(warnings).toEqual([]);
    expect(l).toMatchObject({ name: 'Title and Content', layoutType: 'obj', masterName: 'Corporate', masterPart: 'ppt/slideMasters/slideMaster1.xml', background: { kind: 'rgb', hex: '#112233' }, staticShapes: 1 });
    const [title, body] = l.placeholders;
    expect(title).toMatchObject({ type: 'title', frame: { x: 10, y: 20, w: 300, h: 80 }, anchor: 'middle' });
    expect(title.idx).toBeUndefined();
    expect(title.levels[0]).toEqual({ fontSize: 44, color: { kind: 'theme', slot: 'dk1' }, font: 'major', align: 'left' });
    expect(body).toMatchObject({ type: 'obj', idx: 1, name: 'Content', prompt: 'Click\nto edit', frame: { x: 10, y: 120, w: 300, h: 200 } });
    // body style level 1 (master) + the master placeholder's lstStyle (bold)
    expect(body.levels[0]).toEqual({ fontSize: 28, bold: true, font: 'minor', marginLeft: 24, indent: -24, bullet: { kind: 'char', char: '•' } });
  });

  it('is the master placeholder lstStyle, then the layout placeholder lstStyle, that override the master text style', async () => {
    const { layout: l } = await one(
      layout(`<p:cSld name="L">${tree(sp('<p:ph type="body" idx="1"/>', { lst: `<a:lstStyle>${lvl(1, '', defRPr(' sz="2000" i="1"'))}${lvl(3, '', defRPr(' sz="1000"'))}</a:lstStyle>` }))}</p:cSld>`),
      fullMaster,
      `<p:defaultTextStyle>${lvl(1, ' algn="r"', defRPr(' sz="1800"', '<a:latin typeface="Arial"/>'))}${lvl(4, '', defRPr(' sz="900"'))}</p:defaultTextStyle>`
    );
    const levels = l.placeholders[0].levels;
    // master body style (28) < master placeholder (b="1" only) < layout (20 italic); default text style gives align and font below the master's font
    expect(levels[0]).toEqual({ fontSize: 20, italic: true, align: 'right', font: 'minor', marginLeft: 24, indent: -24, bullet: { kind: 'char', char: '•' }, bold: true });
    expect(levels[1]).toEqual({ fontSize: 24, marginLeft: 72 });
    expect(levels[2]).toEqual({ fontSize: 10 });
    expect(levels[3]).toEqual({ fontSize: 9 });
  });

  it('uses the other style for dates, footers and slide numbers, matched by type, and the layout position when it has one', async () => {
    const { layout: l } = await one(
      layout(`<p:cSld name="L">${tree(sp('<p:ph type="dt" idx="10"/>'), sp('<p:ph type="sldNum" idx="12"/>', { xfrm: xfrm(1, 2, 3, 4, ' rot="5400000"') }), sp('<p:ph type="ftr" idx="11"/>'))}</p:cSld>`),
      fullMaster
    );
    const [dt, num, ftr] = l.placeholders;
    expect(dt).toMatchObject({ type: 'dt', idx: 10, frame: { x: 10, y: 500, w: 100, h: 30 } });
    expect(dt.levels[0]).toEqual({ fontSize: 12, align: 'right' });
    expect(num.frame).toEqual({ x: 1, y: 2, w: 3, h: 4, rot: 90 });
    // the master has no footer placeholder: no frame to inherit, but the other style still applies
    expect(ftr.frame).toBeUndefined();
    expect(ftr.levels[0].fontSize).toBe(12);
  });

  it('matches both title types to the master title, and other body-like types to the master body', async () => {
    const { layout: l } = await one(
      layout(`<p:cSld name="L">${tree(sp('<p:ph type="ctrTitle"/>'), sp('<p:ph type="subTitle" idx="1"/>'), sp('<p:ph type="pic" idx="13"/>'), sp('<p:ph type="chart" idx="14"/>', { body: ' anchor="b"' }))}</p:cSld>`),
      fullMaster
    );
    expect(l.placeholders.map((p) => p.frame?.y)).toEqual([20, 120, 120, 120]);
    expect(l.placeholders.map((p) => p.anchor)).toEqual(['middle', undefined, undefined, 'bottom']);
    // the master with only a ctrTitle also matches a title
    const m = master(`<p:cSld>${tree(sp('<p:ph type="ctrTitle"/>', { xfrm: xfrm(1, 1, 5, 5) }))}</p:cSld>`);
    expect((await one(layout(`<p:cSld>${tree(sp('<p:ph type="title"/>'))}</p:cSld>`), m)).layout.placeholders[0].frame).toEqual({ x: 1, y: 1, w: 5, h: 5 });
  });

  it('reads pictures and graphic frames as placeholders, counts other shapes, and handles a missing or invalid frame', async () => {
    const pic = (ph: string, x = '') => `<p:pic><p:nvPicPr><p:cNvPr id="3" name="Pic"/><p:cNvPicPr/><p:nvPr>${ph}</p:nvPr></p:nvPicPr><p:spPr>${x}</p:spPr></p:pic>`;
    const gf = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Tbl"/><p:cNvGraphicFramePr/><p:nvPr><p:ph type="tbl" idx="15"/></p:nvPr></p:nvGraphicFramePr><p:xfrm><a:off x="${9525}" y="${9525}"/><a:ext cx="${19050}" cy="${19050}"/></p:xfrm></p:graphicFrame>`;
    const { layout: l } = await one(
      layout(`<p:cSld name="L">${tree(pic('<p:ph type="pic" idx="1"/>', xfrm(5, 6, 7, 8)), gf, pic(''), '<p:grpSp/>', '<p:cxnSp/>', '<p:sp><p:nvSpPr><p:nvPr><p:ph idx="6"/></p:nvPr></p:nvSpPr></p:sp>', '<p:extLst/>', sp('<p:ph idx="3"/>', { xfrm: '<a:xfrm><a:off x="0" y="0"/></a:xfrm>' }), sp('<p:ph idx="4"/>', { xfrm: '<a:xfrm><a:off x="0" y="0"/><a:ext cx="-1" cy="5"/></a:xfrm>' }))}</p:cSld>`)
    );
    expect(l.placeholders.map((p) => p.type)).toEqual(['pic', 'tbl', 'obj', 'obj', 'obj']);
    expect(l.placeholders[0].frame).toEqual({ x: 5, y: 6, w: 7, h: 8 });
    expect(l.placeholders[1].frame).toEqual({ x: 1, y: 1, w: 2, h: 2 });
    expect(l.placeholders[2].frame).toBeUndefined();
    expect(l.placeholders[3].frame).toBeUndefined();
    expect(l.staticShapes).toBe(3);
  });

  it('counts the master shapes unless the layout hides them, and names a layout after its file when it has no name', async () => {
    const shown = await one(layout('<p:cSld>' + tree(shape) + '</p:cSld>'), fullMaster);
    expect(shown.layout.name).toBe('slideLayout1');
    expect(shown.layout.staticShapes).toBe(2);
    expect(shown.layout.hideMasterShapes).toBeUndefined();
    const hidden = await one(layout('<p:cSld name="  ">' + tree(shape) + '</p:cSld>', ' showMasterSp="0"'), fullMaster);
    expect(hidden.layout).toMatchObject({ name: 'slideLayout1', staticShapes: 1, hideMasterShapes: true });
    expect((await one(layout('<p:cSld/>', ' showMasterSp="false"'), fullMaster)).layout.hideMasterShapes).toBe(true);
    expect((await one(layout('<p:cSld/>'), master('<p:cSld/>'))).layout.masterName).toBeUndefined();
  });
});

describe('layout backgrounds', () => {
  const bgOf = async (layoutBg: string, masterBg = '') => {
    const r = await one(layout(`<p:cSld name="L">${layoutBg}</p:cSld>`), master(`<p:cSld>${masterBg}</p:cSld>`));
    return { bg: r.layout.background, warnings: r.warnings };
  };
  const solid = (hex: string) => `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></p:bgPr></p:bg>`;

  it('takes the layout background, else the master background, else none', async () => {
    expect((await bgOf(solid('AA0000'), solid('00AA00'))).bg).toEqual({ kind: 'rgb', hex: '#AA0000' });
    expect((await bgOf('', solid('00AA00'))).bg).toEqual({ kind: 'rgb', hex: '#00AA00' });
    expect((await bgOf('')).bg).toBeUndefined();
    expect((await bgOf('<p:bg><p:bgPr><a:noFill/></p:bgPr></p:bg>')).bg).toBeUndefined();
  });

  it('reads a background reference through the colour map (bg1 -> lt1) with no warning for the plain style', async () => {
    const r = await bgOf('', '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg2"/></p:bgRef></p:bg>');
    expect(r.bg).toEqual({ kind: 'theme', slot: 'lt2' });
    expect(r.warnings).toEqual([]);
  });

  it('approximates gradient, picture and theme-style backgrounds and says so', async () => {
    const grad = await bgOf('<p:bg><p:bgPr><a:gradFill/></p:bgPr></p:bg>');
    expect(grad.bg).toBeUndefined();
    expect(grad.warnings).toEqual(['L: the background is a grad fill; it is read as a plain background.']);
    const style = await bgOf('<p:bg><p:bgRef idx="1003"><a:srgbClr val="123456"/></p:bgRef></p:bg>');
    expect(style.bg).toEqual({ kind: 'rgb', hex: '#123456' });
    expect(style.warnings).toEqual(['L: the background is a theme style fill; it is read as a solid colour.']);
    const noColor = await bgOf('<p:bg><p:bgRef idx="1003"/></p:bg>');
    expect(noColor.bg).toBeUndefined();
    expect(noColor.warnings).toEqual([]);
  });

  it('warns once per background source, not once per layout inheriting it', async () => {
    const gradMaster = master('<p:cSld name="Grad"><p:bg><p:bgPr><a:gradFill/></p:bgPr></p:bg></p:cSld>');
    const r = await layoutsOf([{ xml: gradMaster, layouts: [{ xml: layout('<p:cSld name="A"/>') }, { xml: layout('<p:cSld name="B"/>') }, { xml: layout('<p:cSld name="C"><p:bg><p:bgPr><a:pattFill/></p:bgPr></p:bg></p:cSld>') }] }]);
    expect(r.layouts).toHaveLength(3);
    expect(r.warnings).toEqual([
      'Grad: the background is a grad fill; it is read as a plain background.',
      'C: the background is a patt fill; it is read as a plain background.',
    ]);
    const unnamed = await layoutsOf([{ xml: master('<p:cSld><p:bg><p:bgPr><a:blipFill/></p:bgPr></p:bg></p:cSld>'), layouts: [{ xml: layout('<p:cSld name="A"/>') }] }]);
    expect(unnamed.warnings).toEqual(['slideMaster1.xml: the background is a blip fill; it is read as a plain background.']);
  });

  it('applies the layout colour map override to its own background and text colours', async () => {
    const r = await one(
      layout(`<p:cSld name="Dark"><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>${tree(sp('<p:ph type="title"/>'))}</p:cSld>`, '', '<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>'),
      fullMaster
    );
    expect(r.layout.background).toEqual({ kind: 'theme', slot: 'dk1' });
    expect(r.layout.placeholders[0].levels[0].color).toEqual({ kind: 'theme', slot: 'lt1' });
    const plain = await one(layout('<p:cSld/>', '', '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'), master('<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg></p:cSld>'));
    expect(plain.layout.background).toEqual({ kind: 'theme', slot: 'lt1' });
  });
});

describe('masters and layouts of a deck', () => {
  it('lists layouts per master in the order of the layout id list, masters in presentation order', async () => {
    const r = await layoutsOf([
      { layouts: [{ xml: layout('<p:cSld name="A"/>') }, { xml: layout('<p:cSld name="B"/>') }, { xml: layout('<p:cSld name="C"/>') }], order: [2, 0, 1] },
      { xml: master('<p:cSld name="M2"/>'), layouts: [{ xml: layout('<p:cSld name="D"/>') }] },
    ]);
    expect(r.layouts.map((l) => [l.name, l.masterPart])).toEqual([
      ['C', 'ppt/slideMasters/slideMaster1.xml'], ['A', 'ppt/slideMasters/slideMaster1.xml'], ['B', 'ppt/slideMasters/slideMaster1.xml'],
      ['D', 'ppt/slideMasters/slideMaster2.xml'],
    ]);
    expect(r.layouts[3].masterName).toBe('M2');
  });

  it('ignores a layout id list entry that points nowhere and falls back to the relationships when none is usable', async () => {
    const xml = master('<p:cSld/>').replace('</p:sldMaster>', '<p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rIdNope"/></p:sldLayoutIdLst></p:sldMaster>');
    const r = await layoutsOf([{ xml, layouts: [{ xml: layout('<p:cSld name="A"/>') }, { xml: layout('<p:cSld name="B"/>') }] }]);
    expect(r.layouts.map((l) => l.name)).toEqual(['A', 'B']);
  });

  it('warns once about a second master with a different theme and not about an identical one', async () => {
    const different = await layoutsOf([{}, { accent: 'AA0000' }]);
    expect(different.warnings).toEqual(['slideMaster2.xml uses a different theme from the first master; the deck keeps one theme, so its colours and fonts follow the first.']);
    expect((await layoutsOf([{}, {}])).warnings).toEqual([]);
  });

  it('skips a layout or a master that is not the right part, with a warning', async () => {
    const r = await layoutsOf([
      { layouts: [{ raw: 'not xml at all' }, { raw: `<p:sld ${NS}/>` }, { xml: layout('<p:cSld name="Good"/>') }] },
      { raw: '<p:other/>' },
    ]);
    expect(r.layouts.map((l) => l.name)).toEqual(['Good']);
    expect(r.warnings).toEqual([
      'slideLayout1.xml could not be read as a slide layout; it is skipped.',
      'slideLayout2.xml could not be read as a slide layout; it is skipped.',
      'slideMaster2.xml could not be read as a slide master; its layouts are skipped.',
    ]);
  });

  it('finds the master of a slide layout the presentation does not list', async () => {
    const r = await layoutsOf([{ layouts: [{ xml: layout('<p:cSld name="Orphan"/>') }] }], { unlisted: true, slideLayouts: ['../slideLayouts/slideLayout1.xml', '../slideLayouts/missing.xml'] });
    expect(r.layouts.map((l) => l.name)).toEqual(['Orphan']);
  });

  it('ignores a master id without a relationship and one that is not a master', async () => {
    const r = await layoutsOf([{}], { extraMasterIds: '<p:sldMasterId id="9"/><p:sldMasterId id="10" r:id="rIdS1"/>', slideLayouts: ['../slideLayouts/slideLayout1.xml'] });
    expect(r.layouts).toHaveLength(1);
  });

  it('has no layouts for a deck without masters', async () => {
    expect((await layoutsOf([])).layouts).toEqual([]);
  });
});

describe('matchLayoutPlaceholder', () => {
  const ph = (type: string, idx?: number) => ({ type, ...(idx !== undefined ? { idx } : {}), levels: [] });
  const l: PptxLayout = {
    partName: 'p', name: 'L', masterPart: 'm', staticShapes: 0,
    placeholders: [ph('ctrTitle'), ph('subTitle', 1), ph('dt', 10), ph('pic', 13), ph('body', 14)],
  };
  it('prefers the idx, then the same type, then the same kind', () => {
    expect(matchLayoutPlaceholder(l, { type: 'body', idx: 13 })?.type).toBe('pic');
    expect(matchLayoutPlaceholder(l, { type: 'dt', idx: 99 })?.idx).toBe(10);
    expect(matchLayoutPlaceholder(l, { type: 'title' })?.type).toBe('ctrTitle');
    expect(matchLayoutPlaceholder(l, { type: 'ctrTitle', idx: 0 })?.type).toBe('ctrTitle');
    expect(matchLayoutPlaceholder(l, { idx: 0 })?.type).toBe('subTitle');
    expect(matchLayoutPlaceholder(l, {})?.type).toBe('subTitle');
    expect(matchLayoutPlaceholder(l, { type: 'obj' })?.type).toBe('subTitle');
    expect(matchLayoutPlaceholder(l, { type: 'tbl' })).toBeUndefined();
    expect(matchLayoutPlaceholder(undefined, { type: 'title' })).toBeUndefined();
  });

  it('does not pair a slide placeholder naming an idx the layout lacks with a placeholder of the same kind', () => {
    expect(matchLayoutPlaceholder(l, { type: 'obj', idx: 13 })?.type).toBe('pic');
    expect(matchLayoutPlaceholder(l, { type: 'obj', idx: 7 })).toBeUndefined();
    expect(matchLayoutPlaceholder(l, { type: 'subTitle', idx: 7 })?.idx).toBe(1);
    expect(matchLayoutPlaceholder(l, { type: 'title', idx: 7 })).toBeUndefined();
  });
});
