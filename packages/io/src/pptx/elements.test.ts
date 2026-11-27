import JSZip from 'jszip';
import type { PptxElement, PptxLink, PptxPicture, PptxShape, PptxTable, PresentationImportResult } from '@b32nio/spindle-shared/export/pptx';
import { importPptx } from './reader';
import { sniffImage } from './pictures';
import { schemeOf } from './links';

// Pictures, tables and hyperlinks (plan K10a), read from synthetic decks that are real zips.

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

const rel = (id: string, type: string, target: string, external = false): string => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;
const rels = (list: string[]): string => `<Relationships xmlns="${PKG_REL}">${list.join('')}</Relationships>`;
const E = 9525;

interface Deck {
  tree?: string;
  layoutTree?: string;
  slideRels?: string[];
  /** Parts besides the standard ones; a `Uint8Array` is written as bytes. */
  extra?: Record<string, string | Uint8Array>;
  /** Write `ppt/tableStyles.xml` with this content and relate it from the presentation. */
  tableStyles?: string;
}

async function build(d: Deck): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="${CT}"><Override PartName="/ppt/presentation.xml" ContentType="${PRES_TYPE}"/></Types>`);
  zip.file('_rels/.rels', rels([rel('rId1', 'officeDocument', 'ppt/presentation.xml')]));
  zip.file('ppt/_rels/presentation.xml.rels', rels([rel('rId1', 'slide', 'slides/slide1.xml'), rel('rId2', 'slide', 'slides/slide2.xml'), rel('rIdM', 'slideMaster', 'slideMasters/slideMaster1.xml'), ...(d.tableStyles ? [rel('rIdS', 'tableStyles', 'tableStyles.xml')] : [])]));
  zip.file('ppt/presentation.xml', `<p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`);
  for (const n of ['1', '2']) {
    zip.file(`ppt/slides/slide${n}.xml`, `<p:sld ${NS}><p:cSld><p:spTree>${d.tree ?? ''}</p:spTree></p:cSld></p:sld>`);
    zip.file(`ppt/slides/_rels/slide${n}.xml.rels`, rels([rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'), ...(d.slideRels ?? [])]));
  }
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout ${NS}><p:cSld name="L"><p:spTree>${d.layoutTree ?? ''}</p:spTree></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels([rel('rId1', 'slideMaster', '../slideMasters/slideMaster1.xml')]));
  zip.file('ppt/slideMasters/slideMaster1.xml', `<p:sldMaster ${NS}><p:cSld><p:spTree/></p:cSld>${CLR_MAP}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`);
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', rels([rel('rIdT', 'theme', '../theme/theme1.xml'), rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml')]));
  zip.file('ppt/theme/theme1.xml', THEME);
  if (d.tableStyles) zip.file('ppt/tableStyles.xml', d.tableStyles);
  for (const [name, content] of Object.entries(d.extra ?? {})) zip.file(name, content);
  return zip.generateAsync({ type: 'uint8array' });
}

const read = async (d: Deck): Promise<PresentationImportResult> => importPptx(await build(d));
const elementsOf = async (tree: string, d: Deck = {}): Promise<PptxElement[]> => (await read({ ...d, tree })).slides[0].elements as PptxElement[];

const PNG = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3]);
const JPEG = Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 1, 2]);
const ascii = (s: string, tail: number[] = []): Uint8Array => Uint8Array.from([...s].map((c) => c.charCodeAt(0)).concat(tail));
const b64 = (b: Uint8Array): string => Buffer.from(b).toString('base64');

const xfrm = (x: number, y: number, w: number, h: number, attrs = ''): string => `<a:xfrm${attrs}><a:off x="${x * E}" y="${y * E}"/><a:ext cx="${w * E}" cy="${h * E}"/></a:xfrm>`;
const pic = (o: { id?: number; name?: string; cNvPr?: string; blip?: string; blipInner?: string; blipFill?: string; spPr?: string; nvPr?: string } = {}): string =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${o.id ?? 3}" name="${o.name ?? 'P'}"${o.cNvPr ?? ''}/><p:cNvPicPr/><p:nvPr>${o.nvPr ?? ''}</p:nvPr></p:nvPicPr><p:blipFill>${o.blip ?? `<a:blip r:embed="rImg">${o.blipInner ?? ''}</a:blip>`}${o.blipFill ?? '<a:stretch><a:fillRect/></a:stretch>'}</p:blipFill><p:spPr>${o.spPr ?? ''}</p:spPr></p:pic>`;
const image = (target: string): string => rel('rImg', 'image', target);
const sp = (cNvPr = '', body = '<a:p><a:r><a:t>x</a:t></a:r></a:p>'): string => `<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"${cNvPr}>${''}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody></p:sp>`;

describe('sniffImage', () => {
  it.each([
    ['png', PNG, 'image/png'],
    ['jpeg', JPEG, 'image/jpeg'],
    ['gif', ascii('GIF89a'), 'image/gif'],
    ['webp', ascii('RIFF\0\0\0\0WEBPVP8 '), 'image/webp'],
    ['bmp', Uint8Array.from([0x42, 0x4d, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 40, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]), 'image/bmp'],
    ['avif', ascii('\0\0\0\x1cftypavif'), 'image/avif'],
    ['svg', ascii('<svg xmlns="http://www.w3.org/2000/svg"/>'), 'image/svg+xml'],
    ['svg with a declaration and a comment', ascii('<?xml version="1.0"?>\n<!-- c -->\n<!DOCTYPE svg>\n<svg>'), 'image/svg+xml'],
    ['svg after a byte-order mark', new TextEncoder().encode('\uFEFF<svg>'), 'image/svg+xml'],
  ])('recognises %s from its bytes', (_name, bytes, mime) => {
    expect(sniffImage(bytes)).toBe(mime);
  });

  it.each([
    ['a TIFF', ascii('II*\0')],
    ['text', ascii('hello')],
    ['an HTML page', ascii('<html><svg></html>')],
    ['a RIFF that is not WebP', ascii('RIFF\0\0\0\0WAVE')],
    ['text starting with BM', ascii('BMW is a car maker, not a bitmap image')],
    ['a BM header with a short DIB size', Uint8Array.from([0x42, 0x4d, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])],
    ['a truncated BM header', ascii('BM\0\0')],
    ['nothing', new Uint8Array()],
  ])('gives nothing for %s', (_name, bytes) => {
    expect(sniffImage(bytes)).toBeUndefined();
  });
});

describe('p:pic', () => {
  it('reads an embedded picture as a data URL typed from its bytes, with frame, geometry and outline', async () => {
    const [p] = (await elementsOf(
      pic({ id: 9, name: ' Logo ', cNvPr: ' descr=" alt "', spPr: `${xfrm(10, 20, 300, 200, ' rot="5400000" flipH="1"')}<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>` }),
      { slideRels: [image('../media/logo.jpg')], extra: { 'ppt/media/logo.jpg': PNG } }
    )) as [PptxPicture];
    // the extension says JPEG, the bytes say PNG
    expect(p).toMatchObject({
      kind: 'picture', pptxId: '9', name: 'Logo', altText: 'alt', part: 'ppt/media/logo.jpg', mime: 'image/png',
      dataUrl: `data:image/png;base64,${b64(PNG)}`,
      frame: { x: 10, y: 20, w: 300, h: 200, rot: 90, flipH: true },
      geometry: { kind: 'preset', prst: 'roundRect', known: true },
      line: { width: 2 },
    });
    expect(p.unsupported).toBeUndefined();
    expect(p.crop).toBeUndefined();
    expect(p.hidden).toBeUndefined();
  });

  it('encodes a large image in chunks and encodes a picture used twice once', async () => {
    const big = new Uint8Array(100_000).fill(7);
    big.set(PNG);
    const [a, b] = (await elementsOf(pic({ id: 3 }) + pic({ id: 4 }), { slideRels: [image('../media/big.png')], extra: { 'ppt/media/big.png': big } })) as PptxPicture[];
    expect(a.dataUrl).toBe(`data:image/png;base64,${b64(big)}`);
    expect(b.dataUrl).toBe(a.dataUrl);
  });

  it('reads the PNG fallback of an SVG picture', async () => {
    const [p] = (await elementsOf(pic({ blipInner: '<a:extLst><a:ext uri="x"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rSvg"/></a:ext></a:extLst>' }), {
      slideRels: [image('../media/fallback.png'), rel('rSvg', 'image', '../media/v.svg')],
      extra: { 'ppt/media/fallback.png': PNG, 'ppt/media/v.svg': ascii('<svg/>') },
    })) as [PptxPicture];
    expect(p.mime).toBe('image/png');
    expect(p.part).toBe('ppt/media/fallback.png');
  });

  it('reads a picture that is an SVG itself', async () => {
    const [p] = (await elementsOf(pic(), { slideRels: [image('../media/v.svg')], extra: { 'ppt/media/v.svg': ascii('<svg/>') } })) as [PptxPicture];
    expect(p.mime).toBe('image/svg+xml');
  });

  it.each([
    ['emf', '../media/a.emf', 'emf', /EMF or WMF/],
    ['emz', '../media/a.emz', 'emf', /EMF or WMF/],
    ['wmf', '../media/a.wmf', 'wmf', /EMF or WMF/],
    ['wmz', '../media/a.WMZ', 'wmf', /EMF or WMF/],
  ])('keeps a %s picture as a placeholder without reading it', async (_n, target, unsupported, message) => {
    const result = await read({ tree: pic({ cNvPr: ' descr="chart"' }), slideRels: [image(target)], extra: { [`ppt/media/${target.slice(9)}`]: PNG } });
    const [p] = result.slides[0].elements as [PptxPicture];
    expect(p).toMatchObject({ kind: 'picture', altText: 'chart', unsupported, part: `ppt/media/${target.slice(9)}`, });
    expect(p.dataUrl).toBeUndefined();
    expect(result.warnings.some((w) => message.test(w))).toBe(true);
  });

  it('keeps an image in a format browsers cannot show as a placeholder', async () => {
    const result = await read({ tree: pic(), slideRels: [image('../media/a.tif')], extra: { 'ppt/media/a.tif': ascii('II*\0') } });
    expect((result.slides[0].elements as PptxPicture[])[0].unsupported).toBe('format');
    expect(result.warnings).toContain('A picture is in a format browsers cannot show; it is kept as a placeholder.');
  });

  it('keeps a picture whose image part is missing, or whose relationship is, as a placeholder', async () => {
    const missing = await read({ tree: pic(), slideRels: [image('../media/gone.png')] });
    expect((missing.slides[0].elements as PptxPicture[])[0]).toMatchObject({ unsupported: 'missing', part: 'ppt/media/gone.png' });
    const noRel = await read({ tree: pic({ blip: '<a:blip/>' }) });
    const [p] = noRel.slides[0].elements as [PptxPicture];
    expect(p.unsupported).toBe('missing');
    expect(p.part).toBeUndefined();
    expect(noRel.warnings).toContain('A picture has no readable image; it is kept as a placeholder.');
    const noBlip = await read({ tree: pic({ blip: '' }) });
    expect((noBlip.slides[0].elements as PptxPicture[])[0].unsupported).toBe('missing');
  });

  it('never fetches a linked picture: it is a placeholder', async () => {
    const result = await read({ tree: pic({ blip: '<a:blip r:link="rLink"/>' }), slideRels: [rel('rLink', 'image', 'https://example.com/a.png', true)] });
    const [p] = result.slides[0].elements as [PptxPicture];
    expect(p).toMatchObject({ unsupported: 'linked' });
    expect(p.part).toBeUndefined();
    expect(result.warnings).toContain('A picture is linked from outside the file, which is never fetched; it is kept as a placeholder.');
    // an embed that points at an external target is not read either
    const external = await read({ tree: pic(), slideRels: [rel('rImg', 'image', 'https://example.com/a.png', true)] });
    expect((external.slides[0].elements as PptxPicture[])[0]).toMatchObject({ unsupported: 'linked' });
    expect((external.slides[0].elements as PptxPicture[])[0].dataUrl).toBeUndefined();
  });

  it('reads the crop in percent, clamps an outset and ignores a crop that trims nothing', async () => {
    const opts = { slideRels: [image('../media/a.png')], extra: { 'ppt/media/a.png': PNG } };
    const crop = async (rect: string): Promise<PptxPicture['crop']> => ((await elementsOf(pic({ blipFill: `${rect}<a:stretch/>` }), opts)) as PptxPicture[])[0].crop;
    expect(await crop('<a:srcRect l="25000" t="10000" r="25000" b="40000"/>')).toEqual({ l: 25, t: 10, r: 25, b: 40 });
    expect(await crop('<a:srcRect l="-5000" r="90000"/>')).toEqual({ l: 0, t: 0, r: 90, b: 0 });
    expect(await crop('<a:srcRect r="150000"/>')).toBeUndefined();
    expect(await crop('<a:srcRect l="0" t="-1000"/>')).toBeUndefined();
    expect(await crop('')).toBeUndefined();
    // edges that meet or overlap leave nothing visible: ignored
    expect(await crop('<a:srcRect l="60000" r="40000"/>')).toBeUndefined();
    expect(await crop('<a:srcRect t="70000" b="60000"/>')).toBeUndefined();
    expect(await crop('<a:srcRect l="60000" r="39000"/>')).toEqual({ l: 60, t: 0, r: 39, b: 0 });
  });

  it('reads transparency, ignoring full opacity', async () => {
    const [p, q] = (await elementsOf(pic({ blipInner: '<a:alphaModFix amt="40000"/>' }) + pic({ id: 4, blipInner: '<a:alphaModFix amt="100000"/>' }), { slideRels: [image('../media/a.png')], extra: { 'ppt/media/a.png': PNG } })) as PptxPicture[];
    expect(p.alpha).toBeCloseTo(0.4);
    expect(q.alpha).toBeUndefined();
  });

  it('reads the hidden flag and a hyperlink of a picture', async () => {
    const tree = `<p:pic><p:nvPicPr><p:cNvPr id="3" name="P" hidden="true"><a:hlinkClick r:id="rLink"/></p:cNvPr><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rImg"/></p:blipFill><p:spPr/></p:pic>`;
    const [p] = (await elementsOf(tree, { slideRels: [image('../media/a.png'), rel('rLink', 'hyperlink', 'https://example.com/', true)], extra: { 'ppt/media/a.png': PNG } })) as [PptxPicture];
    expect(p.hidden).toBe(true);
    expect(p.link).toEqual({ kind: 'url', href: 'https://example.com/' });
  });

  it('gives a picture placeholder the frame of its layout placeholder', async () => {
    const layoutTree = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Pic"/><p:cNvSpPr/><p:nvPr><p:ph type="pic" idx="13"/></p:nvPr></p:nvSpPr><p:spPr>${xfrm(40, 50, 400, 300)}</p:spPr></p:sp>`;
    const [p] = (await elementsOf(pic({ nvPr: '<p:ph type="pic" idx="13"/>' }), { layoutTree, slideRels: [image('../media/a.png')], extra: { 'ppt/media/a.png': PNG } })) as [PptxPicture];
    expect(p.placeholder).toEqual({ type: 'pic', idx: 13 });
    expect(p.frame).toMatchObject({ x: 40, y: 50, w: 400, h: 300 });
  });

  it('reads a picture with no non-visual properties, and a slide with no shape tree', async () => {
    const [p] = (await elementsOf('<p:pic><p:blipFill><a:blip r:embed="rImg"/></p:blipFill></p:pic>', { slideRels: [image('../media/a.png')], extra: { 'ppt/media/a.png': PNG } })) as [PptxPicture];
    expect(p).toMatchObject({ pptxId: '', name: '', mime: 'image/png' });
    const result = await read({ extra: { 'ppt/slides/slide1.xml': `<p:sld ${NS}><p:cSld/></p:sld>` } });
    expect(result.slides[0].elements).toEqual([]);
  });

  it('reads pictures in drawing order among shapes, also inside alternate content', async () => {
    const tree = `${sp()}<mc:AlternateContent><mc:Choice Requires="x">${pic({ id: 5 })}</mc:Choice><mc:Fallback>${pic({ id: 6 })}</mc:Fallback></mc:AlternateContent>${sp()}`;
    const els = await elementsOf(tree, { slideRels: [image('../media/a.png')], extra: { 'ppt/media/a.png': PNG } });
    expect(els.map((e) => [e.kind, e.pptxId])).toEqual([['shape', '2'], ['picture', '5'], ['shape', '2']]);
    const fallback = await elementsOf(`<mc:AlternateContent><mc:Fallback>${pic({ id: 6 })}</mc:Fallback></mc:AlternateContent><p:grpSp/><p:cxnSp/><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8" name="Chart"/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></p:graphicFrame>`, { slideRels: [image('../media/a.png')], extra: { 'ppt/media/a.png': PNG } });
    expect(fallback.map((e) => e.pptxId)).toEqual(['6']);
  });
});

const link = (target: string, attrs = '', external = true): { tree: string; deck: Deck } => ({
  tree: sp('', `<a:p><a:r><a:rPr><a:hlinkClick r:id="rL"${attrs}/></a:rPr><a:t>x</a:t></a:r></a:p>`),
  deck: { slideRels: [rel('rL', external ? 'hyperlink' : 'slide', target, external)] },
});
const runLink = async (tree: string, deck: Deck): Promise<{ link: PptxLink | undefined; warnings: string[] }> => {
  const result = await read({ ...deck, tree });
  return { link: (result.slides[0].elements as PptxShape[])[0].text?.paragraphs[0].runs[0].link, warnings: result.warnings };
};

describe('hyperlinks', () => {
  it.each(['http://example.com/a?b=1', 'https://example.com/', 'HTTPS://EXAMPLE.COM', 'mailto:a@example.com', 'tel:+4912345'])('keeps %s', async (href) => {
    const { tree, deck } = link(href);
    expect((await runLink(tree, deck)).link).toEqual({ kind: 'url', href });
  });

  it('keeps the tooltip', async () => {
    const { tree, deck } = link('https://example.com/', ' tooltip=" go "');
    expect((await runLink(tree, deck)).link).toEqual({ kind: 'url', href: 'https://example.com/', tooltip: 'go' });
  });

  it.each(['javascript:alert(1)', 'java\tscript:alert(1)', ' JavaScript:alert(1)', 'data:text/html,x', 'vbscript:x', 'ftp://example.com/', 'blob:https://x/y'])('drops %j and warns with the scheme only', async (href) => {
    const { tree, deck } = link(href);
    const { link: kept, warnings } = await runLink(tree, deck);
    expect(kept).toBeUndefined();
    expect(warnings.some((w) => /^A hyperlink with a "[a-z]+:" address was dropped; only http, https, mailto and tel links are kept\.$/.test(w))).toBe(true);
    expect(warnings.join()).not.toContain('alert');
  });

  it('drops a relative address with a warning', async () => {
    const { tree, deck } = link('other.html');
    const { link: kept, warnings } = await runLink(tree, deck);
    expect(kept).toBeUndefined();
    expect(warnings).toContain('A hyperlink with a relative address was dropped.');
  });

  it('keeps a file link as an inert action, with a warning', async () => {
    const { tree, deck } = link('file:///C:/secret.docx');
    const { link: kept, warnings } = await runLink(tree, deck);
    expect(kept).toEqual({ kind: 'action', target: 'file:///C:/secret.docx' });
    expect(warnings).toContain('A link to a file is kept as an inert action; it is never opened.');
  });

  it('keeps a ppaction target inert, with the slide it jumps to', async () => {
    const jump = link('slide2.xml', ' action="ppaction://hlinksldjump"', false);
    expect((await runLink(jump.tree, jump.deck)).link).toEqual({ kind: 'action', target: 'ppaction://hlinksldjump', slidePart: 'ppt/slides/slide2.xml' });
    const next = { tree: sp('', '<a:p><a:r><a:rPr><a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=nextslide"/></a:rPr><a:t>x</a:t></a:r></a:p>'), deck: {} };
    expect((await runLink(next.tree, next.deck)).link).toEqual({ kind: 'action', target: 'ppaction://hlinkshowjump?jump=nextslide' });
    const bare = { tree: sp('', '<a:p><a:r><a:rPr><a:hlinkClick action="ppaction://hlinkshowjump?jump=lastslide"/></a:rPr><a:t>x</a:t></a:r></a:p>'), deck: {} };
    expect((await runLink(bare.tree, bare.deck)).link).toEqual({ kind: 'action', target: 'ppaction://hlinkshowjump?jump=lastslide' });
    const none = { tree: sp('', '<a:p><a:r><a:rPr><a:hlinkClick/></a:rPr><a:t>x</a:t></a:r></a:p>'), deck: {} };
    expect((await runLink(none.tree, none.deck)).link).toBeUndefined();
    const macro = link('https://example.com/', ' action="ppaction://macro?name=Evil"');
    expect((await runLink(macro.tree, macro.deck)).link).toEqual({ kind: 'action', target: 'ppaction://macro?name=Evil' });
  });

  it('reads a jump to a slide that has no action text, and drops an internal link that is not a slide', async () => {
    const jump = link('slide2.xml', '', false);
    expect((await runLink(jump.tree, jump.deck)).link).toEqual({ kind: 'action', target: 'ppaction://hlinksldjump', slidePart: 'ppt/slides/slide2.xml' });
    const other = { tree: link('', '', false).tree, deck: { slideRels: [rel('rL', 'notesSlide', '../notesSlides/n.xml')] } };
    expect((await runLink(other.tree, other.deck)).link).toBeUndefined();
    const jumpAction = link('slide2.xml', ' action="ppaction://hlinksldjump" tooltip="t"', false);
    expect((await runLink(jumpAction.tree, jumpAction.deck)).link).toMatchObject({ tooltip: 't' });
  });

  it('gives nothing for a link without a relationship or without a link at all', async () => {
    expect((await runLink(link('x').tree, {})).link).toBeUndefined();
    expect((await runLink(sp('', '<a:p><a:r><a:t>x</a:t></a:r></a:p>'), {})).link).toBeUndefined();
  });

  it('reads the link of a field run and of a shape', async () => {
    const field = sp('', '<a:p><a:fld id="{1}" type="slidenum"><a:rPr><a:hlinkClick r:id="rL"/></a:rPr><a:t>3</a:t></a:fld></a:p>');
    const deck = { slideRels: [rel('rL', 'hyperlink', 'mailto:a@b.c', true)] };
    expect((await runLink(field, deck)).link).toEqual({ kind: 'url', href: 'mailto:a@b.c' });
    const shape = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"><a:hlinkClick r:id="rL"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`;
    expect(((await elementsOf(shape, deck)) as PptxShape[])[0].link).toEqual({ kind: 'url', href: 'mailto:a@b.c' });
  });

  it('finds the scheme through whitespace and control characters', () => {
    expect(schemeOf('ja\u0000va\nscript:x')).toBe('javascript');
    expect(schemeOf('a/b:c')).toBeUndefined();
    expect(schemeOf('')).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

const cell = (text: string, attrs = '', tcAttrs = '', tcInner = ''): string => `<a:tc${attrs}><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></a:txBody><a:tcPr${tcAttrs}>${tcInner}</a:tcPr></a:tc>`;
const table = (rows: string[][], o: { pr?: string; grid?: number; name?: string; cNvPr?: string; frame?: string; heights?: number[] } = {}): string => {
  const cols = o.grid ?? Math.max(...rows.map((r) => r.length));
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="${o.name ?? 'T'}"${o.cNvPr ?? ''}/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>${o.frame ?? `<p:xfrm><a:off x="${10 * E}" y="${20 * E}"/><a:ext cx="${400 * E}" cy="${100 * E}"/></p:xfrm>`}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>${o.pr ?? '<a:tblPr/>'}<a:tblGrid>${Array.from({ length: cols }, () => `<a:gridCol w="${100 * E}"/>`).join('')}</a:tblGrid>${rows.map((r, i) => `<a:tr h="${(o.heights?.[i] ?? 30) * E}">${r.join('')}</a:tr>`).join('')}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
};
const tableOnly = async (tree: string, d: Deck = {}): Promise<PptxTable> => (await elementsOf(tree, d))[0] as PptxTable;
const textOf = (t: PptxTable, r: number, c: number): string => t.rows[r].cells[c].text.paragraphs.flatMap((p) => p.runs.map((x) => x.text)).join('');

describe('tables', () => {
  it('reads the grid, row heights, frame and cell text', async () => {
    const t = await tableOnly(table([[cell('A'), cell('B')], [cell('C'), cell('D')]], { name: ' Data ', cNvPr: ' descr=" d " hidden="1"', heights: [30, 40] }));
    expect(t).toMatchObject({ kind: 'table', pptxId: '4', name: 'Data', altText: 'd', hidden: true, frame: { x: 10, y: 20, w: 400, h: 100 }, columns: [100, 100] });
    expect(t.rows.map((r) => r.height)).toEqual([30, 40]);
    expect([textOf(t, 0, 0), textOf(t, 0, 1), textOf(t, 1, 0), textOf(t, 1, 1)]).toEqual(['A', 'B', 'C', 'D']);
    expect(t.flags).toBeUndefined();
    expect(t.styleId).toBeUndefined();
  });

  it('reads spans and merged-cell markers', async () => {
    const t = await tableOnly(table([[cell('A', ' gridSpan="2"'), cell('', ' hMerge="1"')], [cell('B', ' rowSpan="2"'), cell('C')], [cell('', ' vMerge="true"'), cell('D', ' gridSpan="1" rowSpan="x"')]]));
    expect(t.rows[0].cells[0].gridSpan).toBe(2);
    expect(t.rows[0].cells[1].hMerge).toBe(true);
    expect(t.rows[1].cells[0].rowSpan).toBe(2);
    expect(t.rows[2].cells[0].vMerge).toBe(true);
    expect(t.rows[2].cells[1].gridSpan).toBeUndefined();
    expect(t.rows[2].cells[1].rowSpan).toBeUndefined();
  });

  it('reads a cell\'s own fill, borders, margins, anchor and vertical text', async () => {
    const inner = '<a:lnL w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnL><a:lnT><a:noFill/></a:lnT><a:lnR w="25400"/><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>';
    const t = await tableOnly(table([[cell('A', '', ' marL="95250" marT="19050" marR="0" marB="-1" anchor="ctr" vert="vert"', inner)]]));
    const c = t.rows[0].cells[0];
    expect(c.margins).toEqual({ l: 10, t: 2, r: 0 });
    expect(c.anchor).toBe('middle');
    expect(c.vert).toBe('vert');
    expect(c.fill).toEqual({ kind: 'solid', color: { kind: 'rgb', hex: '#00FF00' } });
    expect(c.borders).toEqual({ l: { width: 1.3333333333333333, fill: { kind: 'solid', color: { kind: 'rgb', hex: '#FF0000' } } }, t: null, r: { width: 2.6666666666666665 } });
  });

  it('keeps horizontal text and unknown anchors as nothing', async () => {
    const c = (await tableOnly(table([[cell('A', '', ' vert="horz" anchor="x"')]]))).rows[0].cells[0];
    expect(c.vert).toBeUndefined();
    expect(c.anchor).toBeUndefined();
    expect(c.borders).toBeUndefined();
    expect(c.margins).toBeUndefined();
  });

  it('pads short rows and widens the grid to the widest row (extra columns have width 0)', async () => {
    const t = await tableOnly(table([[cell('A'), cell('B'), cell('C')], [cell('D')]], { grid: 2 }));
    expect(t.columns).toHaveLength(3);
    expect(t.columns[2]).toBe(0);
    expect(t.rows[1].cells).toHaveLength(3);
    expect(textOf(t, 1, 1)).toBe('');
    const empty = await tableOnly(table([[cell('A')]], { grid: 3 }));
    expect(empty.rows[0].cells).toHaveLength(3);
  });

  it('reads a very wide row in linear time', async () => {
    const wide = await tableOnly(table([Array.from({ length: 3000 }, (_, i) => cell(String(i)))], { grid: 3000 }));
    expect(wide.rows[0].cells).toHaveLength(3000);
    expect(textOf(wide, 0, 2999)).toBe('2999');
  });

  it('reads a table without a frame, and an empty table', async () => {
    const t = await tableOnly(table([[cell('A')]], { frame: '' }));
    expect(t.frame).toBeUndefined();
    const none = await tableOnly(`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="T"/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl/></a:graphicData></a:graphic></p:graphicFrame>`);
    expect(none).toMatchObject({ columns: [], rows: [] });
    const bad = await tableOnly(`<p:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid><a:gridCol w="x"/></a:tblGrid><a:tr h="x"><a:tc/></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`);
    expect(bad).toMatchObject({ pptxId: '', name: '', columns: [0] });
    expect(bad.rows[0].height).toBe(0);
  });

  it('skips a graphic frame that holds no table', async () => {
    const frame = (uri: string, inner = ''): string => `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="G"/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic></p:graphicFrame>`;
    expect(await elementsOf(frame('http://schemas.openxmlformats.org/drawingml/2006/chart'))).toEqual([]);
    // a table uri without a table inside has nothing to read
    expect(await elementsOf(frame('http://schemas.openxmlformats.org/drawingml/2006/table'))).toEqual([]);
    expect(await elementsOf('<p:graphicFrame/>')).toEqual([]);
  });

  it('reads a table in alternate content', async () => {
    const t = (await elementsOf(`<mc:AlternateContent><mc:Choice Requires="x">${table([[cell('A')]])}</mc:Choice></mc:AlternateContent>`))[0] as PptxTable;
    expect(textOf(t, 0, 0)).toBe('A');
  });

  it('reads hyperlinks in cell text', async () => {
    const withLink = `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:hlinkClick r:id="rL"/></a:rPr><a:t>x</a:t></a:r></a:p></a:txBody></a:tc>`;
    const t = await tableOnly(table([[withLink]]), { slideRels: [rel('rL', 'hyperlink', 'https://example.com/', true)] });
    expect(t.rows[0].cells[0].text.paragraphs[0].runs[0].link).toEqual({ kind: 'url', href: 'https://example.com/' });
  });
});

const line = (slot: string, w = 12700, inner = ''): string => `<a:ln w="${w}"><a:solidFill><a:schemeClr val="${slot}">${inner}</a:schemeClr></a:solidFill></a:ln>`;
const borders = (all: string): string => ['left', 'right', 'top', 'bottom', 'insideH', 'insideV'].map((e) => `<a:${e}>${all}</a:${e}>`).join('');
const STYLE_ID = '{AAAA}';
const styles = (inner: string, extra = '', attrs = ''): string => `<a:tblStyleLst ${NS} def="${STYLE_ID}"><a:tblStyle styleId="${STYLE_ID}"${attrs}>${inner}</a:tblStyle>${extra}</a:tblStyleLst>`;
const fill = (inner: string): string => `<a:fill><a:solidFill>${inner}</a:solidFill></a:fill>`;
const part = (name: string, o: { fill?: string; bdr?: string; tx?: string; txAttrs?: string }): string =>
  `<a:${name}>${o.tx !== undefined || o.txAttrs !== undefined ? `<a:tcTxStyle${o.txAttrs ?? ''}>${o.tx ?? ''}</a:tcTxStyle>` : ''}<a:tcStyle>${o.bdr !== undefined ? `<a:tcBdr>${o.bdr}</a:tcBdr>` : ''}${o.fill ?? ''}</a:tcStyle></a:${name}>`;
const styled = (rows: number, cols: number, flags: string, text = (r: number, c: number): string => `r${r}c${c}`): string =>
  table(Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => cell(text(r, c)))), { pr: `<a:tblPr ${flags}><a:tableStyleId>${STYLE_ID}</a:tableStyleId></a:tblPr>` });
const accent = (n: number): { kind: 'theme'; slot: string } => ({ kind: 'theme', slot: `accent${n}` });

describe('table styles', () => {
  const solid = (n: number, extra = {}): unknown => ({ kind: 'solid', color: { ...accent(n), ...extra } });

  it('resolves the whole-table fill, borders and text into every cell and keeps the style name', async () => {
    const ts = styles(part('wholeTbl', { fill: fill('<a:schemeClr val="accent1"/>'), bdr: borders(line('accent2')), tx: '<a:schemeClr val="accent3"/>', txAttrs: ' b="on" i="on"' }), '', ' styleName="Mine"');
    const t = await tableOnly(styled(2, 2, ''), { tableStyles: ts });
    expect(t).toMatchObject({ styleId: STYLE_ID, styleName: 'Mine' });
    for (const [r, c] of [[0, 0], [0, 1], [1, 0], [1, 1]]) {
      const cellData = t.rows[r].cells[c];
      expect(cellData.fill).toEqual(solid(1));
      expect(Object.keys(cellData.borders ?? {}).sort()).toEqual(['b', 'l', 'r', 't']);
      expect(cellData.borders?.t).toMatchObject({ fill: { kind: 'solid', color: accent(2) }, width: 12700 / E });
      const style = cellData.text.paragraphs[0].style;
      expect(style).toMatchObject({ bold: true, italic: true, color: accent(3) });
    }
  });

  it('uses the outer edges at the table border and the inside edges within', async () => {
    const edges = [['left', 'accent1'], ['right', 'accent2'], ['top', 'accent3'], ['bottom', 'accent4'], ['insideH', 'accent5'], ['insideV', 'accent6']];
    const ts = styles(part('wholeTbl', { bdr: edges.map(([e, slot]) => `<a:${e}>${line(slot)}</a:${e}>`).join('') }));
    const t = await tableOnly(styled(3, 3, ''), { tableStyles: ts });
    const side = (r: number, c: number, s: 'l' | 't' | 'r' | 'b'): string => ((t.rows[r].cells[c].borders?.[s] as { fill: { color: { slot: string } } }).fill.color.slot);
    expect([side(0, 0, 'l'), side(0, 0, 't'), side(0, 0, 'r'), side(0, 0, 'b')]).toEqual(['accent1', 'accent3', 'accent6', 'accent5']);
    expect([side(2, 2, 'l'), side(2, 2, 't'), side(2, 2, 'r'), side(2, 2, 'b')]).toEqual(['accent6', 'accent5', 'accent2', 'accent4']);
    expect([side(1, 1, 'l'), side(1, 1, 't'), side(1, 1, 'r'), side(1, 1, 'b')]).toEqual(['accent6', 'accent5', 'accent6', 'accent5']);
  });

  it('bands rows after the header row and columns after the first column', async () => {
    const ts = styles(part('band1H', { fill: fill('<a:schemeClr val="accent1"/>') }) + part('band2H', { fill: fill('<a:schemeClr val="accent2"/>') }) + part('band1V', { fill: fill('<a:schemeClr val="accent3"/>') }) + part('band2V', { fill: fill('<a:schemeClr val="accent4"/>') }) + part('firstRow', { fill: fill('<a:schemeClr val="accent5"/>') }) + part('firstCol', { fill: fill('<a:schemeClr val="accent6"/>') }));
    const slot = (t: PptxTable, r: number, c: number): string | undefined => (t.rows[r].cells[c].fill as { color: { slot: string } } | undefined)?.color.slot;

    const rows = await tableOnly(styled(4, 1, 'firstRow="1" bandRow="1"'), { tableStyles: ts });
    expect([0, 1, 2, 3].map((r) => slot(rows, r, 0))).toEqual(['accent5', 'accent1', 'accent2', 'accent1']);
    const plain = await tableOnly(styled(3, 1, 'bandRow="1"'), { tableStyles: ts });
    expect([0, 1, 2].map((r) => slot(plain, r, 0))).toEqual(['accent1', 'accent2', 'accent1']);
    const cols = await tableOnly(styled(1, 4, 'firstCol="1" bandCol="1"'), { tableStyles: ts });
    expect([0, 1, 2, 3].map((c) => slot(cols, 0, c))).toEqual(['accent6', 'accent3', 'accent4', 'accent3']);
    // no banding switch, no band fill
    const off = await tableOnly(styled(2, 2, ''), { tableStyles: ts });
    expect(off.rows[0].cells[0].fill).toBeUndefined();
  });

  it('lets the last column and row, and then the header row, win over the bands and the whole table', async () => {
    const ts = styles(part('wholeTbl', { fill: fill('<a:schemeClr val="accent1"/>') }) + part('band1H', { fill: fill('<a:schemeClr val="accent2"/>') }) + part('lastCol', { fill: fill('<a:schemeClr val="accent3"/>') }) + part('lastRow', { fill: fill('<a:schemeClr val="accent4"/>') }) + part('firstRow', { fill: fill('<a:schemeClr val="accent5"/>') }));
    const t = await tableOnly(styled(4, 3, 'firstRow="1" lastRow="1" lastCol="1" bandRow="1" bandCol="1"'), { tableStyles: ts });
    const slot = (r: number, c: number): string => (t.rows[r].cells[c].fill as { color: { slot: string } }).color.slot;
    expect(slot(0, 2)).toBe('accent5');
    expect(slot(3, 2)).toBe('accent4');
    expect(slot(1, 2)).toBe('accent3');
    expect(slot(1, 0)).toBe('accent2');
    expect(slot(2, 0)).toBe('accent1');
    // a banded column with no band parts adds nothing
    expect(slot(2, 1)).toBe('accent1');
  });

  it('resolves a tint as the lighter variant and keeps an explicit luminance', async () => {
    const ts = styles(part('wholeTbl', { fill: fill('<a:schemeClr val="accent1"><a:tint val="20000"/></a:schemeClr>'), bdr: `<a:left>${line('accent2', 12700, '<a:tint val="50000"/><a:lumMod val="60000"/>')}</a:left>` }) + part('firstRow', { fill: fill('<a:srgbClr val="112233"><a:tint val="20000"/></a:srgbClr>') }));
    const t = await tableOnly(styled(2, 1, 'firstRow="1"'), { tableStyles: ts });
    expect(t.rows[1].cells[0].fill).toEqual({ kind: 'solid', color: { kind: 'theme', slot: 'accent1', lumMod: 0.2, lumOff: 0.8 } });
    expect((t.rows[1].cells[0].borders?.l as { fill: { color: unknown } }).fill.color).toEqual({ kind: 'theme', slot: 'accent2', lumMod: 0.6 });
    expect(t.rows[0].cells[0].fill).toEqual({ kind: 'solid', color: { kind: 'rgb', hex: '#112233' } });
  });

  it('reads no-fill and no-line parts, text switched off, and text without a colour', async () => {
    const ts = styles(
      part('wholeTbl', { fill: fill('<a:schemeClr val="accent1"/>'), tx: '<a:schemeClr val="accent2"/>', txAttrs: ' b="on" i="on"', bdr: `<a:left>${line('accent1')}</a:left><a:right>${line('accent1')}</a:right>` }) +
        part('firstRow', { fill: '<a:fill><a:noFill/></a:fill>', txAttrs: ' b="off" i="off"', bdr: '<a:left><a:ln><a:noFill/></a:ln></a:left><a:right><a:ln w="1"/></a:right><a:top/>' }) +
        part('lastRow', { fill: '<a:fill><a:gradFill/></a:fill>' })
    );
    const t = await tableOnly(styled(3, 2, 'firstRow="1" lastRow="1"'), { tableStyles: ts });
    const head = t.rows[0].cells[0];
    expect(head.fill).toEqual({ kind: 'none' });
    expect(head.borders?.l).toBeNull();
    expect(t.rows[0].cells[1].borders?.r).toMatchObject({ width: 1 / E });
    expect(head.text.paragraphs[0].style).toMatchObject({ bold: false, italic: false, color: accent(2) });
    // the last row states a fill the reader cannot use, so the whole-table one stays
    expect(t.rows[2].cells[0].fill).toEqual(solid(1));
  });

  it('draws an edge in black when its colour cannot be read', async () => {
    const ts = styles(part('wholeTbl', { bdr: '<a:left><a:ln w="12700"><a:solidFill><a:schemeClr val="nope"/></a:solidFill></a:ln></a:left>' }));
    const t = await tableOnly(styled(1, 1, ''), { tableStyles: ts });
    expect(t.rows[0].cells[0].borders?.l).toMatchObject({ fill: { kind: 'solid', color: { kind: 'rgb', hex: '#000000' } } });
  });

  it('ignores corner cells and a style without an id, and reads a style with no regions', async () => {
    const ts = `<a:tblStyleLst ${NS}><a:tblStyle styleName="anon"><a:wholeTbl><a:tcStyle/></a:wholeTbl></a:tblStyle><a:tblStyle styleId="${STYLE_ID}"/></a:tblStyleLst>`;
    const t = await tableOnly(styled(1, 1, ''), { tableStyles: ts });
    expect(t.rows[0].cells[0].fill).toBeUndefined();
    expect(t.rows[0].cells[0].borders).toBeUndefined();
    expect(t.styleName).toBeUndefined();
  });

  it('lets the cell\'s own fill, borders and run properties win over the style', async () => {
    const ts = styles(part('wholeTbl', { fill: fill('<a:schemeClr val="accent1"/>'), bdr: borders(line('accent2')), tx: '<a:schemeClr val="accent3"/>', txAttrs: ' b="on"' }));
    const own = `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="0"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:rPr><a:t>x</a:t></a:r></a:p></a:txBody><a:tcPr><a:lnL w="9525"><a:noFill/></a:lnL><a:solidFill><a:srgbClr val="010203"/></a:solidFill></a:tcPr></a:tc>`;
    const t = await tableOnly(table([[own]], { pr: `<a:tblPr><a:tableStyleId>${STYLE_ID}</a:tableStyleId></a:tblPr>` }), { tableStyles: ts });
    const c = t.rows[0].cells[0];
    expect(c.fill).toEqual({ kind: 'solid', color: { kind: 'rgb', hex: '#010203' } });
    expect(c.borders?.l).toBeNull();
    expect(c.borders?.r).toMatchObject({ fill: { color: accent(2) } });
    expect(c.text.paragraphs[0].runs[0]).toMatchObject({ bold: false, color: { kind: 'rgb', hex: '#ABCDEF' } });
  });

  it('warns about a style the file does not define, and when there is no style part', async () => {
    const missing = await read({ tree: styled(1, 1, ''), tableStyles: `<a:tblStyleLst ${NS}/>` });
    expect(missing.warnings).toContain(`A table uses the style ${STYLE_ID}, which the file does not define; its cells keep only their own formatting.`);
    expect((missing.slides[0].elements as PptxTable[])[0].styleId).toBe(STYLE_ID);
    const none = await read({ tree: styled(1, 1, '') });
    expect(none.warnings.some((w) => w.includes(STYLE_ID))).toBe(true);
    expect((await read({ tree: table([[cell('A')]], { pr: '<a:tblPr><a:tableStyleId> </a:tableStyleId></a:tblPr>' }) })).warnings.some((w) => w.includes('style'))).toBe(false);
  });

  it('reads the flags that are on', async () => {
    const t = await tableOnly(table([[cell('A')]], { pr: '<a:tblPr firstRow="1" lastCol="true" bandRow="0"/>' }));
    expect(t.flags).toEqual({ firstRow: true, lastCol: true });
  });
});
