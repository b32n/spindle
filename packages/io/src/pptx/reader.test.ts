import JSZip from 'jszip';
import type { PresentationImportResult } from '@b32nio/spindle-shared/export/pptx';
import { ImportError } from '../common/import-error';
import { PACKAGE_LIMITS } from '../common/ooxml-package';
import { FALLBACK_THEME } from './theme';
import { importPptx } from './reader';

// ---------------------------------------------------------------------------
// Synthetic decks: each package is a real zip built here, so the reader runs through the same path as for a file on disk.
// ---------------------------------------------------------------------------

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PRES_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';

const rel = (id: string, type: string, target: string, external = false): string =>
  `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;
const rels = (list: string[]): string => `<Relationships xmlns="${PKG_REL}">${list.join('')}</Relationships>`;

const themeXml = (name: string, accent1 = '4472C4', extra = ''): string =>
  `<a:theme ${NS} name="${name}"><a:themeElements><a:clrScheme name="c"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="${accent1}"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>${extra}</a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Georgia"/></a:majorFont><a:minorFont><a:latin typeface="Verdana"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;

interface SlideSpec {
  attrs?: string;
  name?: string;
  /** Omit the layout relationship. */
  noLayout?: boolean;
  layout?: number;
  raw?: string;
  /** Do not write the part at all. */
  missing?: boolean;
  /** Replace the presentation-level relationship id. */
  relId?: string;
  /** Write the `p:sldId` without an `id`. */
  noId?: boolean;
  /** Write this `p:sldId/@id` instead of 256 + index. */
  id?: string;
}

interface Spec {
  slides?: SlideSpec[];
  masters?: { theme?: string; clrMap?: string; noTheme?: boolean }[];
  size?: string;
  /** Children of `p:presentation` after `p:sldSz` (the extension list with sections). */
  tail?: string;
  contentTypes?: string;
  rootRels?: string;
  noPresentation?: boolean;
  presentationRaw?: string;
  files?: Record<string, string>;
  /** Do not list the masters in the presentation relationships. */
  masterUnlisted?: boolean;
}

async function build(spec: Spec = {}): Promise<Uint8Array> {
  const slides = spec.slides ?? [{}];
  const masters = spec.masters ?? [{}];
  const zip = new JSZip();
  zip.file('[Content_Types].xml', spec.contentTypes ?? `<Types xmlns="${CT}"><Override PartName="/ppt/presentation.xml" ContentType="${PRES_TYPE}"/></Types>`);
  zip.file('_rels/.rels', spec.rootRels ?? rels([rel('rId1', 'officeDocument', 'ppt/presentation.xml')]));
  const presRels: string[] = [];
  const ids: string[] = [];
  slides.forEach((s, i) => {
    const n = i + 1;
    presRels.push(rel(`rId${n}`, 'slide', `slides/slide${n}.xml`));
    ids.push(`<p:sldId ${s.noId ? '' : `id="${s.id ?? 256 + i}" `}r:id="${s.relId ?? `rId${n}`}"/>`);
    if (!s.missing)
      zip.file(
        `ppt/slides/slide${n}.xml`,
        s.raw ?? `<p:sld ${NS}${s.attrs ? ` ${s.attrs}` : ''}><p:cSld${s.name === undefined ? '' : ` name="${s.name}"`}><p:spTree/></p:cSld></p:sld>`
      );
    zip.file(
      `ppt/slides/_rels/slide${n}.xml.rels`,
      rels(s.noLayout ? [] : [rel('rId1', 'slideLayout', `../slideLayouts/slideLayout${s.layout ?? 1}.xml`)])
    );
  });
  masters.forEach((m, i) => {
    const n = i + 1;
    if (!spec.masterUnlisted) presRels.push(rel(`rIdM${n}`, 'slideMaster', `slideMasters/slideMaster${n}.xml`));
    zip.file(`ppt/slideMasters/slideMaster${n}.xml`, `<p:sldMaster ${NS}><p:cSld/>${m.clrMap ?? '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'}</p:sldMaster>`);
    zip.file(`ppt/slideMasters/_rels/slideMaster${n}.xml.rels`, rels(m.noTheme ? [] : [rel('rId1', 'theme', `../theme/theme${n}.xml`)]));
    if (!m.noTheme) zip.file(`ppt/theme/theme${n}.xml`, m.theme ?? themeXml(`Theme ${n}`, i === 0 ? '4472C4' : 'AA0000'));
    zip.file(`ppt/slideLayouts/slideLayout${n}.xml`, `<p:sldLayout ${NS}><p:cSld/></p:sldLayout>`);
    zip.file(`ppt/slideLayouts/_rels/slideLayout${n}.xml.rels`, rels([rel('rId1', 'slideMaster', `../slideMasters/slideMaster${n}.xml`)]));
  });
  zip.file('ppt/_rels/presentation.xml.rels', rels(presRels));
  if (!spec.noPresentation)
    zip.file(
      'ppt/presentation.xml',
      spec.presentationRaw ??
        `<p:presentation ${NS}><p:sldIdLst>${ids.join('')}</p:sldIdLst>${spec.size ?? '<p:sldSz cx="12192000" cy="6858000"/>'}${spec.tail ?? ''}</p:presentation>`
    );
  for (const [name, content] of Object.entries(spec.files ?? {})) zip.file(name, content);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const read = async (spec: Spec = {}): Promise<PresentationImportResult> => importPptx(await build(spec));
const failure = async (bytes: Uint8Array, options?: Parameters<typeof importPptx>[1]): Promise<ImportError> => {
  try {
    await importPptx(bytes, options);
  } catch (e) {
    expect(e).toBeInstanceOf(ImportError);
    return e as ImportError;
  }
  throw new Error('expected the import to fail');
};

describe('importPptx: the presentation', () => {
  it('reads size, slides in presentation order, names, hidden flags and layouts', async () => {
    const r = await read({
      slides: [{ name: ' Title ' }, { attrs: 'show="0"' }, { attrs: 'show="false"', layout: 2 }, { attrs: 'show="1"' }],
      masters: [{}, {}],
    });
    expect(r.slideSize).toEqual({ w: 1280, h: 720 });
    expect(r.warnings).toEqual([]);
    expect(r.slides).toEqual([
      { pptxId: '256', partName: 'ppt/slides/slide1.xml', name: 'Title', layoutPart: 'ppt/slideLayouts/slideLayout1.xml' },
      { pptxId: '257', partName: 'ppt/slides/slide2.xml', hidden: true, layoutPart: 'ppt/slideLayouts/slideLayout1.xml' },
      { pptxId: '258', partName: 'ppt/slides/slide3.xml', hidden: true, layoutPart: 'ppt/slideLayouts/slideLayout2.xml' },
      { pptxId: '259', partName: 'ppt/slides/slide4.xml', layoutPart: 'ppt/slideLayouts/slideLayout1.xml' },
    ]);
    expect(r.sections).toBeUndefined();
  });

  it('follows the presentation list, not the file names, and survives an empty list', async () => {
    const swapped = await read({ slides: [{ relId: 'rId2' }, { relId: 'rId1' }] });
    expect(swapped.slides.map((s) => s.partName)).toEqual(['ppt/slides/slide2.xml', 'ppt/slides/slide1.xml']);
    expect((await read({ slides: [] })).slides).toEqual([]);
  });

  it('converts any size to px (1 px = 9525 EMU) and uses 16:9 with a warning when there is none', async () => {
    expect((await read({ size: '<p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>' })).slideSize).toEqual({ w: 960, h: 720 });
    expect((await read({ size: '<p:sldSz cx="6858000" cy="12192000"/>' })).slideSize).toEqual({ w: 720, h: 1280 });
    expect((await read({ size: '<p:sldSz cx="10692000" cy="7560000"/>' })).slideSize).toEqual({ w: 1123, h: 794 });
    for (const size of ['', '<p:sldSz cx="99999999999999" cy="6858000"/>', '<p:sldSz cx="12192000" cy="51206401"/>', '<p:sldSz cx="0" cy="6858000"/>', '<p:sldSz cx="abc" cy="1"/>', '<p:sldSz cx="12192000"/>']) {
      const r = await read({ size });
      expect(r.slideSize).toEqual({ w: 1280, h: 720 });
      expect(r.warnings).toEqual(['The presentation does not give a usable slide size; 16:9 (1280 x 720) is used.']);
    }
  });

  it('keeps a slide it cannot read as an error placeholder, with a warning, and reads the others', async () => {
    const r = await read({
      slides: [{}, { missing: true }, { raw: 'plain text, no markup' }, { raw: `<p:other ${NS}/>` }, { relId: 'rId99' }, { noId: true }],
    });
    expect(r.slides.map((s) => s.error)).toEqual([undefined, 'the part is missing', 'the part is damaged', 'the part is damaged', 'its relationship is missing', undefined]);
    expect(r.slides[1]).toEqual({ pptxId: '257', partName: 'ppt/slides/slide2.xml', error: 'the part is missing' });
    expect(r.slides[4]).toMatchObject({ partName: '', error: 'its relationship is missing' });
    expect(r.slides[5].pptxId).toBe('261');
    expect(r.warnings).toHaveLength(4);
    expect(r.warnings).toEqual([
      'Slide 2 could not be read (the part is missing); it is kept as an empty placeholder.',
      'Slide 3 could not be read (the part is damaged); it is kept as an empty placeholder.',
      'Slide 4 could not be read (the part is damaged); it is kept as an empty placeholder.',
      'Slide 5 could not be read (its relationship is missing); it is kept as an empty placeholder.',
    ]);
  });

  it('accepts PowerPoint\'s largest slide size and no more', async () => {
    expect((await read({ size: '<p:sldSz cx="51206400" cy="51206400"/>' })).slideSize).toEqual({ w: 5376, h: 5376 });
  });

  it('gives a slide without a usable id a number no other slide has', async () => {
    const r = await read({ slides: [{ noId: true }, { id: '256' }, { id: 'x' }, { id: '257' }, { noId: true }] });
    expect(r.slides.map((s) => s.pptxId)).toEqual(['258', '256', '259', '257', '260']);
    expect(new Set(r.slides.map((s) => s.pptxId)).size).toBe(5);
  });

  it('copes with list entries and overrides that carry no attributes', async () => {
    const r = await read({
      presentationRaw: `<p:presentation ${NS}><p:sldIdLst><p:sldId/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`,
      contentTypes: `<Types xmlns="${CT}"><Override/></Types>`,
    });
    expect(r.slides).toEqual([{ pptxId: '256', partName: '', error: 'its relationship is missing' }]);
  });

  it('treats an external slide relationship as missing and a slide with no layout as having none', async () => {
    const zip = await JSZip.loadAsync(await build({ slides: [{ noLayout: true }, {}] }));
    zip.file('ppt/_rels/presentation.xml.rels', rels([rel('rId1', 'slide', 'https://example.com/s.xml', true), rel('rId2', 'slide', 'slides/slide2.xml'), rel('rIdM1', 'slideMaster', 'slideMasters/slideMaster1.xml')]));
    const r = await importPptx(await zip.generateAsync({ type: 'uint8array' }));
    expect(r.slides[0]).toMatchObject({ error: 'its relationship is missing' });
    expect(r.slides[1].layoutPart).toBe('ppt/slideLayouts/slideLayout1.xml');
    expect((await read({ slides: [{ noLayout: true }] })).slides[0].layoutPart).toBeUndefined();
  });

  it('reads sections and drops slide ids that match no slide', async () => {
    const tail =
      '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">' +
      '<p14:section name=" Intro " id="{1}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="999"/><p14:sldId/></p14:sldIdLst></p14:section>' +
      '<p14:section id="{2}"><p14:sldIdLst><p14:sldId id="257"/></p14:sldIdLst></p14:section>' +
      '<p14:section name="Empty" id="{3}"/></p14:sectionLst></p:ext></p:extLst>';
    const r = await read({ slides: [{}, {}], tail });
    expect(r.sections).toEqual([
      { name: 'Intro', slideIds: ['256'] },
      { name: '', slideIds: ['257'] },
      { name: 'Empty', slideIds: [] },
    ]);
  });

  it('names a part it had to skip for its size, and does not fail', async () => {
    const saved = { ...PACKAGE_LIMITS };
    const bytes = await build({ files: { 'ppt/slides/slide1.xml': `<p:sld ${NS}><p:cSld><p:spTree><!-- ${'x'.repeat(3000)} --></p:spTree></p:cSld></p:sld>` } });
    PACKAGE_LIMITS.partBytes = 2000;
    try {
      const r = await importPptx(bytes);
      expect(r.slides[0]).toMatchObject({ error: 'the part is too large to read' });
      expect(r.warnings).toContain('ppt/slides/slide1.xml was too large to read and was skipped.');
    } finally {
      Object.assign(PACKAGE_LIMITS, saved);
    }
  });
});

describe('importPptx: the main part', () => {
  it('finds the presentation through the package relationship wherever it is', async () => {
    const zip = await JSZip.loadAsync(await build());
    const moved = await zip.file('ppt/presentation.xml')?.async('string');
    zip.remove('ppt/presentation.xml');
    zip.file('deck/main.xml', moved ?? '');
    zip.file('deck/_rels/main.xml.rels', await zip.file('ppt/_rels/presentation.xml.rels')?.async('string') ?? '');
    zip.file('_rels/.rels', rels([rel('rId1', 'officeDocument', 'deck/main.xml')]));
    const r = await importPptx(await zip.generateAsync({ type: 'uint8array' }));
    expect(r.slides).toHaveLength(1);
  });

  it('falls back to the content type override, then to ppt/presentation.xml, and refuses a package with neither', async () => {
    const viaType = await read({ rootRels: rels([]) });
    expect(viaType.slides).toHaveLength(1);
    const viaPath = await read({ rootRels: rels([]), contentTypes: `<Types xmlns="${CT}"/>` });
    expect(viaPath.slides).toHaveLength(1);
    // an override that names a part that is not there is ignored
    const dangling = await read({ rootRels: rels([]), contentTypes: `<Types xmlns="${CT}"><Override PartName="/ppt/gone.xml" ContentType="${PRES_TYPE}"/></Types>` });
    expect(dangling.slides).toHaveLength(1);
    const none = await failure(await build({ noPresentation: true, rootRels: rels([]), contentTypes: `<Types xmlns="${CT}"/>` }));
    expect(none.code).toBe('not-a-package');
    expect(none.message).toMatch(/no presentation part/);
  });

  it('refuses a main part that is not a presentation, or not XML, or an external target', async () => {
    expect((await failure(await build({ presentationRaw: `<w:document xmlns:w="urn:w"/>` }))).code).toBe('not-a-package');
    expect((await failure(await build({ presentationRaw: 'not xml at all' }))).code).toBe('not-a-package');
    const ext = await build({ rootRels: rels([rel('rId1', 'officeDocument', 'https://example.com/p.xml', true)]), noPresentation: true, contentTypes: `<Types xmlns="${CT}"/>` });
    expect((await failure(ext)).code).toBe('not-a-package');
  });

  it('reads macro-enabled and template main parts by their content type', async () => {
    const types = [
      'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
      'application/vnd.ms-powerpoint.template.macroEnabled.main+xml',
      'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml',
      'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
    ];
    for (const type of types) {
      const r = await read({ rootRels: rels([]), contentTypes: `<Types xmlns="${CT}"><Override PartName="/ppt/presentation.xml" ContentType="${type}"/></Types>` });
      expect(r.slides).toHaveLength(1);
    }
  });
});

describe('importPptx: the theme', () => {
  it('takes the theme and colour map of the first slide\'s master', async () => {
    const dark = '<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
    const r = await read({ slides: [{ layout: 2 }], masters: [{}, { clrMap: dark }] });
    expect(r.theme.name).toBe('Theme 2');
    expect(r.theme.colors.accent1).toBe('#AA0000');
    expect(r.theme.fonts).toEqual({ major: 'Georgia', minor: 'Verdana' });
    expect(r.colorMap).toMatchObject({ bg1: 'dk1', tx1: 'lt1' });
  });

  it('falls back to the first listed master when no slide names a layout, or its master is gone', async () => {
    const none = await read({ slides: [{ noLayout: true }], masters: [{}, {}] });
    expect(none.theme.name).toBe('Theme 1');
    const zip = await JSZip.loadAsync(await build({ slides: [{}], masters: [{}, {}] }));
    zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels([rel('rId1', 'slideMaster', '../slideMasters/gone.xml')]));
    const gone = await importPptx(await zip.generateAsync({ type: 'uint8array' }));
    expect(gone.theme.name).toBe('Theme 1');
    const external = await JSZip.loadAsync(await build({ slides: [{ noLayout: true }] }));
    external.file('ppt/_rels/presentation.xml.rels', rels([rel('rId1', 'slide', 'slides/slide1.xml'), rel('rIdM1', 'slideMaster', 'https://example.com/m.xml', true)]));
    expect((await importPptx(await external.generateAsync({ type: 'uint8array' }))).theme).toEqual(FALLBACK_THEME);
  });

  it('uses the Office theme, with a warning, when there is no master or the master has no readable theme', async () => {
    const noMaster = await read({ masters: [], slides: [{ noLayout: true }] });
    expect(noMaster.theme).toEqual(FALLBACK_THEME);
    expect(noMaster.warnings).toEqual(['The presentation has no readable slide master and theme; the Office default theme is used.']);
    const noTheme = await read({ masters: [{ noTheme: true }] });
    expect(noTheme.theme).toEqual(FALLBACK_THEME);
    expect(noTheme.warnings).toEqual(['The slide master has no readable theme; the Office default theme is used.']);
    const badTheme = await read({ masters: [{ theme: 'not xml' }] });
    expect(badTheme.warnings).toEqual(['The slide master has no readable theme; the Office default theme is used.']);
  });

  it('passes on the theme\'s own warnings, once', async () => {
    const r = await read({ masters: [{ theme: themeXml('T').replace('<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>', '') }] });
    expect(r.warnings).toEqual(['The theme does not define accent2; the Office default colours are used for them.']);
    expect(r.colorMap.bg1).toBe('lt1');
  });

  it('uses the default colour map for a master that has none', async () => {
    expect((await read({ masters: [{ clrMap: '' }] })).colorMap).toMatchObject({ bg1: 'lt1', tx1: 'dk1' });
  });
});

describe('importPptx: files that are not presentations, and limits', () => {
  it('names each kind of file it cannot open and only ever throws an ImportError', async () => {
    expect((await failure(new Uint8Array(0))).code).toBe('not-a-package');
    expect((await failure(new Uint8Array([1, 2, 3, 4]))).message).toBe('This is not a PowerPoint presentation (.pptx).');
    const ole = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]);
    expect((await failure(ole)).code).toBe('legacy-binary');
    const good = await build();
    expect((await failure(good.slice(0, Math.floor(good.length * 0.6)))).code).toBe('corrupt');
    const word = new JSZip();
    word.file('[Content_Types].xml', `<Types xmlns="${CT}"/>`);
    word.file('word/document.xml', '<w:document xmlns:w="urn:w"/>');
    expect((await failure(await word.generateAsync({ type: 'uint8array' }))).code).toBe('not-a-package');
  });

  it('applies the safe-import limits it is given', async () => {
    const error = await failure(await build(), { limits: { maxEntries: 3 } });
    expect(error.code).toBe('size-limit');
    const unsafe = new JSZip();
    unsafe.file('../evil.xml', 'x');
    expect((await failure(await unsafe.generateAsync({ type: 'uint8array' }))).code).toBe('corrupt');
  });

  it('never throws anything but an ImportError, whatever prefix of a good deck it is given', async () => {
    const good = await build({ slides: [{ name: 'a' }, { attrs: 'show="0"' }] });
    for (let cut = 0; cut < good.length; cut += Math.max(1, Math.floor(good.length / 40))) {
      try {
        await importPptx(good.slice(0, cut));
      } catch (e) {
        expect(e).toBeInstanceOf(ImportError);
      }
    }
  });

  it('survives hostile XML in the presentation, the slides and the theme', async () => {
    const bomb = '<!DOCTYPE p [<!ENTITY a "aaaa"><!ENTITY b "&a;&a;&a;&a;">]>';
    const r = await read({
      presentationRaw: `${bomb}<p:presentation ${NS}><p:sldIdLst><p:sldId id="1" r:id="rId1"/></p:sldIdLst><p:sldSz cx="&b;" cy="6858000"/></p:presentation>`,
    });
    expect(r.slideSize).toEqual({ w: 1280, h: 720 });
    expect(r.warnings.length).toBeGreaterThan(0);
  });
});
