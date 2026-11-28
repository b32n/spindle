import JSZip from 'jszip';
import { DEFAULT_COLOR_MAP } from './color';
import { FALLBACK_THEME } from './theme';
import { exportPptx } from './writer';
import { lintPptxPackage } from './writer-lint';

const base = async (): Promise<JSZip> => {
  const { bytes } = await exportPptx({
    slideSize: { w: 1280, h: 720 },
    theme: FALLBACK_THEME,
    colorMap: DEFAULT_COLOR_MAP,
    slides: [{ pptxId: '256', partName: 'a' }, { pptxId: '257', partName: 'b' }],
    layouts: [],
    fonts: { used: [], unavailable: [] },
    opaqueParts: [],
    warnings: [],
  });
  return JSZip.loadAsync(bytes);
};
const edit = async (name: string, change: (xml: string) => string): Promise<JSZip> => {
  const zip = await base();
  zip.file(name, change((await zip.file(name)?.async('string')) as string));
  return zip;
};

describe('lintPptxPackage', () => {
  it('passes a package the writer made, from bytes or from a zip', async () => {
    const zip = await base();
    expect(await lintPptxPackage(zip)).toEqual([]);
    expect(await lintPptxPackage(await zip.generateAsync({ type: 'uint8array' }))).toEqual([]);
  });

  it('reports a missing or broken content types part', async () => {
    const none = await base();
    none.remove('[Content_Types].xml');
    expect(await lintPptxPackage(none)).toEqual(['[Content_Types].xml is missing or not well-formed']);
    const broken = await edit('[Content_Types].xml', () => '<Types><Default></Types>');
    expect(await lintPptxPackage(broken)).toContain('[Content_Types].xml is missing or not well-formed');
  });

  it('reports a part with no content type and an override for a part that is not there', async () => {
    const zip = await base();
    zip.file('ppt/media/image1.png', 'x');
    zip.file('ppt/orphan.dat', 'x');
    let problems = await lintPptxPackage(zip);
    expect(problems).toContain('ppt/media/image1.png: no content type (no Override, and no Default for .png)');
    expect(problems).toContain('ppt/orphan.dat: no content type (no Override, and no Default for .dat)');
    const gone = await base();
    gone.remove('ppt/viewProps.xml');
    problems = await lintPptxPackage(gone);
    expect(problems).toContain('[Content_Types].xml declares ppt/viewProps.xml, which the package does not have');
    expect(problems.some((p) => p.includes('targets viewProps.xml'))).toBe(true);
  });

  it('reports XML that is not well-formed', async () => {
    const zip = await edit('ppt/slides/slide1.xml', (x) => x.replace('</p:sld>', ''));
    expect(await lintPptxPackage(zip)).toContain('ppt/slides/slide1.xml: not well-formed XML');
  });

  it('reports a duplicate relationship id, an external target that needs no part, and a missing target', async () => {
    const dup = await edit('ppt/slides/_rels/slide1.xml.rels', (x) => x.replace('</Relationships>', '<Relationship Id="rId1" Type="t" Target="https://example.com/" TargetMode="External"/></Relationships>'));
    expect(await lintPptxPackage(dup)).toEqual(['ppt/slides/_rels/slide1.xml.rels: relationship id rId1 is used twice']);
    const ext = await edit('ppt/slides/_rels/slide1.xml.rels', (x) => x.replace('</Relationships>', '<Relationship Id="rId2" Type="t" Target="https://example.com/" TargetMode="External"/></Relationships>'));
    expect(await lintPptxPackage(ext)).toEqual([]);
    const missing = await edit('ppt/slides/_rels/slide1.xml.rels', (x) => x.replace('</Relationships>', '<Relationship Id="rId2" Type="t" Target="../media/none.png"/></Relationships>'));
    expect(await lintPptxPackage(missing)).toEqual(['ppt/slides/_rels/slide1.xml.rels: rId2 targets ../media/none.png, which the package does not have']);
  });

  it('reports a relationship reference the part cannot resolve', async () => {
    const zip = await edit('ppt/slides/slide1.xml', (x) => x.replace('<p:spTree>', '<p:spTree><p:pic><p:blipFill><a:blip r:embed="rId9"/></p:blipFill></p:pic>').replace('<p:nvGrpSpPr>', '<p:nvGrpSpPr>'));
    expect(await lintPptxPackage(zip)).toContain("ppt/slides/slide1.xml: blip@r:embed refers to rId9, which the part's relationships do not have");
  });

  it('reports shape ids that repeat or are not numbers', async () => {
    const twice = await edit('ppt/slides/slide1.xml', (x) => x.replace('</p:spTree>', '<p:sp><p:nvSpPr><p:cNvPr id="1" name="x"/></p:nvSpPr></p:sp></p:spTree>'));
    expect(await lintPptxPackage(twice)).toContain('ppt/slides/slide1.xml: shape id 1 is used twice');
    const text = await edit('ppt/slides/slide1.xml', (x) => x.replace('</p:spTree>', '<p:sp><p:nvSpPr><p:cNvPr id="a" name="x"/></p:nvSpPr></p:sp></p:spTree>'));
    expect(await lintPptxPackage(text)).toContain('ppt/slides/slide1.xml: a shape has no numeric id');
  });

  it('reports a shape tree out of order: header, content, extLst last, and nothing foreign', async () => {
    const header = await edit('ppt/slides/slide1.xml', (x) => x.replace(/<p:nvGrpSpPr>.*<\/p:grpSpPr>/, '<p:grpSpPr/><p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>'));
    expect(await lintPptxPackage(header)).toContain('ppt/slides/slide1.xml: spTree must start with nvGrpSpPr and grpSpPr');
    const ext = await edit('ppt/slides/slide1.xml', (x) => x.replace('</p:spTree>', '<p:extLst/><p:sp/></p:spTree>'));
    expect(await lintPptxPackage(ext)).toContain('ppt/slides/slide1.xml: spTree has children after extLst');
    const foreign = await edit('ppt/slides/slide1.xml', (x) => x.replace('</p:spTree>', '<p:extLst/><p:bogus/></p:spTree>'));
    expect(await lintPptxPackage(foreign)).toContain('ppt/slides/slide1.xml: spTree has an unexpected child bogus');
    const fine = await edit('ppt/slides/slide1.xml', (x) => x.replace('</p:spTree>', '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="g"/></p:nvGrpSpPr><p:grpSpPr/></p:grpSp><p:extLst/></p:spTree>'));
    expect(await lintPptxPackage(fine)).toEqual([]);
    const group = await edit('ppt/slides/slide1.xml', (x) => x.replace('</p:spTree>', '<p:grpSp><p:grpSpPr/></p:grpSp></p:spTree>'));
    expect(await lintPptxPackage(group)).toContain('ppt/slides/slide1.xml: grpSp must start with nvGrpSpPr and grpSpPr');
  });

  it('reports slide, master and layout list ids that are out of range or repeat', async () => {
    const range = await edit('ppt/presentation.xml', (x) => x.replace('id="256"', 'id="12"').replace('id="2147483648"', 'id="5"'));
    expect(await lintPptxPackage(range)).toEqual(['ppt/presentation.xml: sldId id 12 is outside 256..2147483647', 'ppt/presentation.xml: sldMasterId id 5 is outside 2147483648..4294967295']);
    const dup = await edit('ppt/presentation.xml', (x) => x.replace('id="257"', 'id="256"'));
    expect(await lintPptxPackage(dup)).toEqual(['ppt/presentation.xml: sldId id 256 is not unique']);
    const clash = await edit('ppt/slideMasters/slideMaster1.xml', (x) => x.replace('id="2147483649"', 'id="2147483648"'));
    expect(await lintPptxPackage(clash)).toEqual(['ppt/slideMasters/slideMaster1.xml: sldLayoutId id 2147483648 is not unique']);
  });

  it('reports a slide without exactly one layout, a layout without a master and a master without a theme', async () => {
    const twoLayouts = await edit('ppt/slides/_rels/slide1.xml.rels', (x) => x.replace('</Relationships>', '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>'));
    expect(await lintPptxPackage(twoLayouts)).toEqual(['ppt/slides/slide1.xml: needs exactly one slideLayout relationship']);
    const noMaster = await base();
    noMaster.remove('ppt/slideLayouts/_rels/slideLayout1.xml.rels');
    expect(await lintPptxPackage(noMaster)).toEqual(['ppt/slideLayouts/slideLayout1.xml: needs exactly one slideMaster relationship']);
    const noTheme = await edit('ppt/slideMasters/_rels/slideMaster1.xml.rels', (x) => x.replace(/<Relationship[^>]*\/theme"[^>]*\/>/, ''));
    expect(await lintPptxPackage(noTheme)).toEqual(['ppt/slideMasters/slideMaster1.xml: needs exactly one theme relationship']);
  });
});

describe('lintPptxPackage: parts with attributes missing', () => {
  it('treats a missing attribute as empty rather than failing', async () => {
    const zip = await base();
    zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')?.async('string'))?.replace('</Types>', '<Default/><Override/></Types>') as string);
    zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships><Relationship Type="t"/><Relationship Type="t"/></Relationships>');
    zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="p"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>');
    zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="p"><p:sldIdLst><p:sldId/></p:sldIdLst></p:presentation>');
    zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<Relationships><Relationship Id="rId1"/></Relationships>');
    zip.file('ppt/notesSlides/notesSlide1.xml', '<p:notes xmlns:p="p"/>');
    const problems = await lintPptxPackage(zip);
    expect(problems).toEqual(expect.arrayContaining([
      '[Content_Types].xml declares , which the package does not have',
      'ppt/slides/_rels/slide1.xml.rels: relationship id  is used twice',
      'ppt/slides/slide1.xml: a shape has no numeric id',
      'ppt/presentation.xml: sldId id  is outside 256..2147483647',
      'ppt/slideLayouts/slideLayout1.xml: needs exactly one slideMaster relationship',
    ]));
  });
});
