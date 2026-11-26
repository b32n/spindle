import type { OoxmlParts } from '../../../../scripts/fidelity-core/ooxml-package';
import {
  PPTX_FIXTURES_DIR,
  PPTX_MANIFEST_OPTIONS,
  loadPptxManifest,
  verifyPptxClaims,
} from '../../../../scripts/pptx-fidelity';
import { CONSTRUCTS } from '../../../../scripts/pptx-fidelity/constructs';
import {
  maxGroupDepth,
  missingSlideParts,
  relationshipCycles,
  relationshipEdges,
  slideSize,
} from '../../../../scripts/pptx-fidelity/structure';

const mk = (xml: Record<string, string>): OoxmlParts => ({
  kind: 'zip',
  raw: new Uint8Array(),
  names: Object.keys(xml),
  xml,
  sizes: {},
  oversized: [],
});
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const rels = (name: string, ...items: Array<[string, string, string]>) => ({
  [name]: `<Relationships>${items
    .map(
      ([id, type, target]) => `<Relationship Id="${id}" Type="${REL}${type}" Target="${target}"/>`
    )
    .join('')}</Relationships>`,
});

describe('maxGroupDepth', () => {
  it('measures nesting and ignores self-closing groups and group properties', () => {
    expect(maxGroupDepth('')).toBe(0);
    expect(maxGroupDepth('<p:grpSp><p:grpSpPr/></p:grpSp>')).toBe(1);
    expect(
      maxGroupDepth('<p:grpSp><p:grpSp><p:grpSp/></p:grpSp></p:grpSp><p:grpSp></p:grpSp>')
    ).toBe(2);
    expect(maxGroupDepth('<p:nvGrpSpPr/><p:grpSp a="1"><p:grpSp>x</p:grpSp></p:grpSp>')).toBe(2);
  });
});

describe('slideSize', () => {
  const size = (attrs: string) =>
    slideSize(
      mk({ 'ppt/presentation.xml': `<p:presentation><p:sldSz ${attrs}/></p:presentation>` })
    );
  it('classifies one kind per deck', () => {
    expect(size('cx="9144000" cy="6858000"')?.kind).toBe('4x3');
    expect(size('cx="12192000" cy="6858000"')?.kind).toBe('16x9');
    expect(size('cx="10692000" cy="7560000" type="A4"')?.kind).toBe('a4');
    expect(size('cx="10692000" cy="7560000"')?.kind).toBe('custom');
    expect(size('cx="7560000" cy="10692000" type="A4"')?.kind).toBe('portrait');
    expect(size('cy="6858000" cx="10000000"')).toEqual({
      cx: 10000000,
      cy: 6858000,
      kind: 'custom',
    });
  });
  it('returns null without a usable size', () => {
    expect(slideSize(mk({}))).toBeNull();
    expect(slideSize(mk({ 'ppt/presentation.xml': '<p:presentation/>' }))).toBeNull();
    expect(size('cy="6858000"')).toBeNull();
    expect(size('cx="0" cy="6858000"')).toBeNull();
    expect(size('cx="9144000" cy="x"')).toBeNull();
  });
});

describe('relationships', () => {
  it('lists internal edges with their type, resolved against the source part', () => {
    const p = mk({
      ...rels('ppt/slides/_rels/slide1.xml.rels', [
        'rId1',
        'slideLayout',
        '../slideLayouts/slideLayout1.xml',
      ]),
      'ppt/_rels/presentation.xml.rels':
        '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="x" Target="https://example.com" TargetMode="External"/></Relationships>',
      'ppt/slides/slide1.xml': '<p:sld/>',
    });
    expect(relationshipEdges(p)).toEqual([
      {
        from: 'ppt/slides/slide1.xml',
        to: 'ppt/slideLayouts/slideLayout1.xml',
        type: `${REL}slideLayout`,
      },
      { from: 'ppt/presentation.xml', to: 'ppt/slides/slide1.xml', type: '' },
    ]);
  });

  it('counts back edges, once per cycle edge, and walks shared targets once', () => {
    const p = mk({
      ...rels('_rels/a.xml.rels', ['r1', 'slide', 'b.xml'], ['r2', 'slide', 'c.xml']),
      ...rels('_rels/b.xml.rels', ['r1', 'slide', 'c.xml']),
      ...rels('_rels/c.xml.rels', ['r1', 'slide', 'a.xml'], ['r2', 'slide', 'c.xml']),
    });
    // a -> b -> c -> a (back), c -> c (back); a -> c is a cross edge to a finished node
    expect(relationshipCycles(p)).toBe(2);
    expect(relationshipCycles(mk({}))).toBe(0);
    const tree = mk({
      ...rels('_rels/a.xml.rels', ['r1', 'slide', 'b.xml'], ['r2', 'slide', 'c.xml']),
      ...rels('_rels/b.xml.rels', ['r1', 'slide', 'c.xml']),
    });
    expect(relationshipCycles(tree)).toBe(0);
  });

  it('does not count the layout to master and notes slide to slide pointers back at the owner', () => {
    const p = mk({
      ...rels('ppt/slideMasters/_rels/slideMaster1.xml.rels', [
        'r1',
        'slideLayout',
        '../slideLayouts/slideLayout1.xml',
      ]),
      ...rels('ppt/slideLayouts/_rels/slideLayout1.xml.rels', [
        'r1',
        'slideMaster',
        '../slideMasters/slideMaster1.xml',
      ]),
      ...rels('ppt/slides/_rels/slide1.xml.rels', [
        'r1',
        'notesSlide',
        '../notesSlides/notesSlide1.xml',
      ]),
      ...rels('ppt/notesSlides/_rels/notesSlide1.xml.rels', [
        'r1',
        'slide',
        '../slides/slide1.xml',
      ]),
    });
    expect(relationshipCycles(p)).toBe(0);
    const other = mk({
      ...rels('ppt/slides/_rels/slide1.xml.rels', ['r1', 'slide', 'slide2.xml']),
      ...rels('ppt/slides/_rels/slide2.xml.rels', ['r1', 'slide', 'slide1.xml']),
    });
    expect(relationshipCycles(other)).toBe(1);
  });
});

describe('missingSlideParts', () => {
  const pres = (ids: string) => `<p:presentation><p:sldIdLst>${ids}</p:sldIdLst></p:presentation>`;
  it('counts listed slides with no relationship, an external one or no part', () => {
    const p = mk({
      'ppt/presentation.xml': pres(
        '<p:sldId id="1" r:id="rId1"/><p:sldId r:id="rId2"/><p:sldId r:id="rId3"/><p:sldId r:id="rId9"/><p:sldId id="5"/>'
      ),
      'ppt/_rels/presentation.xml.rels':
        `<Relationships><Relationship Id="rId1" Type="${REL}slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="${REL}slide" Target="slides/slide2.xml"/>` +
        `<Relationship Id="rId3" Type="${REL}slide" Target="https://example.com" TargetMode="External"/></Relationships>`,
      'ppt/slides/slide1.xml': '<p:sld/>',
    });
    expect(missingSlideParts(p)).toBe(4);
  });
  it('is zero without the presentation or its relationships', () => {
    expect(missingSlideParts(mk({}))).toBe(0);
    expect(missingSlideParts(mk({ 'ppt/presentation.xml': pres('<p:sldId r:id="rId1"/>') }))).toBe(
      0
    );
    expect(missingSlideParts(mk({ 'ppt/_rels/presentation.xml.rels': '<Relationships/>' }))).toBe(
      0
    );
  });
});

describe('bindings', () => {
  it('load the manifest and verify claims with and without an explicit directory', async () => {
    expect(loadPptxManifest(PPTX_FIXTURES_DIR)).toEqual(loadPptxManifest());
    expect(await verifyPptxClaims(PPTX_FIXTURES_DIR)).toEqual([]);
    expect(PPTX_MANIFEST_OPTIONS.extensions).toEqual(['.pptx', '.pptm', '.ppt']);
    expect(PPTX_MANIFEST_OPTIONS.constructIds).toEqual(CONSTRUCTS.map((c) => c.id));
  });
});
