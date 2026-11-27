import type { OoxmlParts } from '../../../../scripts/fidelity-core/ooxml-package';
import type { Json } from '../../../../scripts/fidelity-core/types';
import { modelRuns, sourceRuns, textParity } from '../../../../scripts/pptx-fidelity/parity';

const parts = (xml: Record<string, string>): OoxmlParts => ({ kind: 'zip', raw: new Uint8Array(), names: Object.keys(xml), xml, sizes: {}, oversized: [] });
const model = (...runs: string[][]): Json => ({ slides: runs.map((r) => ({ elements: [{ text: { paragraphs: [{ runs: r.map((text) => ({ kind: 'text', text })) }] } }] })) });

describe('sourceRuns', () => {
  it('finds every a:t of the shapes the reader reads, decoding references and CDATA and skipping empty runs', () => {
    const xml =
      '<p:sld><!-- <a:t>no</a:t> --><p:sp><a:p><a:r><a:t>A &amp; B &#65;&#x42; &lt;</a:t></a:r><a:r><a:t></a:t></a:r><a:r><a:t/></a:r>' +
      '<a:r><a:t x="y"><![CDATA[<raw> &amp;]]> tail</a:t></a:r></a:p></p:sp></p:sld>';
    expect(sourceRuns(xml)).toEqual({ inScope: ['A & B AB <', '<raw> &amp; tail'], outOfScope: [] });
  });

  it('counts the runs of groups, pictures, graphic frames, connectors and fallback copies as out of scope, and resumes after them', () => {
    const xml =
      '<p:spTree><p:sp><a:t>one</a:t></p:sp><p:grpSp><p:sp><a:t>in group</a:t></p:sp></p:grpSp>' +
      '<p:graphicFrame><a:tbl><a:t>cell</a:t></a:tbl></p:graphicFrame><p:pic><a:t>pic</a:t></p:pic><p:cxnSp><a:t>c</a:t></p:cxnSp>' +
      '<mc:AlternateContent><mc:Choice><p:sp><a:t>two</a:t></p:sp></mc:Choice><mc:Fallback><p:pic><a:t>copy</a:t></p:pic></mc:Fallback></mc:AlternateContent>' +
      '<p:sp><a:t>three</a:t></p:sp></p:spTree>';
    expect(sourceRuns(xml)).toEqual({ inScope: ['one', 'two', 'three'], outOfScope: ['in group', 'cell', 'pic', 'c', 'copy'] });
  });

  it('treats an a:t outside every shape as in scope, so the parity check reports it as missing', () => {
    expect(sourceRuns('<p:spTree><a:t>stray</a:t></p:spTree>')).toEqual({ inScope: ['stray'], outOfScope: [] });
    const xml = { 'ppt/slides/slide1.xml': '<p:sld><p:spTree><a:t>stray</a:t></p:spTree></p:sld>' };
    expect(textParity(parts(xml), model([])).missing).toEqual(['stray']);
  });

  it('survives unbalanced tags and closing tags it never saw opened', () => {
    expect(sourceRuns('</p:grpSp><p:sp><a:t>x</a:t></p:sp></p:nope><a:t>y')).toEqual({ inScope: ['x'], outOfScope: [] });
    expect(sourceRuns('<p:grpSp><a:t>x</a:t>')).toEqual({ inScope: [], outOfScope: ['x'] });
  });
});

describe('modelRuns', () => {
  it('collects text and field runs of every element and ignores breaks, empty runs and malformed input', () => {
    const m = { slides: [{ elements: [{ text: { paragraphs: [{ runs: [{ kind: 'text', text: 'a' }, { kind: 'field', text: '1' }, { kind: 'break', text: '' }, { kind: 'text', text: '' }, { kind: 'x', text: 'z' }] }, {}] } }, {}, null] }, {}, { elements: 3 }] };
    expect(modelRuns(m as unknown as Json)).toEqual(['a', '1']);
    expect(modelRuns({})).toEqual([]);
    expect(modelRuns({ slides: [{ elements: [{ text: {} }] }] })).toEqual([]);
  });
});

describe('textParity', () => {
  const slide = (...t: string[]) => `<p:sld>${t.map((x) => `<p:sp><a:t>${x}</a:t></p:sp>`).join('')}<p:graphicFrame><a:t>cell</a:t></p:graphicFrame></p:sld>`;

  it('is exact when the model holds every run once, whatever the order, and ignores parts that are not slides', () => {
    const p = parts({ 'ppt/slides/slide2.xml': slide('b', 'a', 'a'), 'ppt/slides/slide1.xml': slide('c'), 'ppt/slideLayouts/slideLayout1.xml': slide('layout') });
    expect(textParity(p, model(['a', 'c'], ['a', 'b']))).toEqual({ sourceRuns: 4, modelRuns: 4, outOfScope: 2, inErrorSlides: 0, missing: [], extra: [] });
  });

  it('sets aside the runs of slides the model reports as error placeholders, and tolerates a malformed model', () => {
    const p = parts({ 'ppt/slides/slide1.xml': slide('a'), 'ppt/slides/slide2.xml': slide('b', 'c') });
    const m = { slides: [{ partName: 'ppt/slides/slide1.xml', elements: [{ text: { paragraphs: [{ runs: [{ kind: 'text', text: 'a' }] }] } }] }, { partName: 'ppt/slides/slide2.xml', error: 'the part is damaged' }, null, { error: 'x' }] } as unknown as Json;
    expect(textParity(p, m)).toEqual({ sourceRuns: 1, modelRuns: 1, outOfScope: 1, inErrorSlides: 3, missing: [], extra: [] });
    expect(textParity(p, {}).inErrorSlides).toBe(0);
  });

  it('lists lost, invented and duplicated runs (as a multiset), at most twenty of each', () => {
    const p = parts({ 'ppt/slides/slide1.xml': slide('a', 'a', 'b') });
    expect(textParity(p, model(['a', 'x']))).toMatchObject({ missing: ['a', 'b'], extra: ['x'] });
    const many = Array.from({ length: 30 }, (_, i) => `t${i}`);
    expect(textParity(parts({ 'ppt/slides/slide1.xml': slide(...many) }), model([])).missing).toHaveLength(20);
  });
});
