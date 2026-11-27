import type { OoxmlParts } from '../../../../scripts/fidelity-core/ooxml-package';
import type { Json } from '../../../../scripts/fidelity-core/types';
import { modelRuns, notesRuns, sourceRuns, textParity } from '../../../../scripts/pptx-fidelity/parity';

const parts = (xml: Record<string, string>): OoxmlParts => ({ kind: 'zip', raw: new Uint8Array(), names: Object.keys(xml), xml, sizes: {}, oversized: [] });
const model = (...runs: string[][]): Json => ({ slides: runs.map((r) => ({ elements: [{ text: { paragraphs: [{ runs: r.map((text) => ({ kind: 'text', text })) }] } }] })) });

describe('sourceRuns', () => {
  it('finds every a:t of the shapes the reader reads, decoding references and CDATA and skipping empty runs', () => {
    const xml =
      '<p:sld><!-- <a:t>no</a:t> --><p:sp><a:p><a:r><a:t>A &amp; B &#65;&#x42; &lt;</a:t></a:r><a:r><a:t></a:t></a:r><a:r><a:t/></a:r>' +
      '<a:r><a:t x="y"><![CDATA[<raw> &amp;]]> tail</a:t></a:r></a:p></p:sp></p:sld>';
    expect(sourceRuns(xml)).toEqual({ inScope: ['A & B AB <', '<raw> &amp; tail'], outOfScope: [] });
  });

  it('counts the runs of fallback copies as out of scope (groups and connectors are in scope since K10b), and resumes after them', () => {
    const xml =
      '<p:spTree><p:sp><a:t>one</a:t></p:sp><p:grpSp><p:sp><a:t>in group</a:t></p:sp></p:grpSp>' +
      '<p:graphicFrame><a:tbl><a:t>cell</a:t></a:tbl></p:graphicFrame><p:pic><a:t>pic</a:t></p:pic><p:cxnSp><a:t>c</a:t></p:cxnSp>' +
      '<mc:AlternateContent><mc:Choice><p:sp><a:t>two</a:t></p:sp></mc:Choice><mc:Fallback><p:pic><a:t>copy</a:t></p:pic></mc:Fallback></mc:AlternateContent>' +
      '<p:sp><a:t>three</a:t></p:sp></p:spTree>';
    expect(sourceRuns(xml)).toEqual({ inScope: ['one', 'in group', 'cell', 'pic', 'c', 'two', 'three'], outOfScope: ['copy'] });
  });

  it('reads a Fallback as in scope when the alternate content has no Choice, as the reader does', () => {
    const only = '<mc:AlternateContent><mc:Fallback><p:sp><a:t>fb</a:t></p:sp></mc:Fallback></mc:AlternateContent>';
    expect(sourceRuns(only)).toEqual({ inScope: ['fb'], outOfScope: [] });
    // a Choice of another alternate content does not make this Fallback a copy
    const two = `<mc:AlternateContent><mc:Choice><p:sp><a:t>c</a:t></p:sp></mc:Choice></mc:AlternateContent>${only}`;
    expect(sourceRuns(two)).toEqual({ inScope: ['c', 'fb'], outOfScope: [] });
    // a Choice or Fallback outside alternate content marks nothing out of scope
    expect(sourceRuns('<mc:Choice/><mc:Fallback><p:sp><a:t>x</a:t></p:sp></mc:Fallback>')).toEqual({ inScope: ['x'], outOfScope: [] });
    const xml = { 'ppt/slides/slide1.xml': `<p:sld>${only}</p:sld>` };
    expect(textParity(parts(xml), model(['fb']))).toMatchObject({ sourceRuns: 1, outOfScope: 0, missing: [], extra: [] });
  });

  it('treats an a:t outside every shape as in scope, so the parity check reports it as missing', () => {
    expect(sourceRuns('<p:spTree><a:t>stray</a:t></p:spTree>')).toEqual({ inScope: ['stray'], outOfScope: [] });
    const xml = { 'ppt/slides/slide1.xml': '<p:sld><p:spTree><a:t>stray</a:t></p:spTree></p:sld>' };
    expect(textParity(parts(xml), model([])).missing).toEqual(['stray']);
  });

  it('survives unbalanced tags and closing tags it never saw opened', () => {
    expect(sourceRuns('</p:grpSp><p:sp><a:t>x</a:t></p:sp></p:nope><a:t>y')).toEqual({ inScope: ['x'], outOfScope: [] });
    expect(sourceRuns('<p:grpSp><a:t>x</a:t>')).toEqual({ inScope: ['x'], outOfScope: [] });
    expect(sourceRuns('<mc:AlternateContent><mc:Choice></mc:Choice><mc:Fallback><a:t>x</a:t>')).toEqual({ inScope: [], outOfScope: ['x'] });
  });
});

describe('modelRuns', () => {
  it('collects text and field runs of every element and ignores breaks, empty runs and malformed input', () => {
    const m = { slides: [{ elements: [{ text: { paragraphs: [{ runs: [{ kind: 'text', text: 'a' }, { kind: 'field', text: '1' }, { kind: 'break', text: '' }, { kind: 'text', text: '' }, { kind: 'x', text: 'z' }] }, {}] } }, {}, null] }, {}, { elements: 3 }] };
    expect(modelRuns(m as unknown as Json)).toEqual(['a', '1']);
    expect(modelRuns({})).toEqual([]);
    expect(modelRuns({ slides: [{ elements: [{ text: {} }] }] })).toEqual([]);
  });

  it('collects the runs inside groups, at any depth', () => {
    const run = (t: string) => ({ text: { paragraphs: [{ runs: [{ kind: 'text', text: t }] }] } });
    const m = { slides: [{ elements: [{ kind: 'group', children: [run('a'), { kind: 'group', children: [run('b')] }, { kind: 'group' }] }, run('c')] }] };
    expect(modelRuns(m as unknown as Json)).toEqual(['a', 'b', 'c']);
  });

  it('collects the runs of table cells as well, and ignores a table without rows or cells', () => {
    const cell = (t: string) => ({ text: { paragraphs: [{ runs: [{ kind: 'text', text: t }] }] } });
    const m = { slides: [{ elements: [{ kind: 'table', rows: [{ cells: [cell('a'), cell('b')] }, { cells: [cell('c')] }, {}] }, { kind: 'table' }] }] };
    expect(modelRuns(m as unknown as Json)).toEqual(['a', 'b', 'c']);
  });
});

describe('textParity', () => {
  const slide = (...t: string[]) => `<p:sld>${t.map((x) => `<p:sp><a:t>${x}</a:t></p:sp>`).join('')}<mc:AlternateContent><mc:Choice></mc:Choice><mc:Fallback><a:t>copy</a:t></mc:Fallback></mc:AlternateContent></p:sld>`;

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

describe('speaker notes', () => {
  const sp = (type: string, ...t: string[]) => `<p:sp><p:nvSpPr><p:nvPr><p:ph type="${type}"/></p:nvPr></p:nvSpPr><p:txBody>${t.map((x) => `<a:p><a:r><a:t>${x}</a:t></a:r></a:p>`).join('')}</p:txBody></p:sp>`;
  const page = (...shapes: string[]) => `<p:notes><p:cSld><p:spTree>${shapes.join('')}</p:spTree></p:cSld></p:notes>`;
  const rels = (...targets: string[]) => `<Relationships>${targets.join('')}</Relationships>`;
  const notesRel = (target: string, extra = '') => `<Relationship Id="rN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="${target}"${extra}/>`;
  const withNotes = (...t: string[]) => ({ slides: [{ notes: { paragraphs: t.map((x) => ({ runs: [{ kind: 'text', text: x }] })) } }] }) as unknown as Json;

  it('notesRuns takes the runs of the body placeholders as in scope and the rest of the page as out of scope', () => {
    const xml = page(sp('body', 'a', 'b'), sp('sldNum', '1'), sp('hdr', 'h'), sp('body', 'c'));
    expect(notesRuns(xml)).toEqual({ inScope: ['a', 'b', 'c'], outOfScope: ['1', 'h'] });
    expect(notesRuns(page(sp('sldImg')))).toEqual({ inScope: [], outOfScope: [] });
    expect(notesRuns(page())).toEqual({ inScope: [], outOfScope: [] });
    expect(notesRuns('<p:notes><p:sp><p:ph  type = \'body\'/><a:t>q</a:t></p:sp></p:notes>').inScope).toEqual(['q']);
  });

  it('modelRuns counts the runs of a slide\'s notes', () => {
    const m = { slides: [{ elements: [], notes: { paragraphs: [{ runs: [{ kind: 'text', text: 'n' }] }] } }, { notes: null }] };
    expect(modelRuns(m as unknown as Json)).toEqual(['n']);
  });

  it('compares the notes body runs of the page each slide points at with the model, and sets the rest aside', () => {
    const p = parts({
      'ppt/slides/slide1.xml': '<p:sld/>',
      'ppt/slides/_rels/slide1.xml.rels': rels(notesRel('../notesSlides/notesSlide1.xml')),
      'ppt/notesSlides/notesSlide1.xml': page(sp('body', 'x', 'y'), sp('ftr', 'footer')),
      'ppt/notesSlides/notesSlide9.xml': page(sp('body', 'orphan')),
    });
    expect(textParity(p, withNotes('y', 'x'))).toEqual({ sourceRuns: 2, modelRuns: 2, outOfScope: 2, inErrorSlides: 0, missing: [], extra: [] });
    expect(textParity(p, withNotes('x'))).toMatchObject({ missing: ['y'] });
    expect(textParity(p, withNotes('x', 'y', 'z'))).toMatchObject({ extra: ['z'] });
  });

  it('counts the notes of a page several slides share once for each slide that holds it (a failed slide holds none)', () => {
    const slideRels = (n: number) => ({ [`ppt/slides/slide${n}.xml`]: '<p:sld/>', [`ppt/slides/_rels/slide${n}.xml.rels`]: rels(notesRel('../notesSlides/notesSlide1.xml')) });
    const notes = { 'ppt/notesSlides/notesSlide1.xml': page(sp('body', 'a', 'b'), sp('hdr', 'h')) };
    const held = { notes: { paragraphs: [{ runs: [{ kind: 'text', text: 'a' }] }, { runs: [{ kind: 'text', text: 'b' }] }] } };
    const both = { slides: [{ partName: 'ppt/slides/slide1.xml', ...held }, { partName: 'ppt/slides/slide2.xml', ...held }] } as unknown as Json;
    expect(textParity(parts({ ...slideRels(1), ...slideRels(2), ...notes }), both)).toMatchObject({ sourceRuns: 4, modelRuns: 4, missing: [], extra: [], outOfScope: 1 });
    const oneFailed = { slides: [{ partName: 'ppt/slides/slide1.xml', ...held }, { partName: 'ppt/slides/slide2.xml', error: 'x' }] } as unknown as Json;
    expect(textParity(parts({ ...slideRels(1), ...slideRels(2), ...notes }), oneFailed)).toMatchObject({ sourceRuns: 2, modelRuns: 2, missing: [], extra: [] });
  });

  it('sets aside the notes of a slide that failed, resolves an absolute target and ignores an external or foreign relationship', () => {
    const slideRels = (r: string) => ({ 'ppt/slides/slide1.xml': '<p:sld/>', 'ppt/slides/_rels/slide1.xml.rels': rels(r) });
    const failed = { slides: [{ partName: 'ppt/slides/slide1.xml', error: 'x' }] } as unknown as Json;
    const notes = { 'ppt/notesSlides/notesSlide1.xml': page(sp('body', 'a', 'b'), sp('hdr', 'h')) };
    expect(textParity(parts({ ...slideRels(notesRel('/ppt/notesSlides/notesSlide1.xml')), ...notes }), failed)).toMatchObject({ sourceRuns: 0, inErrorSlides: 3 });
    // external, other relationship types and a missing rels part leave the page unreachable: out of scope
    const foreign = '<Relationship Id="r" Type="http://x/slideLayout" Target="../notesSlides/notesSlide1.xml"/>';
    for (const r of [notesRel('http://x/n.xml', ' TargetMode="External"'), foreign]) expect(textParity(parts({ ...slideRels(r), ...notes }), withNotes())).toMatchObject({ sourceRuns: 0, outOfScope: 3 });
    expect(textParity(parts({ 'ppt/slides/slide1.xml': '<p:sld/>', ...notes }), withNotes())).toMatchObject({ outOfScope: 3 });
  });
});
