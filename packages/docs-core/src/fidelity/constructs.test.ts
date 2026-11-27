import { unverifiedClaims } from '../../../../scripts/fidelity-core/claims';
import type { OoxmlParts } from '../../../../scripts/fidelity-core/ooxml-package';
import { CONSTRUCTS } from '../../../../scripts/docx-fidelity/constructs';

const parts = (xml: Record<string, string>, extraNames: string[] = []): OoxmlParts => ({
  kind: 'zip',
  raw: new Uint8Array(),
  names: [...Object.keys(xml), ...extraNames],
  xml,
  sizes: {},
  oversized: [],
});
const findConstruct = (id: string) => CONSTRUCTS.find((c) => c.id === id);
const detectConstructs = (p: OoxmlParts) => CONSTRUCTS.filter((c) => c.count(p) > 0).map((c) => c.id);
const body = (inner: string, rest: Record<string, string> = {}) => parts({ 'word/document.xml': `<w:document><w:body>${inner}</w:body></w:document>`, ...rest });
const count = (id: string, p: OoxmlParts) => findConstruct(id)!.count(p);

describe('probes', () => {
  it('has unique ids and a label for each', () => {
    expect(new Set(CONSTRUCTS.map((c) => c.id)).size).toBe(CONSTRUCTS.length);
    expect(CONSTRUCTS.every((c) => c.label.length > 0)).toBe(true);
    expect(findConstruct('nope')).toBeUndefined();
  });

  it('finds nothing in an empty document', () => {
    const empty = parts({});
    expect(detectConstructs(empty).filter((id) => !id.startsWith('edge.'))).toEqual([]);
  });

  it('counts revisions and authors', () => {
    const p = body(
      '<w:ins w:author="A"/><w:del w:author="B"/><w:moveFrom w:author="C"/><w:moveTo w:author="C"/><w:rPrChange w:author="A"/><w:pPrChange w:author="B"/>'
    );
    expect(count('revisions.insertions', p)).toBe(1);
    expect(count('revisions.deletions', p)).toBe(1);
    expect(count('revisions.moves', p)).toBe(1);
    expect(count('revisions.formatting', p)).toBe(2);
    expect(count('revisions.multiAuthor', p)).toBe(3);
    expect(count('revisions.multiAuthor', body('<w:ins w:author="A"/><w:del w:author="B"/>'))).toBe(0);
    expect(count('revisions.moves', body('<w:moveFrom w:author="C"/>'))).toBe(0);
  });

  it('counts comments, notes and fields', () => {
    const p = body(
      '<w:instrText> TOC \\o "1-3" </w:instrText><w:instrText> PAGE </w:instrText><w:instrText> PAGEREF x </w:instrText><w:instrText> REF x </w:instrText><w:instrText> SEQ Fig </w:instrText><w:instrText> DATE </w:instrText><w:instrText> NUMPAGES </w:instrText><w:fldSimple w:instr=" PAGE "/><w:bookmarkStart w:id="1" w:name="x"/>',
      {
        'word/comments.xml': '<w:comments><w:comment w:id="1"/></w:comments>',
        'word/commentsExtended.xml': '<w15:commentsEx><w15:commentEx w15:paraIdParent="1"/></w15:commentsEx>',
        'word/footnotes.xml': '<w:footnotes><w:footnote w:type="separator" w:id="-1"/><w:footnote w:id="1"/></w:footnotes>',
        'word/endnotes.xml': '<w:endnotes><w:endnote w:type="separator" w:id="-1"/><w:endnote w:id="1"/></w:endnotes>',
      }
    );
    for (const id of ['comments', 'comments.threaded', 'notes.footnotes', 'notes.endnotes', 'fields.toc', 'fields.page', 'fields.pageref', 'fields.ref', 'fields.seq', 'fields.date', 'fields.other', 'fields.bookmarks']) {
      expect([id, count(id, p)]).toEqual([id, expect.any(Number)]);
      expect(count(id, p)).toBeGreaterThan(0);
    }
    expect(count('notes.footnotes', p)).toBe(1);
    expect(count('fields.page', p)).toBe(2);
    expect(count('fields.ref', body('<w:instrText> PAGEREF x </w:instrText>'))).toBe(0);
  });

  it('counts hyperlinks, headings and lists', () => {
    const p = body('<w:hyperlink w:anchor="x"/><w:pStyle w:val="Heading1"/><w:pStyle w:val="heading 2"/><w:numPr><w:ilvl w:val="2"/><w:numId w:val="3"/></w:numPr><w:numPr><w:numId w:val="4"/></w:numPr>', {
      'word/_rels/document.xml.rels': '<Relationships><Relationship Id="r" Type="http://x/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>',
      'word/numbering.xml': '<w:numbering><w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num><w:num w:numId="4"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
    });
    expect(count('hyperlinks.external', p)).toBe(1);
    expect(count('hyperlinks.internal', p)).toBe(1);
    expect(count('headings', p)).toBe(2);
    expect(count('lists.numbered', p)).toBe(2);
    expect(count('lists.multilevel', p)).toBe(1);
    expect(count('lists.restart', p)).toBe(1);
  });

  it('counts style relationships', () => {
    const p = parts({ 'word/styles.xml': '<w:styles><w:basedOn w:val="N"/><w:link w:val="x"/></w:styles>' });
    expect(count('styles.inheritance', p)).toBe(1);
    expect(count('styles.linked', p)).toBe(1);
  });

  it('counts table features', () => {
    const nested = body('<w:tbl><w:tr><w:tc><w:tbl><w:tr/></w:tbl></w:tc></w:tr></w:tbl>');
    expect(count('tables.nested', nested)).toBe(1);
    expect(count('tables.nested', body('<w:tbl><w:tc></w:tc></w:tbl><w:tbl/>'))).toBe(0);
    const p = body('<w:gridSpan w:val="2"/><w:vMerge/><w:tblBorders/><w:tcPr><w:tcW/><w:shd w:fill="x"/></w:tcPr><w:tblHeader/>');
    for (const id of ['tables.mergeH', 'tables.mergeV', 'tables.borders', 'tables.shading', 'tables.headerRow']) expect(count(id, p)).toBeGreaterThan(0);
    expect(count('tables.autofit', body('<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'))).toBe(1);
    expect(count('tables.autofit', body('<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>'))).toBe(0);
    expect(count('tables.autofit', body('<w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblLayout w:type="autofit"/></w:tblPr>'))).toBe(1);
  });

  it('counts section features', () => {
    const p = body(
      '<w:pgSz w:w="1" w:h="2" w:orient="landscape"/><w:pgSz w:w="1" w:h="2"/><w:cols w:num="2"/><w:cols w:num="1"/><w:titlePg/><w:headerReference w:type="even" r:id="x"/><w:pgNumType w:start="1"/>',
      { 'word/settings.xml': '<w:settings><w:evenAndOddHeaders/></w:settings>' }
    );
    expect(count('sections.landscape', p)).toBe(1);
    expect(count('sections.portrait', p)).toBe(1);
    expect(count('sections.columns', p)).toBe(1);
    expect(count('sections.firstPage', p)).toBe(1);
    expect(count('sections.oddEven', p)).toBe(1);
    expect(count('sections.pageNumberRestart', p)).toBe(1);
    expect(count('sections.oddEven', body('<w:headerReference w:type="even"/>'))).toBe(0);
  });

  it('counts drawing features', () => {
    const p = body(
      '<wp:inline><pic:pic/></wp:inline><wp:anchor><wp:wrapSquare/><pic:pic/></wp:anchor><a:srcRect l="1000"/><wps:wsp/><w:txbxContent/><m:oMath/><w:sdt/><w:ffData/><w:object/><w:vanish/><w:framePr w:dropCap="drop"/><w:noHyphen/>'
    );
    for (const id of ['images.inline', 'images.floating', 'images.wrapped', 'images.cropped', 'shapes', 'textboxes', 'equations', 'sdt', 'forms.legacy', 'objects.embedded', 'text.hidden', 'typography.dropCap', 'typography.hyphenation']) {
      expect([id, count(id, p) > 0]).toEqual([id, true]);
    }
    expect(count('images.cropped', body('<a:srcRect/>'))).toBe(0);
  });

  it('counts package-level features', () => {
    const p = parts(
      {
        'word/document.xml': '<w:document><w:rFonts w:asciiTheme="minorHAnsi"/></w:document>',
        'word/theme/theme1.xml': '<a:theme/>',
        'word/styles.xml': '<w:styles><w:rFonts w:asciiTheme="minorHAnsi"/></w:styles>',
        'word/header1.xml': '<w:hdr><w:t>Mark</w:t><v:shape id="PowerPlusWaterMarkObject1"/></w:hdr>',
      },
      ['word/fonts/font1.odttf', 'word/charts/chart1.xml', 'customXml/item1.xml', 'word/diagrams/data1.xml', 'word/diagrams/drawing1.xml']
    );
    for (const id of ['fonts.embedded', 'themes', 'charts', 'customXml', 'smartart', 'watermark', 'headers.content']) expect([id, count(id, p) > 0]).toEqual([id, true]);
    expect(count('smartart', parts({}, ['word/diagrams/data1.xml']))).toBe(0);
    expect(count('themes', parts({ 'word/styles.xml': '<w:styles w:asciiTheme="x"/>' }))).toBe(0);
    expect(count('themes', parts({ 'word/document.xml': '<w:document><w:color w:themeColor="accent1"/></w:document>' }, ['word/theme/theme1.xml']))).toBe(1);
  });

  it('counts text features', () => {
    const p = body('<w:bidi/><w:t>漢字 😀 → </w:t><w:sym w:font="Symbol" w:char="F0B7"/><w:tab w:val="right" w:leader="dot"/><w:ptab w:leader="dot"/>', {
      'word/settings.xml': '<w:settings><w:autoHyphenation/></w:settings>',
    });
    for (const id of ['i18n.rtl', 'i18n.cjk', 'text.emoji', 'text.symbols', 'typography.hyphenation', 'typography.tabLeaders']) expect([id, count(id, p) > 0]).toEqual([id, true]);
  });

  it('detects long documents by word count', () => {
    expect(count('doc.long', body(`<w:t>${'word '.repeat(15000)}</w:t>`))).toBeGreaterThanOrEqual(15000);
    expect(count('doc.long', body('<w:t>short</w:t>'))).toBe(0);
  });

  it('detects structural defects', () => {
    const p = parts({
      'word/document.xml': '<w:document><w:body></w:document>',
      'word/_rels/document.xml.rels':
        '<Relationships><Relationship Id="a" Type="t" Target="styles.xml"/><Relationship Id="b" Type="t" Target="gone.xml"/><Relationship Id="c" Type="t" Target="/word/abs.xml"/><Relationship Id="d" Type="t" Target="https://example.com" TargetMode="External"/><Relationship Id="e" Type="t"/></Relationships>',
      'word/styles.xml': '<w:styles/>',
    });
    expect(count('edge.brokenRels', p)).toBe(2); // gone.xml and /word/abs.xml
    expect(count('edge.invalidXml', p)).toBe(1);
    expect(count('edge.missingOptionalParts', p)).toBe(2); // settings and numbering
    expect(count('edge.brokenRels', parts({ 'word/_rels/document.xml.rels': '<Relationships><Relationship Target="media/a.png"/></Relationships>' }, ['word/media/a.png']))).toBe(0);
  });

  it('detects a zip bomb by the import limit and entity declarations', () => {
    const MB = 1024 * 1024;
    const sized = (sizes: OoxmlParts['sizes']): OoxmlParts => ({ ...parts({}), sizes });
    expect(count('edge.zipBomb', sized({ 'a.bin': { compressed: 24 * 1024, uncompressed: 24 * MB } }))).toBe(1);
    expect(count('edge.zipBomb', sized({ 'a.bin': { compressed: 0, uncompressed: 24 * MB } }))).toBe(1);
    expect(count('edge.zipBomb', sized({ 'small.bin': { compressed: 1, uncompressed: 1 * MB } }))).toBe(0); // under the 16 MiB floor
    expect(count('edge.zipBomb', sized({ 'big.bin': { compressed: 20 * MB, uncompressed: 24 * MB } }))).toBe(0); // ordinary ratio
    expect(count('edge.entityExpansion', parts({ 'word/document.xml': '<!DOCTYPE d [<!ENTITY a "x"><!ENTITY b "&a;&a;">]><d/>', 'word/styles.xml': '<s/>' }))).toBe(2);
    expect(count('edge.entityExpansion', parts({ 'word/document.xml': '<d/>' }))).toBe(0);
  });

  it('reports unverified claims', () => {
    const p = body('<w:sdt/>');
    expect(unverifiedClaims(['sdt', 'equations', 'unknown'], CONSTRUCTS, p)).toEqual(['equations', 'unknown']);
  });
});
