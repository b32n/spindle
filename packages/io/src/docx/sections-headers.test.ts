import JSZip from 'jszip';
import { exportDocx } from './writer';
import { importDocx } from './reader';
import type { DocHeaderFooter, DocSection } from '@b32nio/spindle-shared/export/docx';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const para = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;

/** A docx whose body is `body` (ending in `sectPr`), with header/footer parts `parts` (file name -> paragraphs text) and optional settings. */
async function docx(sectPr: string, parts: Record<string, string> = {}, settings = '', body = para('body')): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${NS}><w:body>${body}${sectPr}</w:body></w:document>`);
  const rels = Object.keys(parts)
    .map((name, i) => `<Relationship Id="rId${10 + i}" Type="${REL}/${name.startsWith('header') ? 'header' : 'footer'}" Target="${name}.xml"/>`)
    .join('');
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`);
  Object.entries(parts).forEach(([name, text]) => {
    const tag = name.startsWith('header') ? 'hdr' : 'ftr';
    zip.file(`word/${name}.xml`, `<?xml version="1.0"?><w:${tag} ${NS}>${para(text)}</w:${tag}>`);
  });
  if (settings) zip.file('word/settings.xml', `<?xml version="1.0"?><w:settings ${NS}>${settings}</w:settings>`);
  return zip.generateAsync({ type: 'uint8array' });
}
/** The rId a `docx()` call gave to the part at `index` (in key order). */
const rid = (index: number) => `rId${10 + index}`;
const text = (hf: DocHeaderFooter | undefined, key: 'paragraphs' | 'firstPageParagraphs' | 'evenPageParagraphs') =>
  hf?.[key]?.map((p) => p.content.map((c) => (c as { text: string }).text).join(''));

describe('header and footer variants (import)', () => {
  it('reads the default, first-page and even-page header by their type, whatever the order of the references', async () => {
    const sectPr = `<w:sectPr><w:headerReference w:type="even" r:id="${rid(2)}"/><w:headerReference w:type="first" r:id="${rid(1)}"/><w:headerReference w:type="default" r:id="${rid(0)}"/><w:titlePg/></w:sectPr>`;
    const { sections } = await importDocx(await docx(sectPr, { header1: 'default', header2: 'first', header3: 'even' }, '<w:evenAndOddHeaders/>'));
    const header = sections[0].header;
    expect(text(header, 'paragraphs')).toEqual(['default']);
    expect(text(header, 'firstPageParagraphs')).toEqual(['first']);
    expect(text(header, 'evenPageParagraphs')).toEqual(['even']);
    expect(header).toMatchObject({ differentFirstPage: true, differentOddEven: true });
  });

  it('treats a reference with no type as the default one', async () => {
    const { sections } = await importDocx(await docx(`<w:sectPr><w:footerReference r:id="${rid(0)}"/></w:sectPr>`, { footer1: 'plain' }));
    expect(text(sections[0].footer, 'paragraphs')).toEqual(['plain']);
    expect(sections[0].footer).not.toHaveProperty('differentFirstPage');
  });

  it('does not use a first-page header as the default one: a first-only header leaves the other pages blank', async () => {
    const { sections } = await importDocx(await docx(`<w:sectPr><w:headerReference w:type="first" r:id="${rid(0)}"/><w:titlePg/></w:sectPr>`, { header1: 'cover' }));
    expect(sections[0].header).toEqual({ paragraphs: [], differentFirstPage: true, firstPageParagraphs: [{ content: [{ type: 'text', text: 'cover' }] }] });
  });

  it('switches the first page on with w:titlePg alone (a blank first page header) and keeps an inert first-page part', async () => {
    const on = await importDocx(await docx(`<w:sectPr><w:headerReference r:id="${rid(0)}"/><w:titlePg/></w:sectPr>`, { header1: 'h' }));
    expect(on.sections[0].header).toMatchObject({ differentFirstPage: true });
    expect(on.sections[0].header?.firstPageParagraphs).toBeUndefined();
    const off = await importDocx(await docx(`<w:sectPr><w:headerReference r:id="${rid(0)}"/><w:headerReference w:type="first" r:id="${rid(1)}"/></w:sectPr>`, { header1: 'h', header2: 'f' }));
    expect(off.sections[0].header).not.toHaveProperty('differentFirstPage');
    expect(text(off.sections[0].header, 'firstPageParagraphs')).toEqual(['f']);
    const explicitOff = await importDocx(await docx(`<w:sectPr><w:headerReference r:id="${rid(0)}"/><w:titlePg w:val="0"/></w:sectPr>`, { header1: 'h' }));
    expect(explicitOff.sections[0].header).not.toHaveProperty('differentFirstPage');
  });

  it('reads the even/odd setting from settings.xml and honours a value of off', async () => {
    const sectPr = `<w:sectPr><w:footerReference r:id="${rid(0)}"/><w:footerReference w:type="even" r:id="${rid(1)}"/></w:sectPr>`;
    const on = await importDocx(await docx(sectPr, { footer1: 'odd', footer2: 'even' }, '<w:evenAndOddHeaders/>'));
    expect(on.sections[0].footer).toMatchObject({ differentOddEven: true });
    expect(text(on.sections[0].footer, 'evenPageParagraphs')).toEqual(['even']);
    const off = await importDocx(await docx(sectPr, { footer1: 'odd', footer2: 'even' }, '<w:evenAndOddHeaders w:val="false"/>'));
    expect(off.sections[0].footer).not.toHaveProperty('differentOddEven');
    expect(text(off.sections[0].footer, 'evenPageParagraphs')).toEqual(['even']);
  });

  it('warns about a variant whose relationship cannot be resolved, and keeps the others', async () => {
    const { sections, warnings } = await importDocx(await docx(`<w:sectPr><w:headerReference r:id="${rid(0)}"/><w:headerReference w:type="first" r:id="rIdMissing"/><w:titlePg/></w:sectPr>`, { header1: 'h' }));
    expect(warnings).toContain('Header/footer relationship "rIdMissing" could not be resolved.');
    expect(text(sections[0].header, 'paragraphs')).toEqual(['h']);
  });

  it('reads each section its own variants', async () => {
    const body = para('one') + `<w:p><w:pPr><w:sectPr><w:headerReference r:id="${rid(0)}"/></w:sectPr></w:pPr></w:p>` + para('two');
    const { sections } = await importDocx(await docx(`<w:sectPr><w:headerReference r:id="${rid(1)}"/><w:headerReference w:type="first" r:id="${rid(2)}"/><w:titlePg/></w:sectPr>`, { header1: 'A', header2: 'B', header3: 'Bfirst' }, '', body));
    expect(text(sections[0].header, 'paragraphs')).toEqual(['A']);
    expect(sections[0].header).not.toHaveProperty('differentFirstPage');
    expect(text(sections[1].header, 'paragraphs')).toEqual(['B']);
    expect(text(sections[1].header, 'firstPageParagraphs')).toEqual(['Bfirst']);
  });
});

describe('section properties (import)', () => {
  const section = async (sectPr: string) => (await importDocx(await docx(`<w:sectPr>${sectPr}</w:sectPr>`))).sections[0];

  it('reads the page number restart and format', async () => {
    expect((await section('<w:pgNumType w:start="10" w:fmt="lowerRoman"/>')).pageNumbering).toEqual({ start: 10, format: 'lowerRoman' });
    expect((await section('<w:pgNumType w:start="0"/>')).pageNumbering).toEqual({ start: 0 });
    expect((await section('<w:pgNumType w:fmt="upperLetter"/>')).pageNumbering).toEqual({ format: 'upperLetter' });
  });

  it('keeps no page numbering for a bare or decimal element, and ignores an unsafe format name', async () => {
    expect(await section('<w:pgNumType/>')).not.toHaveProperty('pageNumbering');
    expect(await section('<w:pgNumType w:fmt="decimal"/>')).not.toHaveProperty('pageNumbering');
    expect(await section('<w:pgNumType w:fmt="a&quot;b"/>')).not.toHaveProperty('pageNumbering');
  });

  it('reads vertical alignment other than top', async () => {
    expect((await section('<w:vAlign w:val="center"/>')).verticalAlign).toBe('center');
    expect((await section('<w:vAlign w:val="both"/>')).verticalAlign).toBe('both');
    expect(await section('<w:vAlign w:val="top"/>')).not.toHaveProperty('verticalAlign');
  });

  it('reads equal columns: count, spacing and separator', async () => {
    expect((await section('<w:cols w:num="3" w:space="360" w:sep="1"/>')).columns).toEqual({ count: 3, spacingPx: 24, separator: true });
    expect((await section('<w:cols w:num="2"/>')).columns).toEqual({ count: 2 });
  });

  it('reads columns of different widths, and takes the count from them when w:num is missing', async () => {
    const cols = (await section('<w:cols w:equalWidth="0" w:space="720"><w:col w:w="2880" w:space="720"/><w:col w:w="5760"/></w:cols>')).columns;
    expect(cols).toEqual({ count: 2, spacingPx: 48, widthsPx: [192, 384] });
  });

  it('keeps no columns for a single column', async () => {
    expect(await section('<w:cols w:num="1" w:space="720"/>')).not.toHaveProperty('columns');
    expect(await section('<w:cols w:space="720"/>')).not.toHaveProperty('columns');
  });
});

describe('header, footer and section properties (export and round trip)', () => {
  const base = (extra: Partial<DocSection> = {}): DocSection => ({
    pageWidthPx: 816,
    pageHeightPx: 1056,
    marginTopPx: 96,
    marginRightPx: 96,
    marginBottomPx: 96,
    marginLeftPx: 96,
    orientation: 'portrait',
    blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'body' }] }],
    ...extra,
  });
  const p = (t: string) => ({ content: [{ type: 'text' as const, text: t }] });
  const roundTrip = async (...sections: DocSection[]) => {
    const { bytes } = await exportDocx({ title: 'T', sections });
    const zip = await JSZip.loadAsync(bytes);
    return { zip, imported: await importDocx(bytes), xml: await zip.file('word/document.xml')!.async('string') };
  };

  it('writes a first-page header and footer and switches the title page on', async () => {
    const { xml, imported } = await roundTrip(
      base({
        header: { paragraphs: [p('H')], differentFirstPage: true, firstPageParagraphs: [p('H first')] },
        footer: { paragraphs: [p('F')], differentFirstPage: true },
      })
    );
    expect(xml).toContain('<w:titlePg/>');
    expect(xml.match(/<w:headerReference w:type="first"/g)).toHaveLength(1);
    // the footer is different on the first page with nothing in it: no first footer part is written, so the first page has none
    expect(xml.match(/<w:footerReference w:type="first"/g)).toBeNull();
    const { header, footer } = imported.sections[0];
    expect(header).toMatchObject({ differentFirstPage: true });
    expect(text(header, 'paragraphs')).toEqual(['H']);
    expect(text(header, 'firstPageParagraphs')).toEqual(['H first']);
    expect(footer).toMatchObject({ differentFirstPage: true });
    expect(footer?.firstPageParagraphs).toBeUndefined();
  });

  it('writes an explicitly empty first-page header as a part of its own', async () => {
    const { xml, imported } = await roundTrip(base({ header: { paragraphs: [p('H')], differentFirstPage: true, firstPageParagraphs: [] } }));
    expect(xml.match(/<w:headerReference w:type="first"/g)).toHaveLength(1);
    expect(imported.sections[0].header?.firstPageParagraphs?.every((x) => x.content.length === 0)).toBe(true);
  });

  it('repeats a footer that is not different on the first page when the header is', async () => {
    const { imported } = await roundTrip(base({ header: { paragraphs: [p('H')], differentFirstPage: true }, footer: { paragraphs: [p('F')] } }));
    expect(text(imported.sections[0].footer, 'firstPageParagraphs')).toEqual(['F']);
  });

  it('writes even-page headers and switches even/odd headers on in the settings', async () => {
    const { zip, imported } = await roundTrip(
      base({ header: { paragraphs: [p('odd')], differentOddEven: true, evenPageParagraphs: [p('even')] }, footer: { paragraphs: [p('foot')] } })
    );
    const settings = await zip.file('word/settings.xml')!.async('string');
    expect(settings).toMatch(/<w:evenAndOddHeaders(?![^>]*w:val="(?:false|0|off)")/);
    const { header, footer } = imported.sections[0];
    expect(header).toMatchObject({ differentOddEven: true });
    expect(text(header, 'evenPageParagraphs')).toEqual(['even']);
    // the footer has no even variant of its own, so it repeats on even pages
    expect(text(footer, 'evenPageParagraphs')).toEqual(['foot']);
  });

  it('does not switch even/odd headers on for a document without them', async () => {
    const { zip } = await roundTrip(base({ header: { paragraphs: [p('H')] } }));
    const settings = zip.file('word/settings.xml') ? await zip.file('word/settings.xml')!.async('string') : '';
    // the library always writes the element; it is off unless a header or footer is different on even pages
    expect(settings).not.toMatch(/<w:evenAndOddHeaders(?![^>]*w:val="(?:false|0|off)")/);
  });

  it('keeps variant parts that are switched off (inert) through a round trip', async () => {
    const hf: DocHeaderFooter = { paragraphs: [p('H')], firstPageParagraphs: [p('inert first')], evenPageParagraphs: [p('inert even')] };
    const { imported } = await roundTrip(base({ header: hf }));
    const header = imported.sections[0].header!;
    expect(text(header, 'firstPageParagraphs')).toEqual(['inert first']);
    expect(header).not.toHaveProperty('differentFirstPage');
    expect(text(header, 'evenPageParagraphs')).toEqual(['inert even']);
    expect(header).not.toHaveProperty('differentOddEven');
  });

  it('writes and reads back page numbering, vertical alignment and columns', async () => {
    const { imported, xml } = await roundTrip(
      base({
        pageNumbering: { start: 5, format: 'lowerRoman' },
        verticalAlign: 'both',
        columns: { count: 2, spacingPx: 24, separator: true },
      })
    );
    expect(xml).toContain('<w:pgNumType w:start="5" w:fmt="lowerRoman"/>');
    expect(imported.sections[0]).toMatchObject({
      pageNumbering: { start: 5, format: 'lowerRoman' },
      verticalAlign: 'both',
      columns: { count: 2, spacingPx: 24, separator: true },
    });
  });

  it('writes columns of different widths', async () => {
    const { imported } = await roundTrip(base({ columns: { count: 2, spacingPx: 48, widthsPx: [192, 384] } }));
    expect(imported.sections[0].columns).toEqual({ count: 2, spacingPx: 48, widthsPx: [192, 384] });
  });

  it('does not write a page number format Word does not know', async () => {
    const { imported } = await roundTrip(base({ pageNumbering: { start: 2, format: 'notAFormat' } }));
    expect(imported.sections[0].pageNumbering).toEqual({ start: 2 });
  });

  it('keeps the properties of each section apart', async () => {
    const { imported } = await roundTrip(base({ pageNumbering: { format: 'lowerRoman' } }), base({ pageNumbering: { start: 1 } }), base());
    expect(imported.sections.map((s) => s.pageNumbering)).toEqual([{ format: 'lowerRoman' }, { start: 1 }, undefined]);
  });
});
