import JSZip from 'jszip';
import { exportDocx } from './writer';
import { importDocx } from './reader';
import { detectDocxLoss } from './loss';
import { parseSdt, unsupportedControls, isKeptBlockControl } from './content-controls';
import { parseXml, childrenOf, find } from './xml-tree';
import type { DocBlock, DocContentControl, DocInline, DocSection } from '@b32nio/spindle-shared/export/docx';

const BASE: Omit<DocSection, 'blocks'> = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginTopPx: 96,
  marginRightPx: 96,
  marginBottomPx: 96,
  marginLeftPx: 96,
  orientation: 'portrait',
};
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (inner: string) => `<w:p>${inner}</w:p>`;
const sdt = (pr: string, content: string) => `<w:sdt><w:sdtPr>${pr}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
type Paragraph = Extract<DocBlock, { type: 'paragraph' }>;
type ControlBlock = Extract<DocBlock, { type: 'contentControl' }>;

async function bodyDocx(bodyXml: string): Promise<Uint8Array> {
  const exported = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'x' }] }] }] });
  const zip = await JSZip.loadAsync(exported.bytes);
  let xml = await zip.file('word/document.xml')!.async('string');
  // Declare the extension namespaces the fixtures use, whatever the writer declared.
  xml = xml.replace(/<w:p>[\s\S]*?<\/w:p>/, bodyXml);
  zip.file('word/document.xml', xml);
  return zip.generateAsync({ type: 'uint8array' });
}
const importBody = async (bodyXml: string) => importDocx(await bodyDocx(bodyXml));
const blocksOf = async (bodyXml: string) => (await importBody(bodyXml)).sections[0].blocks;
const inlineControl = async (pr: string, content = r('v')): Promise<DocContentControl> => {
  const [block] = await blocksOf(p(sdt(pr, content)).replace('<w:p>', `<w:p>`));
  const first = (block as Paragraph).content[0];
  if (first.type !== 'contentControl') throw new Error('not a control');
  return first.control;
};

async function exported(blocks: DocBlock[]) {
  const out = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks }] });
  const zip = await JSZip.loadAsync(out.bytes);
  return { xml: await zip.file('word/document.xml')!.async('string'), bytes: out.bytes, warnings: out.warnings };
}

const text = (t: string): DocInline => ({ type: 'text', text: t });

describe('reading content controls', () => {
  it('reads each control type with its properties', async () => {
    expect(await inlineControl('<w:alias w:val="Name"/><w:tag w:val="name"/><w:text/>')).toEqual({ type: 'text', alias: 'Name', tag: 'name' });
    expect(await inlineControl('<w:text w:multiLine="1"/>')).toEqual({ type: 'text', multiline: true });
    expect(await inlineControl('<w:richText/>')).toEqual({ type: 'richText' });
    expect(await inlineControl('')).toEqual({ type: 'richText' });
    expect(await inlineControl('<w:dropDownList><w:listItem w:displayText="North" w:value="N"/><w:listItem w:value="E"/><w:listItem w:displayText="West"/></w:dropDownList>')).toEqual({
      type: 'dropDown',
      items: [
        { text: 'North', value: 'N' },
        { text: 'E', value: 'E' },
        { text: 'West', value: 'West' },
      ],
    });
    expect((await inlineControl('<w:comboBox><w:listItem w:displayText="a" w:value="b"/></w:comboBox>')).type).toBe('comboBox');
    expect(await inlineControl('<w:date w:fullDate="2026-03-02T00:00:00Z"><w:dateFormat w:val="d MMMM yyyy"/><w:lid w:val="en-US"/></w:date>')).toEqual({
      type: 'date',
      dateFormat: 'd MMMM yyyy',
      dateValue: '2026-03-02T00:00:00Z',
    });
    expect(await inlineControl('<w:date/>')).toEqual({ type: 'date' });
    expect((await inlineControl('<w:picture/>')).type).toBe('picture');
    expect(await inlineControl('<w:docPartObj><w:docPartGallery w:val="Cover Pages"/><w:docPartUnique/></w:docPartObj>')).toEqual({ type: 'docPart', gallery: 'Cover Pages' });
    expect(await inlineControl('<w:docPartList><w:docPartGallery w:val="Equations"/></w:docPartList>')).toEqual({ type: 'docPart', gallery: 'Equations' });
    expect(await inlineControl('<w:docPartObj/>')).toEqual({ type: 'docPart' });
  });

  it('reads a check box state, as on or off, and a missing state as off', async () => {
    expect(await inlineControl('<w14:checkbox><w14:checked w14:val="1"/></w14:checkbox>')).toEqual({ type: 'checkbox', checked: true });
    expect(await inlineControl('<w14:checkbox><w14:checked w14:val="true"/></w14:checkbox>')).toEqual({ type: 'checkbox', checked: true });
    expect(await inlineControl('<w14:checkbox><w14:checked w14:val="0"/></w14:checkbox>')).toEqual({ type: 'checkbox', checked: false });
    expect(await inlineControl('<w14:checkbox><w14:checked/></w14:checkbox>')).toEqual({ type: 'checkbox', checked: true });
    expect(await inlineControl('<w14:checkbox/>')).toEqual({ type: 'checkbox', checked: false });
  });

  it('reads the lock, the placeholder and whether the placeholder is showing', async () => {
    expect(await inlineControl('<w:lock w:val="sdtContentLocked"/>')).toEqual({ type: 'richText', lock: 'sdtContentLocked' });
    expect(await inlineControl('<w:lock w:val="contentLocked"/>')).toEqual({ type: 'richText', lock: 'contentLocked' });
    expect(await inlineControl('<w:lock w:val="sdtLocked"/>')).toEqual({ type: 'richText', lock: 'sdtLocked' });
    expect(await inlineControl('<w:lock w:val="unlocked"/>')).toEqual({ type: 'richText' });
    expect(await inlineControl('<w:placeholder><w:docPart w:val="DefaultPlaceholder_1"/></w:placeholder><w:showingPlcHdr/>')).toEqual({
      type: 'richText',
      placeholder: 'DefaultPlaceholder_1',
      showingPlaceholder: true,
    });
    expect(await inlineControl('<w:showingPlcHdr w:val="0"/>')).toEqual({ type: 'richText' });
  });

  it('keeps the formatted content of an inline control, with a hyperlink and a bookmark inside', async () => {
    const body = `<w:p ${W}>${r('a ')}${sdt('<w:alias w:val="C"/>', `<w:bookmarkStart w:id="1" w:name="mark"/><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>`)}${r(' z')}</w:p>`;
    const [block] = await blocksOf(body);
    expect((block as Paragraph).content).toMatchObject([
      { type: 'text', text: 'a ' },
      { type: 'contentControl', control: { alias: 'C' }, content: [{ type: 'bookmark', name: 'mark' }, { type: 'text', text: 'bold', bold: true }] },
      { type: 'text', text: ' z' },
    ]);
  });

  it('nests inline controls and puts a control inside a hyperlink and a hyperlink inside a control', async () => {
    const body = p(
      sdt('<w:alias w:val="outer"/>', r('x') + sdt('<w:alias w:val="inner"/>', r('y'))) +
        `<w:hyperlink w:anchor="mark">${sdt('<w:alias w:val="inlink"/>', r('l'))}</w:hyperlink>` +
        sdt('', `<w:hyperlink w:anchor="mark">${r('m')}</w:hyperlink>`)
    );
    const [block] = await blocksOf(body.replace('<w:p>', `<w:p ${W}>`));
    const content = (block as Paragraph).content;
    expect(content[0]).toMatchObject({ type: 'contentControl', control: { alias: 'outer' }, content: [{ type: 'text', text: 'x' }, { type: 'contentControl', control: { alias: 'inner' }, content: [{ type: 'text', text: 'y' }] }] });
    // Inside a hyperlink the control's runs are links; a hyperlink inside a control keeps its target.
    expect(content[1]).toMatchObject({ type: 'contentControl', control: { alias: 'inlink' } });
    expect(content[2]).toMatchObject({ type: 'contentControl', content: [{ type: 'link', text: 'm', href: '#mark' }] });
  });

  it('an empty control and a control with no content element keep their place', async () => {
    const [block] = await blocksOf(`<w:p ${W}>${r('a')}<w:sdt><w:sdtPr><w:alias w:val="e"/></w:sdtPr></w:sdt>${r('b')}</w:p>`);
    expect((block as Paragraph).content).toMatchObject([{ type: 'text' }, { type: 'contentControl', control: { alias: 'e' }, content: [] }, { type: 'text' }]);
  });

  it('reads a block-level control with several blocks, a table and nested controls; its properties are kept', async () => {
    const cell = `<w:tbl ${W}><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p>${r('cell')}</w:p></w:tc></w:tr></w:tbl>`;
    const body =
      `<w:sdt ${W}><w:sdtPr><w:alias w:val="Outer"/><w:tag w:val="o"/></w:sdtPr><w:sdtContent>` +
      `<w:p>${r('one')}</w:p>${cell}` +
      `<w:sdt><w:sdtPr><w:alias w:val="Inner"/></w:sdtPr><w:sdtContent><w:p>${r('two')}</w:p></w:sdtContent></w:sdt>` +
      `</w:sdtContent></w:sdt>`;
    const [block] = await blocksOf(body);
    expect(block.type).toBe('contentControl');
    const outer = block as ControlBlock;
    expect(outer.control).toEqual({ type: 'richText', alias: 'Outer', tag: 'o' });
    expect(outer.blocks.map((b) => b.type)).toEqual(['paragraph', 'table', 'contentControl']);
    expect((outer.blocks[2] as ControlBlock).control.alias).toBe('Inner');
  });

  it('keeps numbered lists numbering across control boundaries and bookmarks before a control', async () => {
    const item = (t: string, n: number) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${n}"/></w:numPr></w:pPr>${r(t)}</w:p>`;
    const body = `<w:bookmarkStart ${W} w:id="1" w:name="before"/>` + sdt('', p(r('lead'))) + p(r('tail'));
    const [block, tail] = await blocksOf(body);
    expect(((block as ControlBlock).blocks[0] as Paragraph).content).toEqual([{ type: 'bookmark', name: 'before' }, { type: 'text', text: 'lead' }]);
    expect((tail as Paragraph).content).toEqual([{ type: 'text', text: 'tail' }]);
    // A list item inside a control stays a list item.
    const [listControl] = await blocksOf(sdt('', item('i', 1)).replace('<w:sdt>', `<w:sdt ${W}>`));
    expect(((listControl as ControlBlock).blocks[0]).type).toBe('listItem');
  });

  it('reads a control around a section break as its content only, so the section split stays correct', async () => {
    const body = `<w:sdt ${W}><w:sdtPr/><w:sdtContent><w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>${r('end of section one')}</w:p></w:sdtContent></w:sdt>${p(r('two'))}`;
    const result = await importBody(body);
    expect(result.sections.length).toBeGreaterThanOrEqual(2);
    expect(result.sections[0].blocks).toEqual([{ type: 'paragraph', content: [{ type: 'text', text: 'end of section one' }] }]);
  });

  it('reads controls that are not supported (repeating sections, data binding, group...) as their content, without a wrapper', async () => {
    const kinds = ['<w15:repeatingSection/>', '<w15:repeatingSectionItem/>', '<w:dataBinding w:xpath="/a" w:storeItemID="{1}"/>', '<w:equation/>', '<w:citation/>', '<w:bibliography/>', '<w:group/>'];
    for (const pr of kinds) {
      const blocks = await blocksOf(sdt(pr, p(r('kept'))).replace('<w:sdt>', `<w:sdt ${W}>`));
      expect(blocks).toEqual([{ type: 'paragraph', content: [{ type: 'text', text: 'kept' }] }]);
      const inline = await blocksOf(`<w:p ${W}>${sdt(pr, r('kept'))}</w:p>`);
      expect((inline[0] as Paragraph).content).toEqual([{ type: 'text', text: 'kept' }]);
    }
  });

  it('a control inside the result of a field is read as its content', async () => {
    const field =
      `<w:p ${W}><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      sdt('', r('cached')) +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
    const [block] = await blocksOf(field);
    expect((block as Paragraph).content).toMatchObject([{ type: 'field', kind: 'REF', result: 'cached' }]);
  });

  it('keeps a control wrapping the whole content of a table cell, and reads one around a row or a cell as content', async () => {
    const table = (rows: string) => `<w:tbl ${W}><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl>`;
    const wrapped = await blocksOf(table(`<w:tr><w:tc>${sdt('<w:alias w:val="C"/>', p(r('first')) + p(r('second')))}</w:tc></w:tr>`));
    const cell = (wrapped[0] as Extract<DocBlock, { type: 'table' }>).rows[0].cells[0];
    expect(cell.blocks).toMatchObject([{ type: 'contentControl', control: { alias: 'C' }, blocks: [{ type: 'paragraph' }, { type: 'paragraph' }] }]);
    expect(cell.content).toEqual([{ type: 'text', text: 'first' }]);
    const rowLevel = await blocksOf(table(sdt('', `<w:tr><w:tc>${p(r('row'))}</w:tc></w:tr>`)));
    expect((rowLevel[0] as Extract<DocBlock, { type: 'table' }>).rows[0].cells[0].content).toEqual([{ type: 'text', text: 'row' }]);
  });

  it('keeps a nested table and a control around it inside a control', async () => {
    const inner = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc>${sdt('', p(r('deep')))}</w:tc></w:tr></w:tbl>`;
    const [block] = await blocksOf(sdt('<w:alias w:val="o"/>', inner).replace('<w:sdt>', `<w:sdt ${W}>`));
    const table = (block as ControlBlock).blocks[0] as Extract<DocBlock, { type: 'table' }>;
    expect(table.rows[0].cells[0].blocks).toMatchObject([{ type: 'contentControl', blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'deep' }] }] }]);
  });

  it('control inside a header or footer is read as its content', async () => {
    const out = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [{ type: 'paragraph', content: [] }], header: { paragraphs: [{ content: [text('h')] }] } }] });
    const zip = await JSZip.loadAsync(out.bytes);
    const name = Object.keys(zip.files).find((n) => /^word\/header\d*\.xml$/.test(n))!;
    const xml = (await zip.file(name)!.async('string')).replace(/<w:p>[\s\S]*?<\/w:p>/, `<w:p>${sdt('<w:alias w:val="H"/>', r('Head'))}</w:p>`);
    zip.file(name, xml);
    const result = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(result.sections[0].header?.paragraphs).toEqual([{ content: [{ type: 'text', text: 'Head' }] }]);
  });
});

describe('parseSdt and the helpers behind the loss report', () => {
  const node = (xml: string) => find(parseXml(`<w:body ${W}>${xml}</w:body>`)[0] ? childrenOf(parseXml(`<w:body ${W}>${xml}</w:body>`)[0]) : [], 'w:sdt')!;

  it('says why a control cannot be kept', () => {
    expect(parseSdt(node(sdt('<w15:repeatingSection/>', '')))).toEqual({ unsupported: 'repeating section' });
    expect(parseSdt(node(sdt('<w:dataBinding/>', '')))).toEqual({ unsupported: 'data-bound control' });
    expect(parseSdt(node(sdt('<w:alias w:val="a"/>', '')))).toEqual({ control: { type: 'richText', alias: 'a' } });
    expect(parseSdt(node('<w:sdt/>'))).toEqual({ control: { type: 'richText' } });
  });

  it('keeps a block control unless a section break is inside it', () => {
    expect(isKeptBlockControl(node(sdt('', p(r('a')))))).toBe(true);
    expect(isKeptBlockControl(node(sdt('<w:group/>', p(r('a')))))).toBe(false);
    expect(isKeptBlockControl(node(sdt('', '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>')))).toBe(false);
  });

  it('counts unsupported controls by reason, wherever they sit', () => {
    const root = parseXml(
      `<w:document ${W}><w:body>` +
        sdt('<w15:repeatingSection/>', sdt('<w15:repeatingSectionItem/>', p(r('x'))) + sdt('<w15:repeatingSectionItem/>', p(r('y')))) +
        `<w:tbl>${sdt('', '<w:tr/>')}<w:tr>${sdt('', '<w:tc/>')}</w:tr></w:tbl>` +
        sdt('', '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>') +
        p(sdt('<w:dataBinding/>', r('z'))) +
        p(sdt('<w:alias w:val="fine"/>', r('ok'))) +
        `</w:body></w:document>`
    );
    expect(Object.fromEntries(unsupportedControls(root))).toEqual({
      'repeating section': 1,
      'repeating section item': 2,
      'row-level control': 1,
      'cell-level control': 1,
      'control around a section break': 1,
      'data-bound control': 1,
    });
  });
});

describe('writing content controls', () => {
  it('writes the properties in order and a distinct w:id per control', async () => {
    const { xml } = await exported([
      {
        type: 'paragraph',
        content: [
          { type: 'contentControl', control: { type: 'text', alias: 'A', tag: 't', lock: 'sdtLocked', placeholder: 'ph', showingPlaceholder: true, multiline: true }, content: [text('x')] },
          { type: 'contentControl', control: { type: 'richText' }, content: [] },
        ],
      },
    ]);
    expect(xml).toContain('<w:sdt><w:sdtPr><w:alias w:val="A"/><w:tag w:val="t"/><w:id w:val="1000001"/><w:lock w:val="sdtLocked"/><w:placeholder><w:docPart w:val="ph"/></w:placeholder><w:showingPlcHdr/><w:text w:multiLine="1"/></w:sdtPr><w:sdtContent>');
    const ids = [...xml.matchAll(/<w:id w:val="(\d+)"\/>/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids.length).toBe(2);
  });

  it('writes each type with its own element', async () => {
    const control = (c: DocContentControl): DocBlock => ({ type: 'paragraph', content: [{ type: 'contentControl', control: c, content: [text('v')] }] });
    const { xml } = await exported([
      control({ type: 'dropDown', items: [{ text: 'North', value: 'N' }] }),
      control({ type: 'comboBox', items: [] }),
      control({ type: 'date', dateFormat: 'dd/MM/yyyy', dateValue: '2026-03-02T00:00:00Z' }),
      control({ type: 'date' }),
      control({ type: 'checkbox', checked: true }),
      control({ type: 'checkbox' }),
      control({ type: 'picture' }),
      control({ type: 'docPart', gallery: 'Cover Pages' }),
      control({ type: 'docPart' }),
    ]);
    expect(xml).toContain('<w:dropDownList><w:listItem w:displayText="North" w:value="N"/></w:dropDownList>');
    expect(xml).toContain('<w:comboBox/>');
    expect(xml).toContain('<w:date w:fullDate="2026-03-02T00:00:00Z"><w:dateFormat w:val="dd/MM/yyyy"/>');
    expect(xml).toContain('<w:date><w:lid');
    expect(xml).toContain('<w14:checkbox><w14:checked w14:val="1"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/><w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>');
    expect(xml).toContain('<w14:checked w14:val="0"/>');
    expect(xml).toContain('<w:picture/>');
    expect(xml).toContain('<w:docPartObj><w:docPartGallery w:val="Cover Pages"/><w:docPartUnique/></w:docPartObj>');
    expect(xml).toContain('<w:docPartObj><w:docPartUnique/></w:docPartObj>');
  });

  it('writes a block control around blocks; an empty one gets a paragraph; a table at the end of a cell gets one after it', async () => {
    const table: DocBlock = { type: 'table', rows: [{ cells: [{ content: [text('t')] }] }] };
    const { xml } = await exported([
      { type: 'contentControl', control: { type: 'richText', alias: 'B' }, blocks: [] },
      { type: 'contentControl', control: { type: 'richText' }, blocks: [{ type: 'paragraph', content: [text('in')] }, table] },
      {
        type: 'table',
        rows: [{ cells: [{ content: [], blocks: [{ type: 'contentControl', control: { type: 'richText' }, blocks: [table] }] }] }],
      },
    ]);
    expect(xml).toMatch(/<w:sdt><w:sdtPr><w:alias w:val="B"\/>.*?<\/w:sdtPr><w:sdtContent><w:p\b/);
    // The cell holding a control that ends with a table ends with a paragraph.
    expect(xml).toMatch(/<\/w:sdt><w:p\/><\/w:tc>/);
  });

  it('round-trips controls of every type, nested, with comments anchored inside them', async () => {
    const blocks: DocBlock[] = [
      {
        type: 'paragraph',
        content: [
          { type: 'contentControl', control: { type: 'text', alias: 'A' }, content: [{ type: 'text', text: 'one', commentThreadId: 't1' }] },
          { type: 'contentControl', control: { type: 'checkbox', checked: false }, content: [text('☐')] },
        ],
      },
      { type: 'contentControl', control: { type: 'richText', alias: 'O' }, blocks: [{ type: 'contentControl', control: { type: 'richText', alias: 'I' }, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'two', commentThreadId: 't1' }] }] }] },
    ];
    const out = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks }], comments: [{ id: 't1', quote: 'one', comments: [{ id: '1', author: 'A', text: 'hi' }] }] });
    const result = await importDocx(out.bytes);
    const read = result.sections[0].blocks;
    expect(read[0]).toMatchObject({ type: 'paragraph', content: [{ type: 'contentControl', control: { alias: 'A' } }, { type: 'contentControl', control: { type: 'checkbox', checked: false } }] });
    expect((read[1] as ControlBlock).blocks[0]).toMatchObject({ type: 'contentControl', control: { alias: 'I' } });
    // The comment's range runs from the first tagged run (inside the first control) to the last (inside the nested ones).
    expect(result.comments?.[0].id).toBeDefined();
    const doc = await JSZip.loadAsync(out.bytes).then((z) => z.file('word/document.xml')!.async('string'));
    expect(doc.match(/commentRangeStart/g)).toHaveLength(1);
    expect(doc.match(/commentRangeEnd/g)).toHaveLength(1);
  });

  it('keeps the named styles used only inside controls', async () => {
    const blocks: DocBlock[] = [
      { type: 'contentControl', control: { type: 'richText' }, blocks: [{ type: 'paragraph', namedStyleId: 'Callout', content: [text('c')] }] },
      { type: 'paragraph', content: [{ type: 'contentControl', control: { type: 'text' }, content: [{ type: 'text', text: 'a', namedStyleId: 'Accent' }] }] },
    ];
    const out = await exportDocx({
      title: 'T',
      sections: [{ ...BASE, blocks }],
      namedStyles: { Callout: { name: 'Callout', kind: 'paragraph' }, Accent: { name: 'Accent', kind: 'character' }, Unused: { name: 'Unused', kind: 'paragraph' } },
    });
    const result = await importDocx(out.bytes);
    expect(Object.keys(result.namedStyles ?? {}).sort()).toEqual(['Accent', 'Callout']);
    expect(((result.sections[0].blocks[0] as ControlBlock).blocks[0] as Paragraph).namedStyleId).toBe('Callout');
  });

  it('a numbered list inside a control restarts its own numbering', async () => {
    const item = (t: string, start?: number): DocBlock => ({ type: 'listItem', listType: 'numbered', level: 0, content: [text(t)], ...(start && { start }) });
    const out = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [item('a'), item('b'), { type: 'contentControl', control: { type: 'richText' }, blocks: [item('c', 3), item('d')] }, item('e', 4)] }] });
    const read = (await importDocx(out.bytes)).sections[0].blocks;
    const inner = (read[2] as ControlBlock).blocks as Array<Extract<DocBlock, { type: 'listItem' }>>;
    expect(inner.map((b) => b.start)).toEqual([3, undefined]);
    expect((read[3] as Extract<DocBlock, { type: 'listItem' }>).start).toBe(4);
  });
});

describe('detecting what is lost about content controls', () => {
  it('reports only the controls it cannot keep, by reason, in both detail levels', async () => {
    const bytes = await bodyDocx(
      sdt('<w:alias w:val="fine"/>', p(r('ok'))).replace('<w:sdt>', `<w:sdt ${W}>`) +
        sdt('<w15:repeatingSection/>', sdt('<w15:repeatingSectionItem/>', p(r('x')))) +
        `<w:p>${sdt('<w:dataBinding/>', r('bound'))}</w:p>`
    );
    const full = await detectDocxLoss(bytes, { detail: 'full' });
    const item = full.items.find((i) => i.id === 'docx.contentControls');
    expect(item).toMatchObject({ count: 3, where: ['data-bound control ×1', 'repeating section ×1', 'repeating section item ×1'] });
    const summary = await detectDocxLoss(bytes);
    const short = summary.items.find((i) => i.id === 'docx.contentControls');
    expect(short?.count).toBe(3);
    expect(short).not.toHaveProperty('where');
  });

  it('reports nothing when every control is kept', async () => {
    const bytes = await bodyDocx(sdt('', p(r('a'))).replace('<w:sdt>', `<w:sdt ${W}>`) + `<w:p>${sdt('<w:text/>', r('b'))}</w:p>`);
    expect((await detectDocxLoss(bytes, { detail: 'full' })).items.find((i) => i.id === 'docx.contentControls')).toBeUndefined();
  });
});
