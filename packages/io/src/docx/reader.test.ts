import JSZip from 'jszip';
import { exportDocx } from './writer';
import { importDocx } from './reader';
import type { DocBlock, DocSection, DocumentExportInput } from '@b32nio/spindle-shared/export/docx';

const BASE_SECTION: Omit<DocSection, 'blocks'> = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginTopPx: 96,
  marginRightPx: 96,
  marginBottomPx: 96,
  marginLeftPx: 96,
  orientation: 'portrait',
};

const ONE_PX_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

async function roundTrip(input: DocumentExportInput) {
  const exported = await exportDocx(input);
  return importDocx(exported.bytes);
}

function section(blocks: DocSection['blocks'], extra: Partial<DocSection> = {}): DocumentExportInput {
  return { title: 'Test', sections: [{ ...BASE_SECTION, ...extra, blocks }] };
}

describe('importDocx', () => {
  it('rejects a non-zip byte stream', async () => {
    await expect(importDocx(new Uint8Array([0, 1, 2, 3]))).rejects.toThrow();
  });

  it('surfaces a warning instead of throwing for a zip with no word/document.xml', async () => {
    const zip = new JSZip();
    zip.file('hello.txt', 'not a docx');
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    const result = await importDocx(bytes);
    expect(result.sections).toEqual([]);
    expect(result.warnings.length).toBeGreaterThan(0);
  });

  it('round-trips the title', async () => {
    const result = await roundTrip(section([{ type: 'paragraph', content: [{ type: 'text', text: 'Hi' }] }]));
    expect(result.title).toBe('Test');
  });

  it('round-trips plain text and run formatting', async () => {
    const result = await roundTrip(
      section([
        {
          type: 'paragraph',
          content: [
            { type: 'text', text: 'Bold', bold: true },
            { type: 'text', text: 'Italic', italic: true },
            { type: 'text', text: 'Underline', underline: true },
            { type: 'text', text: 'Strike', strikethrough: true },
            { type: 'text', text: 'Super', superscript: true },
            { type: 'text', text: 'Sub', subscript: true },
            { type: 'text', text: 'Colored', color: '#ff0000' },
            { type: 'text', text: 'Highlighted', backgroundColor: '#00ff00' },
            { type: 'text', text: 'Sized', fontSize: 18 },
            { type: 'text', text: 'Fonted', fontFamily: 'Georgia' },
          ],
        },
      ])
    );
    expect(result.sections[0].blocks).toEqual([
      {
        type: 'paragraph',
        content: [
          { type: 'text', text: 'Bold', bold: true },
          { type: 'text', text: 'Italic', italic: true },
          { type: 'text', text: 'Underline', underline: true },
          { type: 'text', text: 'Strike', strikethrough: true },
          { type: 'text', text: 'Super', superscript: true },
          { type: 'text', text: 'Sub', subscript: true },
          { type: 'text', text: 'Colored', color: '#ff0000' },
          { type: 'text', text: 'Highlighted', backgroundColor: '#00ff00' },
          { type: 'text', text: 'Sized', fontSize: 18 },
          { type: 'text', text: 'Fonted', fontFamily: 'Georgia' },
        ],
      },
    ]);
    expect(result.warnings).toEqual([]);
  });

  it('round-trips each heading level via the style-name map, not a hard-coded styleId', async () => {
    const blocks: DocSection['blocks'] = [1, 2, 3, 4, 5, 6].map((level) => ({
      type: 'heading',
      level: level as 1 | 2 | 3 | 4 | 5 | 6,
      content: [{ type: 'text', text: `H${level}` }],
    }));
    const result = await roundTrip(section(blocks));
    // The `docx` library's built-in heading styles carry colour, size and italic, which the importer now
    // resolves into the runs (see the style-resolution tests).
    const resolved: object[] = [
      { color: '#2e74b5', fontSize: 16 },
      { color: '#2e74b5', fontSize: 13 },
      { color: '#1f4d78', fontSize: 12 },
      { italic: true, color: '#2e74b5' },
      { color: '#2e74b5' },
      { color: '#1f4d78' },
    ];
    expect(result.sections[0].blocks).toEqual(
      [1, 2, 3, 4, 5, 6].map((level) => ({ type: 'heading', level, namedStyleId: `Heading${level}`, content: [{ type: 'text', text: `H${level}`, ...resolved[level - 1] }] }))
    );
  });

  it('round-trips bullet and numbered list items, including nesting level', async () => {
    const result = await roundTrip(
      section([
        { type: 'listItem', listType: 'bullet', level: 0, content: [{ type: 'text', text: 'Bullet' }] },
        { type: 'listItem', listType: 'bullet', level: 1, content: [{ type: 'text', text: 'Nested bullet' }] },
        { type: 'listItem', listType: 'numbered', level: 0, content: [{ type: 'text', text: 'Numbered' }] },
      ])
    );
    expect(result.sections[0].blocks).toEqual([
      { type: 'listItem', listType: 'bullet', level: 0, content: [{ type: 'text', text: 'Bullet' }] },
      { type: 'listItem', listType: 'bullet', level: 1, content: [{ type: 'text', text: 'Nested bullet' }] },
      { type: 'listItem', listType: 'numbered', level: 0, content: [{ type: 'text', text: 'Numbered' }] },
    ]);
  });

  it('round-trips a table with a colspan and a vertical rowspan merge', async () => {
    const result = await roundTrip(
      section([
        {
          type: 'table',
          rows: [
            { cells: [{ content: [{ type: 'text', text: 'A1' }], colspan: 2 }] },
            {
              cells: [
                { content: [{ type: 'text', text: 'B1' }], rowspan: 2 },
                { content: [{ type: 'text', text: 'B2' }] },
              ],
            },
            { cells: [{ content: [{ type: 'text', text: 'C2' }] }] },
          ],
        },
      ])
    );
    // The writer's default grid comes back as explicit single borders and the grid widths, on top of the structure.
    expect(result.sections[0].blocks).toMatchObject([
      {
        type: 'table',
        rows: [
          { cells: [{ content: [{ type: 'text', text: 'A1' }], colspan: 2 }] },
          {
            cells: [
              { content: [{ type: 'text', text: 'B1' }], rowspan: 2 },
              { content: [{ type: 'text', text: 'B2' }] },
            ],
          },
          { cells: [{ content: [{ type: 'text', text: 'C2' }] }] },
        ],
      },
    ]);
    const table = result.sections[0].blocks[0] as Extract<DocBlock, { type: 'table' }>;
    expect(table.borders?.top?.style).toBe('single');
    expect(table.borders?.insideV?.style).toBe('single');
    expect(table.colWidthsPx).toHaveLength(2);
  });

  it('round-trips an external hyperlink', async () => {
    const result = await roundTrip(
      section([{ type: 'paragraph', content: [{ type: 'text', text: 'Before ' }, { type: 'link', text: 'Spindle', href: 'https://example.com' }] }])
    );
    expect(result.sections[0].blocks).toEqual([
      {
        type: 'paragraph',
        content: [
          { type: 'text', text: 'Before ' },
          // The "Hyperlink" character style the writer applies (blue, underlined) is resolved into the run.
          { type: 'link', text: 'Spindle', href: 'https://example.com', color: '#0563c1', underline: true },
        ],
      },
    ]);
  });

  it('round-trips a block-level image and an inline image mixed with text', async () => {
    const result = await roundTrip(
      section([
        { type: 'image', src: ONE_PX_PNG, width: 10, height: 10, alt: 'dot' },
        { type: 'paragraph', content: [{ type: 'text', text: 'Caption: ' }, { type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }] },
      ])
    );
    expect(result.sections[0].blocks).toEqual([
      { type: 'image', src: ONE_PX_PNG, width: 10, height: 10, alt: 'dot' },
      { type: 'paragraph', content: [{ type: 'text', text: 'Caption: ' }, { type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }] },
    ]);
    expect(result.warnings).toEqual([]);
  });

  it('reads an image wrapped in mc:AlternateContent (LibreOffice output)', async () => {
    const exported = await exportDocx(
      section([{ type: 'paragraph', content: [{ type: 'text', text: 'Caption: ' }, { type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }] }])
    );
    const zip = await JSZip.loadAsync(exported.bytes);
    let xml = await zip.file('word/document.xml')!.async('string');
    xml = xml
      .replace(/<w:document /, '<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ')
      .replace(/(<w:drawing>[\s\S]*?<\/w:drawing>)/, '<mc:AlternateContent><mc:Choice Requires="wps">$1</mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>');
    expect(xml).toContain('<mc:AlternateContent>');
    zip.file('word/document.xml', xml);
    const result = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(result.sections[0].blocks).toEqual([
      { type: 'paragraph', content: [{ type: 'text', text: 'Caption: ' }, { type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }] },
    ]);
  });

  it('round-trips a page break and a horizontal rule', async () => {
    const result = await roundTrip(
      section([
        { type: 'paragraph', content: [{ type: 'text', text: 'Before' }] },
        { type: 'pageBreak' },
        { type: 'horizontalRule' },
        { type: 'paragraph', content: [{ type: 'text', text: 'After' }] },
      ])
    );
    expect(result.sections[0].blocks).toEqual([
      { type: 'paragraph', content: [{ type: 'text', text: 'Before' }] },
      { type: 'pageBreak' },
      { type: 'horizontalRule' },
      { type: 'paragraph', content: [{ type: 'text', text: 'After' }] },
    ]);
  });

  describe('page breaks stored inside content paragraphs (#76)', () => {
    const BR = '<w:r><w:br w:type="page"/></w:r>';
    const run = (t: string) => `<w:r><w:t>${t}</w:t></w:r>`;

    async function importBody(paragraphsXml: string) {
      const exported = await exportDocx(section([{ type: 'paragraph', content: [{ type: 'text', text: 'x' }] }]));
      const zip = await JSZip.loadAsync(exported.bytes);
      let xml = await zip.file('word/document.xml')!.async('string');
      xml = xml.replace(/<w:p>[\s\S]*?<\/w:p>/, paragraphsXml);
      zip.file('word/document.xml', xml);
      return importDocx(await zip.generateAsync({ type: 'uint8array' }));
    }

    it('trailing break run (LibreOffice) becomes a pageBreak block after the paragraph', async () => {
      const result = await importBody(`<w:p>${run('One')}${BR}</w:p><w:p>${run('Two')}</w:p>`);
      expect(result.sections[0].blocks).toEqual([
        { type: 'paragraph', content: [{ type: 'text', text: 'One' }] },
        { type: 'pageBreak' },
        { type: 'paragraph', content: [{ type: 'text', text: 'Two' }] },
      ]);
    });

    it('a break after a block-level image keeps the image and adds the pageBreak', async () => {
      const exported = await exportDocx(section([{ type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }]));
      const zip = await JSZip.loadAsync(exported.bytes);
      let xml = await zip.file('word/document.xml')!.async('string');
      xml = xml.replace(/<\/w:drawing><\/w:r>/, `</w:drawing></w:r>${BR}`);
      expect(xml).toContain('w:type="page"');
      zip.file('word/document.xml', xml);
      const result = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
      expect(result.sections[0].blocks.map((b) => b.type)).toEqual(['image', 'pageBreak']);
    });

    it('a leading break run (page-break-before) puts the pageBreak first', async () => {
      const result = await importBody(`<w:p>${BR}${run('Two')}</w:p>`);
      expect(result.sections[0].blocks).toEqual([
        { type: 'pageBreak' },
        { type: 'paragraph', content: [{ type: 'text', text: 'Two' }] },
      ]);
    });

    it('an otherwise empty break paragraph is still a single pageBreak', async () => {
      const result = await importBody(`<w:p>${BR}</w:p>`);
      expect(result.sections[0].blocks).toEqual([{ type: 'pageBreak' }]);
    });
  });

  it('round-trips paragraph alignment, spacing, and indent', async () => {
    const result = await roundTrip(
      section([
        {
          type: 'paragraph',
          content: [{ type: 'text', text: 'Styled' }],
          alignment: 'center',
          spaceBeforePt: 12,
          spaceAfterPt: 6,
          leftIndentPt: 24,
          firstLineIndentPt: 18,
        },
      ])
    );
    expect(result.sections[0].blocks).toEqual([
      {
        type: 'paragraph',
        content: [{ type: 'text', text: 'Styled' }],
        alignment: 'center',
        spaceBeforePt: 12,
        spaceAfterPt: 6,
        leftIndentPt: 24,
        firstLineIndentPt: 18,
      },
    ]);
  });

  it('round-trips page size and margins', async () => {
    const result = await roundTrip(
      section([{ type: 'paragraph', content: [] }], {
        pageWidthPx: 900,
        pageHeightPx: 1200,
        marginTopPx: 48,
        marginRightPx: 72,
        marginBottomPx: 48,
        marginLeftPx: 72,
      })
    );
    const s = result.sections[0];
    expect(s.pageWidthPx).toBe(900);
    expect(s.pageHeightPx).toBe(1200);
    expect(s.marginTopPx).toBe(48);
    expect(s.marginRightPx).toBe(72);
  });

  it('round-trips landscape orientation and the on-page size', async () => {
    const result = await roundTrip(
      section([{ type: 'paragraph', content: [] }], { pageWidthPx: 1056, pageHeightPx: 816, orientation: 'landscape' })
    );
    expect(result.sections[0].orientation).toBe('landscape');
    expect(result.sections[0].pageWidthPx).toBe(1056);
    expect(result.sections[0].pageHeightPx).toBe(816);
  });

  describe('text that used to be dropped', () => {
    const W = 'xmlns:w="x"';
    async function importXml(bodyXml: string) {
      const exported = await exportDocx(section([{ type: 'paragraph', content: [{ type: 'text', text: 'x' }] }]));
      const zip = await JSZip.loadAsync(exported.bytes);
      let xml = await zip.file('word/document.xml')!.async('string');
      xml = xml.replace(/<w:p>[\s\S]*?<\/w:p>/, bodyXml);
      zip.file('word/document.xml', xml);
      return importDocx(await zip.generateAsync({ type: 'uint8array' }));
    }
    const r = (inner: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}${inner}</w:r>`;
    const t = (text: string) => `<w:t xml:space="preserve">${text}</w:t>`;
    const para = (inner: string) => `<w:p ${W}>${inner}</w:p>`;
    const firstText = async (xml: string) => (await importXml(xml)).sections[0].blocks[0];

    it('imports a tab as a tab character, keeping the surrounding text and formatting', async () => {
      const block = await firstText(para(r(`${t('Chapter one')}<w:tab/>${t('1')}`, '<w:b/>')));
      expect(block).toEqual({ type: 'paragraph', content: [{ type: 'text', text: 'Chapter one\t1', bold: true }] });
    });

    it('imports an absolute-position tab as a tab character', async () => {
      const block = await firstText(para(r(`${t('a')}<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/>${t('b')}`)));
      expect(block).toMatchObject({ content: [{ type: 'text', text: 'a\tb' }] });
    });

    it.each([
      ['a plain break', '<w:br/>'],
      ['a text-wrapping break', '<w:br w:type="textWrapping"/>'],
      ['a column break', '<w:br w:type="column"/>'],
      ['a carriage return', '<w:cr/>'],
    ])('imports %s as a line break', async (_name, br) => {
      const block = await firstText(para(r(`${t('one')}${br}${t('two')}`)));
      expect(block).toMatchObject({ content: [{ type: 'text', text: 'one\ntwo' }] });
    });

    it('keeps a page break structural and out of the text', async () => {
      const result = await importXml(para(r(`${t('one')}`)) + para(r('<w:br w:type="page"/>')));
      expect(result.sections[0].blocks).toEqual([
        { type: 'paragraph', content: [{ type: 'text', text: 'one' }] },
        { type: 'pageBreak' },
      ]);
    });

    it('imports soft and non-breaking hyphens as their Unicode characters', async () => {
      const block = await firstText(para(r(`${t('co')}<w:softHyphen/>${t('op')}<w:noBreakHyphen/>${t('erate')}`)));
      expect(block).toMatchObject({ content: [{ type: 'text', text: 'co\u00adop\u2011erate' }] });
    });

    it('keeps tabs and breaks inside a hyperlink run', async () => {
      const block = await firstText(para(`<w:hyperlink w:anchor="x">${r(`${t('a')}<w:tab/>${t('b')}`)}</w:hyperlink>`));
      expect(block).toMatchObject({ content: [{ type: 'link', href: '#x', text: 'a\tb' }] });
    });

    it('keeps tabs and breaks of a field result in the field, not in the text around it', async () => {
      const block = await firstText(
        para(
          `${r('<w:fldChar w:fldCharType="begin"/>')}${r('<w:instrText> PAGEREF _Toc1 </w:instrText>')}${r('<w:fldChar w:fldCharType="separate"/>')}` +
            `${r(`${t('Chapter')}<w:tab/>${t('4')}<w:br/>`)}${r('<w:fldChar w:fldCharType="end"/>')}${r(t('after'))}`
        )
      );
      expect(block).toMatchObject({ content: [{ type: 'field', kind: 'PAGEREF', result: 'Chapter\t4\n' }, { type: 'text', text: 'after' }] });
    });

    it('keeps the text of smart tags, including a hyperlink inside one and nested tags', async () => {
      const block = await firstText(
        para(
          r(t('Visit ')) +
            `<w:smartTag w:uri="urn:x" w:element="place"><w:smartTagPr><w:attr w:name="a" w:val="b"/></w:smartTagPr>${r(t('Paris'))}` +
            `<w:smartTag w:uri="urn:x" w:element="city">${r(t(' France'))}</w:smartTag></w:smartTag>`
        )
      );
      expect(block).toMatchObject({ content: [{ type: 'text', text: 'Visit ' }, { type: 'text', text: 'Paris' }, { type: 'text', text: ' France' }] });
    });

    it('keeps the text of inline custom XML, and an inline content control around its text', async () => {
      const block = await firstText(
        para(
          `<w:customXml w:element="e"><w:customXmlPr/>${r(t('custom '))}</w:customXml>` +
            `<w:sdt><w:sdtPr><w:alias w:val="A"/></w:sdtPr><w:sdtContent>${r(t('control'))}</w:sdtContent></w:sdt>`
        )
      );
      expect(block).toMatchObject({
        content: [
          { type: 'text', text: 'custom ' },
          { type: 'contentControl', control: { type: 'richText', alias: 'A' }, content: [{ type: 'text', text: 'control' }] },
        ],
      });
    });

    it('keeps tracked insertions and deletions as marked text', async () => {
      const block = await firstText(
        para(r(t('kept ')) + `<w:ins w:id="1" w:author="A" w:date="2024-01-02T03:04:05Z">${r(t('inserted '))}</w:ins>` + `<w:del w:id="2" w:author="B"><w:r><w:delText>deleted</w:delText></w:r></w:del>` + r(t('end')))
      );
      expect(block).toMatchObject({
        content: [
          { type: 'text', text: 'kept ' },
          { type: 'text', text: 'inserted ', revision: { insertion: { id: '1', author: 'A', date: '2024-01-02T03:04:05Z' } } },
          { type: 'text', text: 'deleted', revision: { deletion: { id: '2', author: 'B' } } },
          { type: 'text', text: 'end' },
        ],
      });
    });

    it('keeps the base text of a ruby annotation and drops the reading marks', async () => {
      const block = await firstText(
        para(
          r(
            `<w:ruby><w:rubyPr/><w:rt><w:r>${t('かん')}</w:r></w:rt><w:rubyBase><w:r>${t('漢')}</w:r></w:rubyBase></w:ruby>`
          ) + r(t('字'))
        )
      );
      expect(block).toMatchObject({ content: [{ type: 'text', text: '漢' }, { type: 'text', text: '字' }] });
    });

    it('keeps block-level content controls as wrappers around their blocks, and the text of custom XML', async () => {
      const result = await importXml(
        `<w:sdt ${W}><w:sdtPr/><w:sdtContent><w:p>${r(t('inside'))}</w:p></w:sdtContent></w:sdt>` +
          `<w:customXml ${W} w:element="e"><w:p>${r(t('wrapped'))}</w:p></w:customXml>` +
          `<w:sdt ${W}><w:sdtContent><w:sdt><w:sdtContent><w:p>${r(t('nested'))}</w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt>`
      );
      const paragraph = (text: string) => ({ type: 'paragraph', content: [{ type: 'text', text }] });
      expect(result.sections[0].blocks).toEqual([
        { type: 'contentControl', control: { type: 'richText' }, blocks: [paragraph('inside')] },
        paragraph('wrapped'),
        {
          type: 'contentControl',
          control: { type: 'richText' },
          blocks: [{ type: 'contentControl', control: { type: 'richText' }, blocks: [paragraph('nested')] }],
        },
      ]);
    });

    it('keeps the text of content controls around table rows and cells, and keeps one around cell content', async () => {
      const cell = (inner: string) => `<w:tc><w:p>${r(t(inner))}</w:p></w:tc>`;
      const result = await importXml(
        `<w:tbl ${W}><w:tblPr/><w:tblGrid/>` +
          `<w:sdt><w:sdtContent><w:tr>${cell('row-level')}</w:tr></w:sdtContent></w:sdt>` +
          `<w:tr><w:sdt><w:sdtContent>${cell('cell-level')}</w:sdtContent></w:sdt>` +
          `<w:tc><w:sdt><w:sdtContent><w:p>${r(t('in-cell'))}</w:p></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>`
      );
      const table = result.sections[0].blocks[0];
      expect(table.type).toBe('table');
      if (table.type !== 'table') return;
      expect(table.rows.map((row) => row.cells.map((c) => c.content))).toEqual([
        [[{ type: 'text', text: 'row-level' }]],
        [[{ type: 'text', text: 'cell-level' }], [{ type: 'text', text: 'in-cell' }]],
      ]);
      // A control around the content of a cell stays a wrapper in the cell's blocks.
      expect(table.rows[1].cells[1].blocks).toEqual([
        { type: 'contentControl', control: { type: 'richText' }, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'in-cell' }] }] },
      ]);
    });

    it('keeps a content control in a header', async () => {
      const exported = await exportDocx(
        section([{ type: 'paragraph', content: [] }], { header: { paragraphs: [{ content: [{ type: 'text', text: 'h' }] }] } })
      );
      const zip = await JSZip.loadAsync(exported.bytes);
      const name = Object.keys(zip.files).find((n) => /^word\/header\d*\.xml$/.test(n))!;
      let xml = await zip.file(name)!.async('string');
      xml = xml.replace(/<w:p>[\s\S]*?<\/w:p>/, `<w:sdt><w:sdtContent><w:p><w:r><w:t>Header</w:t><w:tab/><w:t>Right</w:t></w:r></w:p></w:sdtContent></w:sdt>`);
      zip.file(name, xml);
      const result = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
      expect(result.sections[0].header?.paragraphs).toEqual([{ content: [{ type: 'text', text: 'Header\tRight' }] }]);
    });
  });

  describe('tabs and line breaks round trip', () => {
    it('writes \\t as w:tab and \\n as w:br, and reads them back with their formatting', async () => {
      const input = section([
        {
          type: 'paragraph',
          content: [
            { type: 'text', text: 'Title\t12\nsecond line', bold: true },
            { type: 'text', text: ' plain' },
            { type: 'link', text: 'a\tlink\nhere', href: 'https://example.com/' },
          ],
        },
      ]);
      const exported = await exportDocx(input);
      const xml = await (await JSZip.loadAsync(exported.bytes)).file('word/document.xml')!.async('string');
      expect(xml).toContain('<w:tab/>');
      expect(xml).toContain('<w:br/>');
      expect(xml).not.toMatch(/<w:t[^>]*>[^<]*[\t\n][^<]*<\/w:t>/);
      const result = await importDocx(exported.bytes);
      const inlines = (result.sections[0].blocks[0] as { content: Array<{ type: string; text?: string; bold?: boolean; href?: string }> }).content;
      expect(inlines.map((i) => i.text).join('|')).toBe('Title\t12\nsecond line| plain|a\tlink\nhere');
      expect(inlines.filter((i) => i.bold).map((i) => i.text)).toEqual(['Title\t12\nsecond line']);
      expect(inlines[inlines.length - 1]).toMatchObject({ type: 'link', href: 'https://example.com/' });
    });

    it('writes a line break at the start and end of a run, and consecutive tabs', async () => {
      const exported = await exportDocx(section([{ type: 'paragraph', content: [{ type: 'text', text: '\n\t\tx\n' }] }]));
      const result = await importDocx(exported.bytes);
      expect(result.sections[0].blocks[0]).toEqual({ type: 'paragraph', content: [{ type: 'text', text: '\n\t\tx\n' }] });
    });

    it('writes tabs and breaks in header text', async () => {
      const exported = await exportDocx(
        section([{ type: 'paragraph', content: [] }], { header: { paragraphs: [{ content: [{ type: 'text', text: 'Left\tRight\nNext' }] }] } })
      );
      const result = await importDocx(exported.bytes);
      expect(result.sections[0].header?.paragraphs[0].content).toEqual([{ type: 'text', text: 'Left\tRight\nNext' }]);
    });
  });

  describe('drawing kinds and warnings', () => {
    const NS = 'xmlns:wp="x" xmlns:a="x" xmlns:r="x" xmlns:pic="x" xmlns:wps="x" xmlns:o="x"';
    const drawing = (uri: string, inner: string, kind: 'inline' | 'anchor' = 'inline') =>
      `<w:p><w:r><w:drawing ${NS}><wp:${kind}><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="Obj"/>` +
      `<a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic></wp:${kind}></w:drawing></w:r></w:p>`;
    const OOXML = 'http://schemas.openxmlformats.org/drawingml/2006';
    const picture = (blipAttrs: string, kind: 'inline' | 'anchor' = 'inline') =>
      drawing(`${OOXML}/picture`, `<pic:pic><pic:blipFill><a:blip ${blipAttrs}/></pic:blipFill></pic:pic>`, kind);

    async function importBody(paragraphsXml: string, extraRel?: string) {
      const exported = await exportDocx(section([{ type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }]));
      const zip = await JSZip.loadAsync(exported.bytes);
      let xml = await zip.file('word/document.xml')!.async('string');
      xml = xml.replace(/<w:p>[\s\S]*?<\/w:p>/, paragraphsXml);
      zip.file('word/document.xml', xml);
      if (extraRel) {
        const relsPath = 'word/_rels/document.xml.rels';
        const rels = await zip.file(relsPath)!.async('string');
        zip.file(relsPath, rels.replace('</Relationships>', `${extraRel}</Relationships>`));
      }
      return importDocx(await zip.generateAsync({ type: 'uint8array' }));
    }

    it('imports a resolvable inline picture with no warning', async () => {
      const exported = await exportDocx(section([{ type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }]));
      const result = await importDocx(exported.bytes);
      expect(result.sections[0].blocks[0]).toMatchObject({ type: 'image' });
      expect(result.warnings).toEqual([]);
    });

    it('keeps "could not be resolved" for a picture whose blip target is missing', async () => {
      const result = await importBody(picture('r:embed="rId999"'));
      expect(result.warnings).toEqual(['An embedded image could not be resolved and was skipped.']);
    });

    it('keeps "could not be resolved" for a picture with no blip reference at all', async () => {
      const result = await importBody(picture(''));
      expect(result.warnings).toEqual(['An embedded image could not be resolved and was skipped.']);
    });

    it('gives a linked (r:link) picture its own warning and never resolves it', async () => {
      const rel = '<Relationship Id="rId900" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/linked.png" TargetMode="External"/>';
      const result = await importBody(picture('r:link="rId900"'), rel);
      expect(result.warnings).toEqual(['A linked picture was skipped — images linked from outside the document are never fetched.']);
      expect(result.sections[0].blocks.some((b) => b.type === 'image')).toBe(false);
    });

    it('prefers r:embed over r:link when both are present', async () => {
      const exported = await exportDocx(section([{ type: 'image', src: ONE_PX_PNG, width: 5, height: 5 }]));
      const zip = await JSZip.loadAsync(exported.bytes);
      const xml = await zip.file('word/document.xml')!.async('string');
      const embed = xml.match(/r:embed="([^"]+)"/)![1];
      const result = await importBody(picture(`r:embed="${embed}" r:link="rId900"`));
      expect(result.warnings).toEqual([]);
      expect(result.sections[0].blocks[0]).toMatchObject({ type: 'image' });
    });

    it('keeps the floating-image warning for an anchored picture', async () => {
      const result = await importBody(picture('r:embed="rId1"', 'anchor'));
      expect(result.warnings).toEqual(['A floating (anchored) image was skipped — only inline images are supported.']);
    });

    it.each([
      ['a chart', `${OOXML}/chart`, '<c:chart xmlns:c="x" r:id="rId5"/>', 'A chart was not imported — charts are not supported yet.'],
      ['a chartex chart', 'http://schemas.microsoft.com/office/drawing/2014/chartex', '<cx:chart xmlns:cx="x"/>', 'A chart was not imported — charts are not supported yet.'],
      ['a SmartArt diagram', `${OOXML}/diagram`, '<dgm:relIds xmlns:dgm="x"/>', 'A SmartArt diagram was not imported — diagrams are not supported yet.'],
      ['a shape', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape', '<wps:wsp/>', 'A shape or text box was not imported — shapes and text boxes are not supported yet.'],
      ['a shape group', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup', '<wpg:wgp xmlns:wpg="x"/>', 'A shape or text box was not imported — shapes and text boxes are not supported yet.'],
      ['an OLE preview', 'urn:example:ole', '<o:OLEObject/>', 'An embedded object preview was not imported — embedded objects are not supported yet.'],
      ['an unknown drawing', 'urn:example:other', '<x:thing xmlns:x="x"/>', 'A drawing object of an unsupported kind was not imported.'],
    ])('reports %s as what it is, not as a missing image', async (_name, uri, inner, expected) => {
      const result = await importBody(drawing(uri, inner));
      expect(result.warnings).toEqual([expected]);
    });

    it('reports an anchored non-picture by its kind', async () => {
      const result = await importBody(drawing(`${OOXML}/chart`, '<c:chart xmlns:c="x"/>', 'anchor'));
      expect(result.warnings).toEqual(['A chart was not imported — charts are not supported yet.']);
    });

    it('classifies a drawing inside mc:AlternateContent', async () => {
      const inner = drawing('http://schemas.microsoft.com/office/word/2010/wordprocessingShape', '<wps:wsp/>', 'anchor')
        .replace('<w:p><w:r>', '<w:p><w:r><mc:AlternateContent xmlns:mc="x"><mc:Choice>')
        .replace('</w:r></w:p>', '</mc:Choice></mc:AlternateContent></w:r></w:p>');
      const result = await importBody(inner);
      expect(result.warnings).toEqual(['A shape or text box was not imported — shapes and text boxes are not supported yet.']);
    });
  });

  describe('page size from pgSz', () => {
    async function importPgSz(attrs: string) {
      const exported = await exportDocx(section([{ type: 'paragraph', content: [] }]));
      const zip = await JSZip.loadAsync(exported.bytes);
      let xml = await zip.file('word/document.xml')!.async('string');
      xml = xml.replace(/<w:pgSz[^>]*\/>/, `<w:pgSz ${attrs}/>`);
      zip.file('word/document.xml', xml);
      return (await importDocx(await zip.generateAsync({ type: 'uint8array' }))).sections[0];
    }

    it('reads landscape written with w < h (natural size plus orient)', async () => {
      const s = await importPgSz('w:w="12240" w:h="15840" w:orient="landscape"');
      expect(s).toMatchObject({ orientation: 'landscape', pageWidthPx: 1056, pageHeightPx: 816 });
    });

    it('reads landscape written with w > h', async () => {
      const s = await importPgSz('w:w="15840" w:h="12240" w:orient="landscape"');
      expect(s).toMatchObject({ orientation: 'landscape', pageWidthPx: 1056, pageHeightPx: 816 });
    });

    it('treats w > h without an orient attribute as landscape', async () => {
      const s = await importPgSz('w:w="15840" w:h="12240"');
      expect(s).toMatchObject({ orientation: 'landscape', pageWidthPx: 1056, pageHeightPx: 816 });
    });

    it('reads portrait unchanged', async () => {
      const s = await importPgSz('w:w="12240" w:h="15840" w:orient="portrait"');
      expect(s).toMatchObject({ orientation: 'portrait', pageWidthPx: 816, pageHeightPx: 1056 });
    });

    it('keeps a custom landscape size', async () => {
      const s = await importPgSz('w:w="7200" w:h="14400" w:orient="landscape"');
      expect(s).toMatchObject({ orientation: 'landscape', pageWidthPx: 960, pageHeightPx: 480 });
    });
  });

  it('round-trips multiple sections with independent orientation', async () => {
    const exported = await exportDocx({
      title: 'T',
      sections: [
        { ...BASE_SECTION, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'One' }] }] },
        { ...BASE_SECTION, pageWidthPx: 1056, pageHeightPx: 816, orientation: 'landscape', blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'Two' }] }] },
      ],
    });
    const result = await importDocx(exported.bytes);
    expect(result.sections).toHaveLength(2);
    expect(result.sections[0].orientation).toBe('portrait');
    expect(result.sections[1].orientation).toBe('landscape');
    expect(result.sections[0].blocks).toEqual([{ type: 'paragraph', content: [{ type: 'text', text: 'One' }] }]);
    expect(result.sections[1].blocks).toEqual([{ type: 'paragraph', content: [{ type: 'text', text: 'Two' }] }]);
  });

  it('round-trips header and footer content, including page-number/total-pages fields', async () => {
    const result = await roundTrip(
      section([{ type: 'paragraph', content: [{ type: 'text', text: 'Body' }] }], {
        header: { paragraphs: [{ content: [{ type: 'text', text: 'My Doc' }] }] },
        footer: { paragraphs: [{ content: [{ type: 'pageNumber' }, { type: 'text', text: ' of ' }, { type: 'totalPages' }] }] },
      })
    );
    expect(result.sections[0].header).toEqual({ paragraphs: [{ content: [{ type: 'text', text: 'My Doc' }] }] });
    expect(result.sections[0].footer).toEqual({
      paragraphs: [{ content: [{ type: 'pageNumber' }, { type: 'text', text: ' of ' }, { type: 'totalPages' }] }],
    });
  });

  it('reports no warnings for a document it fully understands', async () => {
    const result = await roundTrip(section([{ type: 'paragraph', content: [{ type: 'text', text: 'Plain' }] }]));
    expect(result.warnings).toEqual([]);
  });
});
