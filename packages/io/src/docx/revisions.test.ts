import JSZip from 'jszip';
import { exportDocx } from './writer';
import { importDocx } from './reader';
import { detectDocxLoss } from './loss';
import { parseXml, walk } from './xml-tree';
import type { DocBlock, DocInline, DocSection } from '@b32nio/spindle-shared/export/docx';

const BASE: Omit<DocSection, 'blocks'> = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginTopPx: 96,
  marginRightPx: 96,
  marginBottomPx: 96,
  marginLeftPx: 96,
  orientation: 'portrait',
};
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
type Paragraph = Extract<DocBlock, { type: 'paragraph' }>;

const t = (text: string) => `<w:t xml:space="preserve">${text}</w:t>`;
const r = (inner: string, props = '') => `<w:r>${props && `<w:rPr>${props}</w:rPr>`}${inner}</w:r>`;
const ins = (id: string, author: string, inner: string, date = '') => `<w:ins w:id="${id}" w:author="${author}"${date && ` w:date="${date}"`}>${inner}</w:ins>`;
const del = (id: string, author: string, inner: string) => `<w:del w:id="${id}" w:author="${author}">${inner}</w:del>`;
const delRun = (text: string) => r(`<w:delText xml:space="preserve">${text}</w:delText>`);

async function importBody(bodyXml: string) {
  const exported = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'x' }] }] }] });
  const zip = await JSZip.loadAsync(exported.bytes);
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace(/<w:p>[\s\S]*?<\/w:p>/, bodyXml));
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  return { result: await importDocx(bytes), bytes };
}
const inlinesOf = async (paragraphXml: string): Promise<DocInline[]> => {
  const { result } = await importBody(`<w:p ${W}>${paragraphXml}</w:p>`);
  return (result.sections[0].blocks[0] as Paragraph).content;
};

async function exportXml(content: DocInline[]) {
  const out = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [{ type: 'paragraph', content }] }] });
  const zip = await JSZip.loadAsync(out.bytes);
  return { xml: await zip.file('word/document.xml')!.async('string'), bytes: out.bytes };
}

const A = { id: '1', author: 'Ada', date: '2024-05-06T07:08:09Z' };
const B = { id: '2', author: 'Bo' };

describe('reading tracked changes', () => {
  it('keeps inserted and deleted text with author, date and id', async () => {
    const content = await inlinesOf(r(t('keep ')) + ins('1', 'Ada', r(t('new ')), '2024-05-06T07:08:09Z') + del('2', 'Bo', delRun('old ')) + r(t('end')));
    expect(content).toEqual([
      { type: 'text', text: 'keep ' },
      { type: 'text', text: 'new ', revision: { insertion: A } },
      { type: 'text', text: 'old ', revision: { deletion: B } },
      { type: 'text', text: 'end' },
    ]);
  });

  it('reads text inserted by one author and deleted by another as having both', async () => {
    const content = await inlinesOf(ins('1', 'Ada', del('2', 'Bo', delRun('both')), '2024-05-06T07:08:09Z'));
    expect(content).toEqual([{ type: 'text', text: 'both', revision: { insertion: A, deletion: B } }]);
  });

  it('keeps the formatting of runs inside a change, and gives all runs of one change the same id', async () => {
    const content = await inlinesOf(ins('7', 'Ada', r(t('a'), '<w:b/>') + r(t('b'), '<w:i/>')));
    expect(content).toMatchObject([
      { type: 'text', text: 'a', bold: true, revision: { insertion: { id: '7', author: 'Ada' } } },
      { type: 'text', text: 'b', italic: true, revision: { insertion: { id: '7', author: 'Ada' } } },
    ]);
  });

  it('puts a change inside a hyperlink on the link, and one around a hyperlink on the link it holds', async () => {
    const inside = await inlinesOf(`<w:hyperlink w:anchor="here">${ins('3', 'Ada', r(t('go')))}</w:hyperlink>`);
    expect(inside).toMatchObject([{ type: 'link', text: 'go', href: '#here', revision: { insertion: { id: '3' } } }]);
    const around = await inlinesOf(ins('4', 'Ada', `<w:hyperlink w:anchor="there">${r(t('go'))}</w:hyperlink>`));
    expect(around).toMatchObject([{ type: 'link', text: 'go', href: '#there', revision: { insertion: { id: '4' } } }]);
  });

  it('reads tabs and line breaks in deleted text, and changes through custom xml, smart tags and content controls', async () => {
    const content = await inlinesOf(
      del('5', 'Bo', r('<w:delText>a</w:delText><w:tab/><w:br/><w:delText>b</w:delText>')) +
        `<w:smartTag w:uri="x" w:element="y">${ins('6', 'Ada', r(t('tagged')))}</w:smartTag>` +
        ins('8', 'Ada', `<w:sdt><w:sdtPr><w:alias w:val="c"/></w:sdtPr><w:sdtContent>${r(t('ctl'))}</w:sdtContent></w:sdt>`)
    );
    expect(content[0]).toMatchObject({ type: 'text', text: 'a\t\nb', revision: { deletion: { id: '5' } } });
    expect(content[1]).toMatchObject({ text: 'tagged', revision: { insertion: { id: '6' } } });
    expect(content[2]).toMatchObject({ type: 'contentControl' });
    expect((content[2] as { content: DocInline[] }).content[0]).toMatchObject({ text: 'ctl', revision: { insertion: { id: '8' } } });
  });

  it('gives a change without an id or author a number of its own and "Unknown"', async () => {
    const content = await inlinesOf('<w:ins>' + r(t('x')) + '</w:ins><w:del>' + delRun('y') + '</w:del>');
    expect(content).toMatchObject([
      { revision: { insertion: { id: 'r1', author: 'Unknown' } } },
      { revision: { deletion: { id: 'r2', author: 'Unknown' } } },
    ]);
  });

  it('keeps changes in table cells', async () => {
    const { result } = await importBody(`<w:tbl ${W}><w:tr><w:tc><w:p>${ins('1', 'Ada', r(t('cell')))}</w:p></w:tc></w:tr></w:tbl>`);
    const table = result.sections[0].blocks[0] as Extract<DocBlock, { type: 'table' }>;
    expect(JSON.stringify(table)).toContain('"revision":{"insertion":{"id":"1","author":"Ada"}}');
  });

  it('keeps an inserted field as it is and drops a deleted one, saying so', async () => {
    const field = (instr: string, tag: 'ins' | 'del') => {
      const wrap = tag === 'ins' ? ins : del;
      const text = tag === 'ins' ? t : (s: string) => `<w:delText>${s}</w:delText>`;
      const instrTag = tag === 'ins' ? 'instrText' : 'delInstrText';
      return wrap(
        '9',
        'Ada',
        r('<w:fldChar w:fldCharType="begin"/>') + r(`<w:${instrTag}> ${instr} </w:${instrTag}>`) + r('<w:fldChar w:fldCharType="separate"/>') + r(text('3')) + r('<w:fldChar w:fldCharType="end"/>')
      );
    };
    const inserted = await importBody(`<w:p ${W}>${field('REF a', 'ins')}</w:p>`);
    expect((inserted.result.sections[0].blocks[0] as Paragraph).content).toMatchObject([{ type: 'field', kind: 'REF' }]);
    expect(inserted.result.warnings.some((w) => /insertion around the field "REF a"/.test(w))).toBe(true);
    const deleted = await importBody(`<w:p ${W}>${field('REF a', 'del')}</w:p>`);
    expect((deleted.result.sections[0].blocks[0] as Paragraph).content).toEqual([]);
    expect(deleted.result.warnings.some((w) => /deletion around the field "REF a" .* dropped/.test(w))).toBe(true);
  });

  it('drops a deleted image and a deleted note reference with a warning', async () => {
    const drawing = '<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="9525" cy="9525"/></wp:inline></w:drawing>';
    const { result } = await importBody(`<w:p ${W}>${del('1', 'Bo', r(drawing) + r('<w:footnoteReference w:id="1"/>'))}</w:p>`);
    expect((result.sections[0].blocks[0] as Paragraph).content).toEqual([]);
    expect(result.warnings.some((w) => /deleted image was not kept/.test(w))).toBe(true);
    expect(result.warnings.some((w) => /deleted footnote reference was not kept/.test(w))).toBe(true);
  });
});

describe('tracked changes the importer does not keep as changes', () => {
  const FIELD_BEGIN = r('<w:fldChar w:fldCharType="begin"/>') + r('<w:instrText> REF a </w:instrText>') + r('<w:fldChar w:fldCharType="separate"/>');
  const FIELD_END = r('<w:fldChar w:fldCharType="end"/>');

  it('numbers changes without an id across paragraphs, so no two changes share an id', async () => {
    const { result } = await importBody(`<w:p ${W}><w:ins>${r(t('a'))}</w:ins></w:p><w:p ${W}><w:ins>${r(t('b'))}</w:ins><w:del>${delRun('c')}</w:del></w:p>`);
    const ids = result.sections[0].blocks.flatMap((b) => (b as Paragraph).content).map((i) => Object.values((i as unknown as { revision: Record<string, { id: string }> }).revision)[0].id);
    expect(ids).toEqual(['r1', 'r2', 'r3']);
  });

  it('numbers anonymous changes per import, so two imports at once do not share a counter', async () => {
    const body = (text: string) => `<w:p ${W}><w:ins>${r(t(text))}</w:ins></w:p>`;
    const [a, b] = await Promise.all([importBody(body('a')), importBody(body('b'))]);
    for (const { result } of [a, b]) expect((result.sections[0].blocks[0] as Paragraph).content).toMatchObject([{ revision: { insertion: { id: 'r1' } } }]);
  });

  it('puts the text of a field result without its change, and drops deleted text inside a field result', async () => {
    const inside = await importBody(`<w:p ${W}>${ins('5', 'Ada', FIELD_BEGIN + r(t('3')) + FIELD_END)}</w:p>`);
    expect((inside.result.sections[0].blocks[0] as Paragraph).content).toMatchObject([{ type: 'field', result: '3' }]);
    expect(JSON.stringify((inside.result.sections[0].blocks[0] as Paragraph).content)).not.toContain('revision');
    expect(inside.result.warnings.some((w) => /insertion around the field "REF a"/.test(w))).toBe(true);
    // A run that starts the field and holds its result text.
    const sameRun = await importBody(`<w:p ${W}>${ins('5', 'Ada', r('<w:fldChar w:fldCharType="begin"/><w:instrText> REF a </w:instrText><w:fldChar w:fldCharType="separate"/>' + t('3')) + FIELD_END)}</w:p>`);
    expect(JSON.stringify((sameRun.result.sections[0].blocks[0] as Paragraph).content)).not.toContain('revision');
    // Deleted text in the result of a field begun outside the change.
    const deleted = await importBody(`<w:p ${W}>${FIELD_BEGIN}${r(t('3'))}${del('6', 'Bo', delRun('4'))}${FIELD_END}</w:p>`);
    expect(JSON.stringify((deleted.result.sections[0].blocks[0] as Paragraph).content)).not.toContain('"4"');
    expect(deleted.result.warnings.some((w) => /deleted text inside the field "REF a"/.test(w))).toBe(true);
  });

  it('warns about an inserted image and keeps it as ordinary content', async () => {
    const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';
    const exported = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [{ type: 'paragraph', content: [{ type: 'image', src: png, width: 10, height: 10 }] }] }] });
    const zip = await JSZip.loadAsync(exported.bytes);
    const xml = await zip.file('word/document.xml')!.async('string');
    zip.file('word/document.xml', xml.replace(/<w:r>(?:(?!<\/w:r>)[\s\S])*<w:drawing[\s\S]*?<\/w:r>/, (run) => ins('1', 'Ada', run)));
    const result = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
    expect(result.sections[0].blocks).toMatchObject([{ type: 'image' }]);
    expect(JSON.stringify(result.sections[0].blocks[0])).not.toContain('revision');
    expect(result.warnings.some((w) => /insertion of an image was not kept/.test(w))).toBe(true);
  });

  async function importFooter(footerParagraph: string) {
    const exported = await exportDocx({
      title: 'T',
      sections: [{ ...BASE, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'x' }] }], footer: { paragraphs: [{ content: [{ type: 'text', text: 'F' }] }] } }],
    });
    const zip = await JSZip.loadAsync(exported.bytes);
    const name = Object.keys(zip.files).find((n) => /^word\/footer\d*\.xml$/.test(n))!;
    const xml = await zip.file(name)!.async('string');
    zip.file(name, xml.replace(/<w:p>[\s\S]*?<\/w:p>|<w:p [^>]*>[\s\S]*?<\/w:p>/, `<w:p>${footerParagraph}</w:p>`));
    const bytes = await zip.generateAsync({ type: 'uint8array' });
    return { result: await importDocx(bytes), bytes };
  }

  it('drops deleted text in a footer, keeps inserted text as ordinary text and says so', async () => {
    const { result, bytes } = await importFooter(r(t('keep ')) + ins('1', 'Ada', r(t('new '))) + del('2', 'Bo', delRun('old ')) + r(t('end')));
    const footer = result.sections[0].footer!.paragraphs[0].content;
    expect(footer.map((c) => (c as { text: string }).text).join('')).toBe('keep new end');
    expect(JSON.stringify(footer)).not.toContain('revision');
    expect(result.warnings).toEqual(expect.arrayContaining([expect.stringMatching(/deleted text in a header or footer/), expect.stringMatching(/insertion in a header or footer/)]));
    expect((await detectDocxLoss(bytes, { detail: 'full' })).items.find((i) => i.id === 'docx.trackedHeaderFooterChanges')).toMatchObject({ count: 2 });
  });
});

describe('writing tracked changes', () => {
  const content: DocInline[] = [
    { type: 'text', text: 'keep ' },
    { type: 'text', text: 'new ', bold: true, revision: { insertion: A } },
    { type: 'text', text: 'more', revision: { insertion: A } },
    { type: 'text', text: 'old\tgone', revision: { deletion: B } },
    { type: 'text', text: 'both', revision: { insertion: { id: '3', author: 'Cy' }, deletion: { id: '4', author: 'Di' } } },
    { type: 'text', text: ' end' },
  ];

  it('writes w:ins around inserted runs (one element for the runs of a change), w:del with w:delText for deleted ones, and w:ins around w:del for both', async () => {
    const { xml } = await exportXml(content);
    const doc = parseXml(xml);
    const para: string[] = [];
    walk(doc, (n, tag, path) => {
      if ((tag === 'w:ins' || tag === 'w:del') && path[path.length - 1] === 'w:p') para.push(`${tag}#${(n as Record<string, Record<string, string>>)[':@']['@_w:author']}`);
    });
    expect(para).toEqual(['w:ins#Ada', 'w:del#Bo', 'w:ins#Cy']);
    expect(xml).toContain('w:date="2024-05-06T07:08:09Z"');
    expect(xml).toMatch(/<w:ins [^>]*w:author="Ada"[^>]*><w:r>[\s\S]*?<w:t[^>]*>new <\/w:t><\/w:r><w:r>[\s\S]*?<w:t[^>]*>more<\/w:t><\/w:r><\/w:ins>/);
    expect(xml).toMatch(/<w:del [^>]*w:author="Bo"[^>]*><w:r><w:delText[^>]*>old<\/w:delText><w:tab\/><w:delText[^>]*>gone<\/w:delText><\/w:r><\/w:del>/);
    expect(xml).toMatch(/<w:ins [^>]*w:author="Cy"[^>]*><w:del [^>]*w:author="Di"[^>]*><w:r><w:delText[^>]*>both<\/w:delText>/);
    // No text of a deleted run is a w:t, and ids are unique per change.
    expect(xml.match(/<w:del [\s\S]*?<\/w:del>/g)!.join('')).not.toMatch(/<w:t[ >]/);
    const ids = [...xml.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('round-trips every kind of change through export and import', async () => {
    const { bytes } = await exportXml(content);
    const back = (await importDocx(bytes)).sections[0].blocks[0] as Paragraph;
    expect(back.content).toMatchObject([
      { text: 'keep ' },
      { text: 'new ', bold: true, revision: { insertion: A } },
      { text: 'more', revision: { insertion: A } },
      { text: 'old\tgone', revision: { deletion: B } },
      { text: 'both', revision: { insertion: { id: '3', author: 'Cy' }, deletion: { id: '4', author: 'Di' } } },
      { text: ' end' },
    ]);
  });

  it('writes a changed link with the change inside the hyperlink, for external and internal links', async () => {
    const { xml, bytes } = await exportXml([
      { type: 'link', text: 'ext', href: 'https://example.com/', revision: { insertion: A } },
      { type: 'link', text: 'int', href: '#mark', revision: { deletion: B } },
    ]);
    expect(xml).toMatch(/<w:hyperlink [^>]*><w:ins [^>]*><w:r>/);
    expect(xml).toMatch(/<w:hyperlink [^>]*w:anchor="mark"[^>]*><w:del [^>]*><w:r>[\s\S]*?<w:delText/);
    const back = (await importDocx(bytes)).sections[0].blocks[0] as Paragraph;
    expect(back.content).toMatchObject([
      { type: 'link', text: 'ext', href: 'https://example.com/', revision: { insertion: A } },
      { type: 'link', text: 'int', href: '#mark', revision: { deletion: B } },
    ]);
  });

  it('keeps a numeric id when it is free and numbers the others', async () => {
    const { xml } = await exportXml([
      { type: 'text', text: 'a', revision: { insertion: { id: 'r1', author: 'Ada' } } },
      { type: 'text', text: 'b', revision: { insertion: { id: 'r2', author: 'Ada' } } },
      { type: 'text', text: 'c', revision: { insertion: { id: '1', author: 'Ada' } } },
    ]);
    const ids = [...xml.matchAll(/<w:ins w:id="(\d+)"/g)].map((m) => m[1]);
    expect(ids).toHaveLength(3);
    expect(new Set(ids).size).toBe(3);
    expect(ids).toContain('1');
  });

  it('does not write a revision with a run in a field result or a text without a change', async () => {
    const { xml } = await exportXml([{ type: 'field', instruction: 'REF a', kind: 'REF', result: [{ type: 'text', text: '3', revision: { deletion: B } }] }]);
    expect(xml).not.toContain('<w:del ');
    expect(xml).not.toContain('delText');
  });
});

describe('writing the author and date of a change', () => {
  it('leaves out what XML cannot carry and drops a date that is not a date', async () => {
    const { xml } = await exportXml([
      { type: 'text', text: 'a', revision: { insertion: { id: '1', author: 'Ann\u0000\u0008 Lee\uD800', date: 'yesterday' } } },
      { type: 'text', text: 'b', revision: { insertion: { id: '2', author: '\u0001', date: '2024-01-02T03:04:05.678+02:00' } } },
    ]);
    expect(xml).toContain('w:author="Ann Lee"');
    expect(xml).toContain('w:author="Unknown"');
    expect(xml).not.toContain('yesterday');
    expect(xml).toContain('w:date="2024-01-02T03:04:05.678+02:00"');
    expect(() => parseXml(xml)).not.toThrow();
    // eslint-disable-next-line no-control-regex
    expect(xml).not.toMatch(/[\u0000-\u0008]/);
  });
});

describe('writing the author and date of a change: edge cases', () => {
  it('keeps a valid surrogate pair, drops an unpaired low surrogate and accepts a date without seconds', async () => {
    const { xml } = await exportXml([
      { type: 'text', text: 'a', revision: { insertion: { id: '1', author: 'Ann \u{1F600}\uDC00 Lee', date: '2024-01-05T10:00Z' } } },
    ]);
    expect(xml).toContain('w:author="Ann \u{1F600} Lee"');
    expect(xml).toContain('w:date="2024-01-05T10:00Z"');
    expect(() => parseXml(xml)).not.toThrow();
  });
});

describe('a deleted run before a page break', () => {
  it('counts as content, so the break is after the paragraph', async () => {
    const { result } = await importBody(`<w:p ${W}>${del('1', 'Bo', delRun('gone'))}<w:r><w:br w:type="page"/></w:r></w:p>`);
    expect(result.sections[0].blocks.map((b) => b.type)).toEqual(['paragraph', 'pageBreak']);
  });
});

describe('loss report', () => {
  it('no longer counts changes around runs, only those on a paragraph mark or table row', async () => {
    const { bytes } = await importBody(
      `<w:p ${W}><w:pPr><w:rPr><w:ins w:id="1" w:author="A"/></w:rPr></w:pPr>${ins('2', 'A', r(t('x')))}${del('3', 'A', delRun('y'))}</w:p>` +
        `<w:tbl ${W}><w:tr><w:trPr><w:del w:id="4" w:author="A"/></w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl>`
    );
    const items = (await detectDocxLoss(bytes)).items;
    expect(items.find((i) => i.id === 'docx.trackedInsertions')).toMatchObject({ count: 1 });
    expect(items.find((i) => i.id === 'docx.trackedDeletions')).toMatchObject({ count: 1 });
    expect((await detectDocxLoss(bytes, { detail: 'full' })).items.find((i) => i.id === 'docx.trackedInsertions')).toMatchObject({ count: 1 });
  });

  it('reports nothing for a document whose only changes are around runs', async () => {
    const { bytes } = await importBody(`<w:p ${W}>${ins('2', 'A', r(t('x')))}${del('3', 'A', delRun('y'))}</w:p>`);
    expect((await detectDocxLoss(bytes)).items.filter((i) => i.id.startsWith('docx.tracked'))).toEqual([]);
  });

  it('does not count a change inside a text box or a property-change record', async () => {
    const { bytes } = await importBody(
      `<w:p ${W}><w:r><w:pict><w:txbxContent><w:p><w:pPr><w:rPr><w:ins w:id="1" w:author="A"/></w:rPr></w:pPr></w:p></w:txbxContent></w:pict></w:r>` +
        `<w:pPr><w:pPrChange w:id="5" w:author="A"><w:pPr><w:rPr><w:del w:id="6" w:author="A"/></w:rPr></w:pPr></w:pPrChange></w:pPr></w:p>`
    );
    expect((await detectDocxLoss(bytes)).items.filter((i) => i.id === 'docx.trackedInsertions' || i.id === 'docx.trackedDeletions')).toEqual([]);
  });
});

describe('loss report for changes around whole blocks', () => {
  it('counts a change around a paragraph or table of the body or of a cell', async () => {
    const { bytes } = await importBody(
      `<w:ins ${W} w:id="1" w:author="A"><w:p>${r(t('x'))}</w:p></w:ins><w:del ${W} w:id="2" w:author="A"><w:p>${delRun('y')}</w:p></w:del>` +
        `<w:tbl ${W}><w:tr><w:tc><w:ins w:id="3" w:author="A"><w:p/></w:ins><w:p/></w:tc></w:tr></w:tbl>`
    );
    expect((await detectDocxLoss(bytes)).items.find((i) => i.id === 'docx.trackedBlockChanges')).toMatchObject({ count: 3 });
  });
});
