import JSZip from 'jszip';
import { exportDocx } from './writer';
import { importDocx } from './reader';
import { detectDocxLoss } from './loss';
import { parseNotes, referencedNoteIds } from './notes';
import type { DocBlock, DocInline, DocNoteRefRun, DocSection } from '@b32nio/spindle-shared/export/docx';

const BASE: Omit<DocSection, 'blocks'> = {
  pageWidthPx: 816,
  pageHeightPx: 1056,
  marginTopPx: 96,
  marginRightPx: 96,
  marginBottomPx: 96,
  marginLeftPx: 96,
  orientation: 'portrait',
};
const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
type Paragraph = Extract<DocBlock, { type: 'paragraph' }>;
type NoteRef = DocNoteRefRun;

const r = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const ref = (kind: 'footnote' | 'endnote', id: number | string, extra = '') => `<w:r><w:${kind}Reference w:id="${id}"${extra}/></w:r>`;
const note = (kind: 'footnote' | 'endnote', id: number | string, paragraphs: string[], attrs = '') =>
  `<w:${kind} w:id="${id}"${attrs}>${paragraphs.map((p, i) => `<w:p>${i === 0 ? `<w:r><w:${kind}Ref/></w:r>` : ''}${p}</w:p>`).join('')}</w:${kind}>`;
const separators = (kind: 'footnote' | 'endnote') =>
  `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}><w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`;

interface Parts {
  body?: string;
  footnotes?: string;
  endnotes?: string;
  footnoteRels?: string;
}

async function docxWith(parts: Parts): Promise<Uint8Array> {
  const exported = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'x' }] }] }] });
  const zip = await JSZip.loadAsync(exported.bytes);
  if (parts.body !== undefined) {
    const xml = await zip.file('word/document.xml')!.async('string');
    zip.file('word/document.xml', xml.replace(/<w:p>[\s\S]*?<\/w:p>/, parts.body));
  }
  if (parts.footnotes !== undefined) zip.file('word/footnotes.xml', `<?xml version="1.0"?><w:footnotes ${NS}>${separators('footnote')}${parts.footnotes}</w:footnotes>`);
  if (parts.endnotes !== undefined) zip.file('word/endnotes.xml', `<?xml version="1.0"?><w:endnotes ${NS}>${separators('endnote')}${parts.endnotes}</w:endnotes>`);
  if (parts.footnoteRels !== undefined) zip.file('word/_rels/footnotes.xml.rels', parts.footnoteRels);
  return zip.generateAsync({ type: 'uint8array' });
}
const p = (inner: string) => `<w:p ${NS}>${inner}</w:p>`;
const importParts = async (parts: Parts) => importDocx(await docxWith(parts));
const firstContent = (result: Awaited<ReturnType<typeof importDocx>>, block = 0): DocInline[] => (result.sections[0].blocks[block] as Paragraph).content;
const refs = (content: DocInline[]): NoteRef[] => content.filter((i): i is NoteRef => i.type === 'noteRef');

describe('reading notes', () => {
  it('turns each reference into a noteRef holding the text of its note', async () => {
    const result = await importParts({
      body: p(r('claim') + ref('footnote', 5) + r(' and') + ref('endnote', 2) + r('.')),
      footnotes: note('footnote', 5, [r(' Source: ledger.')]),
      endnotes: note('endnote', 2, [r(' Disputed.', '<w:b/>')]),
    });
    const content = firstContent(result);
    expect(content.map((i) => i.type)).toEqual(['text', 'noteRef', 'text', 'noteRef', 'text']);
    expect(refs(content)[0]).toMatchObject({ kind: 'footnote', content: [{ type: 'text', text: 'Source: ledger.' }] });
    expect(refs(content)[1]).toMatchObject({ kind: 'endnote', content: [{ type: 'text', text: 'Disputed.', bold: true }] });
    expect(result.warnings).toEqual([]);
  });

  it('numbers the model ids in document order per kind, whatever the ids in the file are', async () => {
    const result = await importParts({
      body: p(ref('footnote', 9) + ref('endnote', 3) + ref('footnote', 4) + ref('footnote', 9)),
      footnotes: note('footnote', 4, [r(' four')]) + note('footnote', 9, [r(' nine')]),
      endnotes: note('endnote', 3, [r(' three')]),
    });
    expect(refs(firstContent(result)).map((n) => [n.kind, n.id, n.content[0].type === 'text' ? n.content[0].text : ''])).toEqual([
      ['footnote', '1', 'nine'],
      ['endnote', '1', 'three'],
      ['footnote', '2', 'four'],
      ['footnote', '3', 'nine'],
    ]);
  });

  it('drops the space after the mark, in the text or in a run of its own, and keeps other leading text', async () => {
    const result = await importParts({
      body: p(ref('footnote', 1) + ref('footnote', 2) + ref('footnote', 3)),
      footnotes: note('footnote', 1, [r(' ') + r('LibreOffice style')]) + note('footnote', 2, [r('no space')]) + note('footnote', 3, [r('  two spaces')]),
    });
    expect(refs(firstContent(result)).map((n) => (n.content[0] as { text: string }).text)).toEqual(['LibreOffice style', 'no space', ' two spaces']);
  });

  it('joins the paragraphs of a note with a line break that keeps the formatting of the text before it', async () => {
    const result = await importParts({
      body: p(ref('footnote', 1)),
      footnotes: note('footnote', 1, [r(' first', '<w:i/>'), r('second')]),
    });
    expect(refs(firstContent(result))[0].content).toEqual([
      { type: 'text', text: 'first\n', italic: true },
      { type: 'text', text: 'second' },
    ]);
  });

  it('puts a line break of its own between paragraphs when the first note paragraph ends in something that is not text', async () => {
    const result = await importParts({
      body: p(ref('footnote', 1)),
      footnotes: note('footnote', 1, [`<w:bookmarkStart w:id="1" w:name="b"/>`, r('second')]),
    });
    expect(refs(firstContent(result))[0].content).toEqual([{ type: 'bookmark', name: 'b' }, { type: 'text', text: '\n' }, { type: 'text', text: 'second' }]);
  });

  it('reads a custom mark and keeps the text after it out of the paragraph text', async () => {
    const result = await importParts({
      body: p(r('a') + `<w:r><w:footnoteReference w:customMarkFollows="1" w:id="1"/><w:t>*</w:t></w:r>` + r('b')),
      footnotes: note('footnote', 1, [r(' starred')]),
    });
    const content = firstContent(result);
    expect(content).toMatchObject([{ type: 'text', text: 'a' }, { type: 'noteRef', customMark: '*', content: [{ text: 'starred' }] }, { type: 'text', text: 'b' }]);
  });

  it('resolves links in a note through the notes part, then the document', async () => {
    const rels = `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId77" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://note.example/a" TargetMode="External"/></Relationships>`;
    const result = await importParts({
      body: p(ref('footnote', 1)),
      footnotes: note('footnote', 1, [r(' see ') + `<w:hyperlink r:id="rId77"><w:r><w:t>here</w:t></w:r></w:hyperlink>`]),
      footnoteRels: rels,
    });
    expect(refs(firstContent(result))[0].content).toMatchObject([{ text: 'see ' }, { type: 'link', text: 'here', href: 'https://note.example/a' }]);
  });

  it('warns about a reference to a note that does not exist and keeps reading', async () => {
    const result = await importParts({ body: p(r('a') + ref('footnote', 42) + ref('endnote', 1) + r('b')), footnotes: note('footnote', 1, [r(' x')]) });
    expect(firstContent(result).map((i) => i.type)).toEqual(['text', 'text']);
    expect(result.warnings).toEqual(['A footnote reference points to a footnote that does not exist.', 'A endnote reference points to a endnote that does not exist.']);
  });

  it('does not keep a reference inside a field result, and warns about one in a header', async () => {
    const field = p(`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${ref('footnote', 1)}${r('cached')}<w:r><w:fldChar w:fldCharType="end"/></w:r>`);
    const result = await importParts({ body: field, footnotes: note('footnote', 1, [r(' n')]) });
    expect(firstContent(result)).toMatchObject([{ type: 'field', kind: 'REF', result: 'cached' }]);

    const headerExport = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks: [{ type: 'paragraph', content: [] }], header: { paragraphs: [{ content: [{ type: 'text', text: 'h' }] }] } }] });
    const headerZip = await JSZip.loadAsync(headerExport.bytes);
    const headerName = Object.keys(headerZip.files).find((n) => /^word\/header\d*\.xml$/.test(n))!;
    headerZip.file(headerName, (await headerZip.file(headerName)!.async('string')).replace(/<w:p>[\s\S]*?<\/w:p>/, `<w:p>${ref('footnote', 1)}</w:p>`));
    const withHeader = await importDocx(await headerZip.generateAsync({ type: 'uint8array' }));
    expect(withHeader.warnings).toContain('A footnote reference in a header, footer or note was not imported.');
  });

  it('keeps a reference in a table cell, and warns about a table inside a note', async () => {
    const cell = `<w:tbl ${NS}><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p>${r('in cell')}${ref('footnote', 1)}</w:p></w:tc></w:tr></w:tbl>`;
    const result = await importParts({
      body: cell,
      footnotes: note('footnote', 1, [r(' n')]).replace('</w:footnote>', '<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:footnote>'),
    });
    const table = result.sections[0].blocks[0] as Extract<DocBlock, { type: 'table' }>;
    expect(refs(table.rows[0].cells[0].content)).toHaveLength(1);
    expect(result.warnings).toContain('A table in a note was not imported.');
  });

  it('skips separator notes, reads a damaged part as far as it goes, and survives one that is not XML', async () => {
    const parsed = parseNotes(`<w:footnotes ${NS}>${separators('footnote')}${note('footnote', 1, [r('a')])}</w:footnotes>`, 'footnote');
    expect([...parsed.keys()]).toEqual(['1']);
    expect(parseNotes(undefined, 'footnote').size).toBe(0);
    expect(parseNotes(`<w:other ${NS}/>`, 'footnote').size).toBe(0);
    expect(parseNotes(`<w:footnotes ${NS}><w:footnote w:id="1"><w:p><w:r><w:t>Broken &bogus; <b></w:t></w:r></w:p></w:footnote>`, 'footnote').size).toBe(1);
    // A repeated id keeps the first note.
    const twice = parseNotes(`<w:footnotes ${NS}>${note('footnote', 1, [r('first')])}${note('footnote', 1, [r('second')])}</w:footnotes>`, 'footnote');
    expect(twice.size).toBe(1);

    const result = await importParts({ body: p(r('a') + ref('footnote', 1)), footnotes: '<w:footnote w:id="1"><w:p><w:r><w:t>Broken &bogus; <b></w:t></w:r></w:p></w:footnote>' });
    expect(refs(firstContent(result))[0].content).toEqual([{ type: 'text', text: 'Broken &bogus; ' }]);
  });

  it('finds the notes a body refers to', () => {
    const xml = `<w:body><w:footnoteReference w:id="3"/><w:footnoteReference w:customMarkFollows="1" w:id="4"/><w:endnoteReference w:id="3"/></w:body>`;
    expect([...referencedNoteIds(xml, 'footnote')]).toEqual(['3', '4']);
    expect([...referencedNoteIds(xml, 'endnote')]).toEqual(['3']);
  });
});

const text = (t: string, extra: Partial<Extract<DocInline, { type: 'text' }>> = {}): DocInline => ({ type: 'text', text: t, ...extra });
const doc = (...content: DocInline[]): DocBlock[] => [{ type: 'paragraph', content }];
const footnote = (id: string, ...content: DocInline[]): DocNoteRefRun => ({ type: 'noteRef', kind: 'footnote', id, content });
const endnote = (id: string, ...content: DocInline[]): DocNoteRefRun => ({ type: 'noteRef', kind: 'endnote', id, content });

async function exportParts(blocks: DocBlock[], extra: { namedStyles?: Record<string, { name: string; kind: 'paragraph' | 'character' }> } = {}) {
  const out = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks }], ...extra });
  const zip = await JSZip.loadAsync(out.bytes);
  const read = async (name: string) => (await zip.file(name)?.async('string')) ?? '';
  return {
    bytes: out.bytes,
    warnings: out.warnings,
    zip,
    document: await read('word/document.xml'),
    footnotes: await read('word/footnotes.xml'),
    endnotes: await read('word/endnotes.xml'),
    footnoteRels: await read('word/_rels/footnotes.xml.rels'),
    settings: await read('word/settings.xml'),
  };
}

describe('writing notes', () => {
  it('writes a superscript reference per note and the notes into their parts, numbered from 1 in document order', async () => {
    const out = await exportParts(doc(text('a'), footnote('x', text('first note')), text('b'), endnote('y', text('an endnote', { bold: true })), footnote('z', text('second note'))));
    expect(out.document).toContain('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>');
    expect(out.document).toContain('<w:footnoteReference w:id="2"/>');
    expect(out.document).toContain('<w:endnoteReference w:id="1"/>');
    expect(out.footnotes).toMatch(/<w:footnote w:type="separator" w:id="-1">.*<w:footnote w:type="continuationSeparator" w:id="0">.*<w:footnote w:id="1">.*first note.*<w:footnote w:id="2">.*second note/s);
    expect(out.footnotes).toContain('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>');
    expect(out.endnotes).toContain('<w:endnote w:id="1">');
    expect(out.endnotes).toContain('<w:endnoteRef/>');
    expect(out.endnotes).toContain('<w:b/>');
    // The part keeps its root element with the namespaces.
    expect(out.footnotes).toMatch(/<w:footnotes [^>]*xmlns:w=/);
  });

  it('tells the settings where the separator notes are, before the compatibility settings, once', async () => {
    const out = await exportParts(doc(footnote('1', text('n')), endnote('1', text('e'))));
    expect(out.settings).toContain('<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr><w:compat>');
    const onlyFootnotes = await exportParts(doc(footnote('1', text('n'))));
    expect(onlyFootnotes.settings).toContain('<w:footnotePr>');
    expect(onlyFootnotes.settings).not.toContain('<w:endnotePr>');
    expect(onlyFootnotes.endnotes).not.toContain('<w:endnote w:id="1">');
  });

  it('leaves the parts alone when the document has no notes', async () => {
    const out = await exportParts(doc(text('plain')));
    expect(out.settings).not.toContain('footnotePr');
    expect(out.footnotes).not.toContain('<w:footnote w:id="1">');
  });

  it('writes a custom mark after the reference and as the start of the note', async () => {
    const out = await exportParts(doc({ ...footnote('1', text('starred')), customMark: '*' }));
    expect(out.document).toMatch(/<w:footnoteReference [^>]*w:customMarkFollows="1"[^>]*\/><w:t xml:space="preserve">\*<\/w:t>/);
    expect(out.footnotes).toContain('<w:r><w:t xml:space="preserve">*</w:t></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>');
    expect(out.footnotes).not.toContain('<w:footnoteRef/>');
  });

  it('writes an empty note, a note with a line break, formatting and no nested notes or controls loss', async () => {
    const out = await exportParts(
      doc(
        footnote('1'),
        footnote('2', text('a\nb', { italic: true }), footnote('3', text('nested')), { type: 'contentControl', control: { type: 'richText' }, content: [footnote('4', text('deep')), text('kept')] })
      )
    );
    expect(out.footnotes).toMatch(/<w:footnote w:id="1"><w:p><w:pPr>.*?<\/w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"\/><\/w:rPr><w:footnoteRef\/><\/w:r><w:r><w:t xml:space="preserve"> <\/w:t><\/w:r><\/w:p><\/w:footnote>/);
    expect(out.footnotes).toContain('<w:br/>');
    expect(out.footnotes).toContain('<w:i/>');
    expect(out.footnotes).toContain('kept');
    expect(out.footnotes).not.toContain('nested');
    expect(out.footnotes).not.toContain('deep');
    expect(out.footnotes.match(/<w:footnote w:id="/g)).toHaveLength(2);
  });

  it('writes the links and pictures of a note into the notes part relationships', async () => {
    const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
    const out = await exportParts(
      doc(footnote('1', text('see '), { type: 'link', text: 'here', href: 'https://note.example/x' }, { type: 'image', src: png, width: 5, height: 5 }), endnote('1', text('plain')))
    );
    expect(out.footnoteRels).toMatch(/Type="[^"]*\/hyperlink" Target="https:\/\/note\.example\/x" TargetMode="External"/);
    const image = out.footnoteRels.match(/Type="[^"]*\/image" Target="(media\/note-footnote-[^"]+)"/);
    expect(image).not.toBeNull();
    expect(out.zip.file(`word/${image![1]}`)).not.toBeNull();
    expect(out.footnoteRels).not.toContain('/styles');
    // The reading side resolves both.
    const result = await importDocx(out.bytes);
    const note = refs(firstContent(result))[0];
    expect(note.content.map((c) => c.type)).toEqual(['text', 'link', 'image']);
  });

  it('keeps named styles used in a note', async () => {
    const out = await exportParts(doc(footnote('1', text('styled', { namedStyleId: 'Accent' }))), { namedStyles: { Accent: { name: 'Accent', kind: 'character' } } });
    expect(out.footnotes).toContain('<w:rStyle w:val="Accent"/>');
    const result = await importDocx(out.bytes);
    expect(Object.keys(result.namedStyles ?? {})).toEqual(['Accent']);
  });

  it('writes the notes of every section, for references in tables, lists and controls', async () => {
    const blocks: DocBlock[] = [
      { type: 'table', rows: [{ cells: [{ content: [text('c'), footnote('1', text('in a cell'))] }] }] },
      { type: 'listItem', listType: 'bullet', level: 0, content: [footnote('2', text('in a list'))] },
      { type: 'contentControl', control: { type: 'richText' }, blocks: [{ type: 'paragraph', content: [endnote('1', text('in a control'))] }] },
    ];
    const out = await exportParts(blocks);
    expect(out.footnotes).toContain('in a cell');
    expect(out.footnotes).toContain('in a list');
    expect(out.endnotes).toContain('in a control');
  });

  it('round-trips references and notes through export and import', async () => {
    const blocks = doc(
      text('a'),
      footnote('1', text('one ', { bold: true }), { type: 'link', text: 'link', href: 'https://example.com/z' }, text(' two\nthree')),
      endnote('1', text('end')),
      { ...footnote('2', text('custom')), customMark: '†' },
      footnote('3')
    );
    const out = await exportDocx({ title: 'T', sections: [{ ...BASE, blocks }] });
    const read = (await importDocx(out.bytes)).sections[0].blocks[0] as Paragraph;
    expect(refs(read.content)).toMatchObject([
      { kind: 'footnote', id: '1', content: [{ text: 'one ', bold: true }, { type: 'link', href: 'https://example.com/z' }, { text: ' two\nthree' }] },
      { kind: 'endnote', id: '1', content: [{ text: 'end' }] },
      { kind: 'footnote', id: '2', customMark: '†', content: [{ text: 'custom' }] },
      { kind: 'footnote', id: '3', content: [] },
    ]);
  });
});

describe('detecting what is lost about notes', () => {
  const lossOf = async (parts: Parts, id: string) => (await detectDocxLoss(await docxWith(parts), { detail: 'full' })).items.find((i) => i.id === id);

  it('reports nothing when every note is referenced from the body', async () => {
    const bytes = { body: p(ref('footnote', 1) + ref('endnote', 1)), footnotes: note('footnote', 1, [r(' a')]), endnotes: note('endnote', 1, [r(' b')]) };
    expect(await lossOf(bytes, 'docx.footnotes')).toBeUndefined();
    expect(await lossOf(bytes, 'docx.endnotes')).toBeUndefined();
  });

  it('reports a note nothing in the body refers to, and one whose only reference is in a text box', async () => {
    const parts = {
      body: p(ref('footnote', 1)) + `<w:p><w:r><w:pict><w:txbxContent><w:p>${ref('footnote', 3)}</w:p></w:txbxContent></w:pict></w:r></w:p>`,
      footnotes: note('footnote', 1, [r(' a')]) + note('footnote', 2, [r(' orphan')]) + note('footnote', 3, [r(' in a text box')]),
      endnotes: note('endnote', 1, [r(' orphan endnote')]),
    };
    expect(await lossOf(parts, 'docx.footnotes')).toMatchObject({ count: 2 });
    expect(await lossOf(parts, 'docx.endnotes')).toMatchObject({ count: 1 });
  });

  it('does not count the separator notes', async () => {
    expect(await lossOf({ body: p(r('x')), footnotes: '' }, 'docx.footnotes')).toBeUndefined();
  });

  it('reports the notes of a part that cannot be read as XML', async () => {
    const bytes = await docxWith({ body: p(ref('footnote', 1)) });
    const zip = await JSZip.loadAsync(bytes);
    zip.file('word/footnotes.xml', '<w:footnote w:id="1"><w:p>not a notes part');
    const loss = await detectDocxLoss(await zip.generateAsync({ type: 'uint8array' }), { detail: 'full' });
    expect(loss.items.find((i) => i.id === 'docx.footnotes')).toMatchObject({ count: 1 });
  });
});
