import JSZip from 'jszip';

/** CSV with a header row plus `rows` data rows: id, text, number, formula-free decimal, flag. */
export function bigCsv(rows: number): Buffer {
  const lines = ['id,name,qty,price,flag'];
  for (let i = 1; i <= rows; i++) {
    lines.push(`${i},item ${i},${i % 97},${((i * 7) % 1000) / 10},${i % 2 === 0 ? 'yes' : 'no'}`);
  }
  return Buffer.from(lines.join('\n'));
}

const esc = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;');

/** A minimal .docx of `paragraphs` one-to-two line paragraphs,. */
export async function bigDocx(paragraphs: number): Promise<Buffer> {
  const body: string[] = [];
  for (let i = 1; i <= paragraphs; i++) {
    const text = `Paragraph ${i}: the quick brown fox jumps over the lazy dog and keeps running along the river bank.`;
    body.push(`<w:p><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`);
  }
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body.join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`
  );
  return zip.generateAsync({ type: 'nodebuffer' });
}

const DOCX_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** A minimal .docx around `body` (WordprocessingML for the body of the document). */
async function docxOf(body: string): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${DOCX_NS}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`
  );
  return zip.generateAsync({ type: 'nodebuffer' });
}

/**
 * A .docx with one table whose first cell holds two paragraphs and a nested table (with "Inner one" and "Inner two"),
 * followed by a paragraph "After the table".
 */
export async function nestedTableDocx(): Promise<Buffer> {
  const p = (t: string) => `<w:p><w:r><w:t>${esc(t)}</w:t></w:r></w:p>`;
  const cell = (inner: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
  const inner = `<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/></w:tblGrid><w:tr>${cell(p('Inner one'))}${cell(p('Inner two'))}</w:tr></w:tbl>`;
  const outer = `<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell(p('Cell first') + p('Cell second') + inner + '<w:p/>')}${cell(p('Side'))}</w:tr></w:tbl>`;
  return docxOf(outer + p('After the table'));
}
