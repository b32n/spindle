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

/** A minimal .docx of `paragraphs` one-to-two line paragraphs, with a heading every 40th. */
export async function bigDocx(paragraphs: number): Promise<Buffer> {
  const body: string[] = [];
  for (let i = 1; i <= paragraphs; i++) {
    const text = `Paragraph ${i}: the quick brown fox jumps over the lazy dog and keeps running along the river bank.`;
    body.push(
      i % 40 === 1
        ? `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Section ${Math.ceil(i / 40)}</w:t></w:r></w:p>`
        : `<w:p><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`
    );
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
