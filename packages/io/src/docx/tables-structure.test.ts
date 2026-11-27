import JSZip from 'jszip';
import { exportDocx } from './writer';
import { importDocx } from './reader';
import { detectDocxLoss } from './loss';
import type { DocBlock, DocSection } from '@b32nio/spindle-shared/export/docx';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
type Table = Extract<DocBlock, { type: 'table' }>;
type Cell = Table['rows'][number]['cells'][number];

async function docx(body: string, extra: Record<string, string> = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types/>');
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  for (const [name, xml] of Object.entries(extra)) zip.file(name, xml);
  return zip.generateAsync({ type: 'uint8array' });
}
const p = (text: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p>`;
const tc = (inner: string, tcPr = '') => `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}${inner}</w:tc>`;
const tr = (cells: string, trPr = '') => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells}</w:tr>`;
const tbl = (rows: string, tblPr = '', grid = '') => `<w:tbl><w:tblPr>${tblPr}</w:tblPr>${grid ? `<w:tblGrid>${grid}</w:tblGrid>` : ''}${rows}</w:tbl>`;
const importTable = async (body: string, extra: Record<string, string> = {}): Promise<{ table: Table; warnings: string[] }> => {
  const { sections, warnings } = await importDocx(await docx(body, extra));
  return { table: sections[0].blocks[0] as Table, warnings };
};
const text = (s: string) => ({ type: 'text', text: s });

describe('cells with more than one paragraph', () => {
  it('keeps every paragraph in `blocks` and the first one with text in `content`', async () => {
    const { table } = await importTable(tbl(tr(tc(p('') + p('one') + p('two') + p('')))));
    const cell = table.rows[0].cells[0];
    expect(cell.content).toEqual([text('one')]);
    expect(cell.blocks?.map((b) => b.type)).toEqual(['paragraph', 'paragraph', 'paragraph', 'paragraph']);
    expect((cell.blocks?.[2] as { content: unknown }).content).toEqual([text('two')]);
  });

  it('leaves a cell of one plain paragraph without `blocks`', async () => {
    const { table } = await importTable(tbl(tr(tc(p('solo')))));
    expect(table.rows[0].cells[0]).toEqual({ content: [text('solo')] });
  });

  it('keeps an inline image in a cell as part of its paragraph', async () => {
    const png =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
    const body = tbl(
      tr(
        tc(
          `<w:p xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
        )
      )
    );
    const bytes = await docx(body, {
      'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/i.png"/></Relationships>`,
    });
    const zip = await JSZip.loadAsync(bytes);
    zip.file('word/media/i.png', png, { base64: true });
    const { sections } = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
    const cell = (sections[0].blocks[0] as Table).rows[0].cells[0];
    expect(cell.blocks).toBeUndefined();
    expect(cell.content).toHaveLength(1);
    expect(cell.content[0]).toMatchObject({ type: 'image', width: 100, height: 50 });
  });

  it('keeps a list and a heading of a cell as list items and headings', async () => {
    const numbering = `<w:numbering ${NS}><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
    const { table } = await importTable(
      tbl(tr(tc(p('a', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') + p('b', '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>')))),
      { 'word/numbering.xml': numbering }
    );
    const blocks = table.rows[0].cells[0].blocks;
    expect(blocks).toMatchObject([
      { type: 'listItem', listType: 'bullet', level: 0 },
      { type: 'listItem', listType: 'bullet', level: 1 },
    ]);
  });

  it('keeps a page-break run of a cell paragraph out of the cell structure', async () => {
    const { table } = await importTable(tbl(tr(tc('<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + p('after')))));
    expect(table.rows[0].cells[0].blocks?.map((b) => b.type)).toEqual(['paragraph', 'paragraph']);
  });

  it('keeps a bookmark between paragraphs with the next paragraph', async () => {
    const { table } = await importTable(tbl(tr(tc('<w:bookmarkStart w:id="1" w:name="m"/>' + p('a') + p('b')))));
    expect((table.rows[0].cells[0].blocks?.[0] as { content: unknown }).content).toEqual([{ type: 'bookmark', name: 'm' }, text('a')]);
  });
});

describe('nested tables', () => {
  const nested = (depth: number): string => (depth === 0 ? tbl(tr(tc(p('deep')))) : tbl(tr(tc(p(`level ${depth}`) + nested(depth - 1) + p('')))));

  it('imports a table in a cell as a block of the cell, up to the corpus depth of 3 and beyond', async () => {
    const { table, warnings } = await importTable(nested(3));
    const level = (t: Table, n: number): Table => (n === 0 ? t : level(t.rows[0].cells[0].blocks!.find((b) => b.type === 'table') as Table, n - 1));
    expect(level(table, 3).rows[0].cells[0].content).toEqual([text('deep')]);
    expect(table.rows[0].cells[0].blocks?.map((b) => b.type)).toEqual(['paragraph', 'table', 'paragraph']);
    expect(warnings).toEqual([]);
  });

  it('flattens a table nested more than 20 deep with a warning rather than overflowing the stack', async () => {
    const { warnings } = await importTable(nested(25));
    expect(warnings).toContain('A table nested too deeply was not imported.');
  });

  it('does not give a nested table the conditional text formatting of the outer table style', async () => {
    const styles = `<w:styles ${NS}><w:style w:type="table" w:styleId="TS"><w:name w:val="TS"/><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style></w:styles>`;
    const { table } = await importTable(
      tbl(tr(tc(p('outer') + tbl(tr(tc(p('inner')))) + p(''))), '<w:tblStyle w:val="TS"/><w:tblLook w:val="0020" w:firstRow="1"/>'),
      { 'word/styles.xml': styles }
    );
    const cell = table.rows[0].cells[0];
    expect((cell.blocks?.[0] as { content: Array<{ bold?: boolean }> }).content[0].bold).toBe(true);
    const inner = cell.blocks?.[1] as Table;
    expect(inner.rows[0].cells[0].content[0]).not.toHaveProperty('bold');
  });

  it('numbers a list in a cell on its own', async () => {
    const numbering = `<w:numbering ${NS}><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
    const item = (t: string) => p(t, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
    const { sections } = await importDocx(await docx(item('x') + tbl(tr(tc(item('a') + item('b')))), { 'word/numbering.xml': numbering }));
    const cell = (sections[0].blocks[1] as Table).rows[0].cells[0];
    // the list in the cell begins at 1 (no `start`), and the list before the table is not continued by it
    expect(cell.blocks).toMatchObject([{ type: 'listItem', listType: 'numbered' }, { type: 'listItem', listType: 'numbered' }]);
  });
});

describe('merged cells', () => {
  it('extends a vertically merged cell under a continuation cell with the same grid span', async () => {
    const { table } = await importTable(
      tbl(
        tr(tc(p('a'), '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>') + tc(p('b'))) +
          tr(tc(p(''), '<w:gridSpan w:val="2"/><w:vMerge/>') + tc(p('c'))) +
          tr(tc(p('d')) + tc(p('e')) + tc(p('f'))),
        '',
        '<w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/>'
      )
    );
    expect(table.rows.map((r) => r.cells.length)).toEqual([2, 1, 3]);
    expect(table.rows[0].cells[0]).toMatchObject({ colspan: 2, rowspan: 2 });
  });

  it('counts a merged cell once when several narrower continuation cells sit under it', async () => {
    const { table } = await importTable(
      tbl(
        tr(tc(p('a'), '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>') + tc(p('b'))) +
          tr(tc(p(''), '<w:vMerge/>') + tc(p(''), '<w:vMerge/>') + tc(p('c')))
      )
    );
    expect(table.rows[0].cells[0].rowspan).toBe(2);
    expect(table.rows[1].cells).toHaveLength(1);
  });

  it('follows a merge across rows whose cells have different spans', async () => {
    const { table } = await importTable(
      tbl(
        tr(tc(p('a')) + tc(p('b'), '<w:vMerge w:val="restart"/>') + tc(p('c'))) +
          tr(tc(p('d'), '<w:gridSpan w:val="2"/>') + tc(p('e'))) +
          tr(tc(p('f')) + tc(p(''), '<w:vMerge/>') + tc(p('g')))
      )
    );
    // the merge of the middle column was ended by the row that did not continue it, so the last row's continuation is an ordinary cell
    expect(table.rows[0].cells[1].rowspan).toBeUndefined();
    expect(table.rows[2].cells).toHaveLength(3);
  });

  it('keeps a continuation cell with no merge above it as an ordinary cell', async () => {
    const { table } = await importTable(tbl(tr(tc(p('keep me'), '<w:vMerge/>'))));
    expect(table.rows[0].cells[0].content).toEqual([text('keep me')]);
  });

  it('keeps cells where they are when a row skips grid columns (`gridBefore`)', async () => {
    const { table } = await importTable(
      tbl(tr(tc(p('a')) + tc(p('b')) + tc(p('c'))) + tr(tc(p('x')), '<w:gridBefore w:val="2"/>'), '', '<w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/>')
    );
    const second = table.rows[1].cells;
    expect(second).toHaveLength(2);
    expect(second[0]).toMatchObject({ content: [], colspan: 2, colwidthPx: [66.67, 66.67] });
    expect(second[0].borders).toEqual({ top: { style: 'none' }, right: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' } });
    expect(second[1].content).toEqual([text('x')]);
  });
});

describe('text direction', () => {
  it('records a rotated cell', async () => {
    const { table } = await importTable(tbl(tr(tc(p('up'), '<w:textDirection w:val="btLr"/>') + tc(p('down'), '<w:textDirection w:val="tbRl"/>') + tc(p('flat'), '<w:textDirection w:val="lrTb"/>'))));
    expect(table.rows[0].cells.map((c) => c.textDirection)).toEqual(['btLr', 'tbRl', undefined]);
  });
});

describe('export and round trip', () => {
  const section = (blocks: DocBlock[]): DocSection => ({
    pageWidthPx: 816,
    pageHeightPx: 1056,
    marginTopPx: 96,
    marginRightPx: 96,
    marginBottomPx: 96,
    marginLeftPx: 96,
    orientation: 'portrait',
    blocks,
  });
  const para = (s: string): DocBlock => ({ type: 'paragraph', content: [{ type: 'text', text: s }] });
  const roundTrip = async (blocks: DocBlock[]) => {
    const { bytes } = await exportDocx({ title: 'T', sections: [section(blocks)] });
    return { bytes, imported: await importDocx(bytes) };
  };
  const tableOf = (cell: Cell): DocBlock => ({ type: 'table', rows: [{ cells: [cell] }] });

  it('writes the paragraphs, a nested table and the text direction of a cell and reads them back', async () => {
    const inner = tableOf({ content: [{ type: 'text', text: 'inner' }] });
    const { imported } = await roundTrip([
      tableOf({
        content: [{ type: 'text', text: 'one' }],
        blocks: [para('one'), para('two'), inner, para('')],
        textDirection: 'btLr',
      }),
    ]);
    const cell = (imported.sections[0].blocks[0] as Table).rows[0].cells[0];
    expect(cell.textDirection).toBe('btLr');
    expect(cell.blocks?.map((b) => b.type)).toEqual(['paragraph', 'paragraph', 'table', 'paragraph']);
    expect(((cell.blocks?.[2] as Table).rows[0].cells[0].content[0] as { text: string }).text).toBe('inner');
  });

  it('ends a cell with a paragraph after a nested table (the file format requires it)', async () => {
    const inner = tableOf({ content: [{ type: 'text', text: 'inner' }] });
    const { bytes, imported } = await roundTrip([tableOf({ content: [], blocks: [para('a'), inner] })]);
    const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
    expect(xml).toMatch(/<\/w:tbl><w:p\/><\/w:tc>|<\/w:tbl><w:p><\/w:p><\/w:tc>/);
    expect((imported.sections[0].blocks[0] as Table).rows[0].cells[0].blocks?.map((b) => b.type)).toEqual(['paragraph', 'table', 'paragraph']);
  });

  it('numbers a list in a cell from 1, and neither continues the list before the table nor is continued by it', async () => {
    const item = (s: string, start?: number): DocBlock => ({ type: 'listItem', listType: 'numbered', level: 0, content: [{ type: 'text', text: s }], ...(start !== undefined && { start }) });
    const { imported } = await roundTrip([item('x', 5), tableOf({ content: [], blocks: [item('a'), item('b')] }), item('y')]);
    const [before, t, after] = imported.sections[0].blocks;
    expect(before).toMatchObject({ type: 'listItem', start: 5 });
    expect((t as Table).rows[0].cells[0].blocks).toMatchObject([{ type: 'listItem' }, { type: 'listItem' }]);
    // a list that starts at 1 after the table carries no `start` (it is not a continuation of the 5, 6... before it)
    expect(after).toMatchObject({ type: 'listItem', listType: 'numbered' });
    expect(after).not.toHaveProperty('start');
  });

  it('writes the comment anchors of every paragraph in a cell', async () => {
    const run = (s: string, id: string): DocBlock => ({ type: 'paragraph', content: [{ type: 'text', text: s, commentThreadId: id }] });
    const { bytes } = await exportDocx({
      title: 'T',
      sections: [section([tableOf({ content: [], blocks: [run('first', 't1'), run('second', 't2')] })])],
      comments: [
        { id: 't1', quote: 'first', comments: [{ id: 'c1', author: 'A', text: 'one' }] },
        { id: 't2', quote: 'second', comments: [{ id: 'c2', author: 'A', text: 'two' }] },
      ],
    });
    const { sections } = await importDocx(bytes);
    const blocks = (sections[0].blocks[0] as Table).rows[0].cells[0].blocks as Array<{ content: Array<{ commentThreadId?: string }> }>;
    expect(blocks.map((b) => !!b.content[0].commentThreadId)).toEqual([true, true]);
  });

  it('does not report nested tables or multi-paragraph cells as lost', async () => {
    const { bytes } = await roundTrip([tableOf({ content: [], blocks: [para('a'), tableOf({ content: [] }), para('')] })]);
    expect((await detectDocxLoss(bytes, { detail: 'full' })).items.map((i) => i.id)).not.toContain('docx.tableStructure');
  });
});
