import { Schema, Node as PmNode } from 'prosemirror-model';
import { createBlockPositionMap, noteLabel, proseMirrorToFlowBlocks, proseMirrorToSectionedFlowBlocks } from './pm-to-blocks';
import type { ParagraphBlock, LinkRun } from './flow-blocks';

// Minimal schema: enough to build a paragraph containing a linked text run.
const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { group: 'block', content: 'inline*', toDOM: () => ['p', 0] },
    text: { group: 'inline' },
  },
  marks: {
    link: {
      attrs: { href: {} },
      toDOM: (m) => ['a', { href: m.attrs.href }, 0],
    },
  },
});

function docWithLink(href: string): PmNode {
  const mark = schema.marks.link.create({ href });
  const textNode = schema.text('click me', [mark]);
  const para = schema.nodes.paragraph.create(null, textNode);
  return schema.nodes.doc.create(null, para);
}

describe('proseMirrorToFlowBlocks — link marks', () => {
  it('maps a link mark to a LinkRun carrying the href (not a plain text run)', () => {
    const blocks = proseMirrorToFlowBlocks(docWithLink('https://example.com'));
    const para = blocks[0] as ParagraphBlock;
    expect(para.kind).toBe('paragraph');
    const run = para.runs[0] as LinkRun;
    expect(run.kind).toBe('link');
    expect(run.href).toBe('https://example.com');
    expect(run.text).toBe('click me');
    expect(run.underline).toBe(true);
  });

  it('neutralizes a javascript: href when building the model', () => {
    const blocks = proseMirrorToFlowBlocks(docWithLink('javascript:alert(1)'));
    const run = (blocks[0] as ParagraphBlock).runs[0] as LinkRun;
    expect(run.kind).toBe('link');
    expect(run.href).toBe('');
  });
});

// The real docs schema: paragraph formatting and run flags travel through ProseMirror attrs and marks.
import { docsSchema } from '@b32nio/spindle-docs-core';
import type { ListItemBlock, TextRun } from './flow-blocks';

const j = (json: unknown) => docsSchema.nodeFromJSON({ type: 'doc', content: json });
const text = (s: string, marks?: Array<{ type: string; attrs?: object }>) => (marks ? { type: 'text', text: s, marks } : { type: 'text', text: s });

describe('proseMirrorToFlowBlocks: DOCX paragraph and run properties', () => {
  it('maps indents, spacing and a custom line spacing of a paragraph', () => {
    const doc = j([
      { type: 'paragraph', attrs: { leftIndent: 36, rightIndent: 12, firstLineIndent: -18, spaceAfter: 0, lineSpacing: { type: 'multiple', value: 1.2 } }, content: [text('x')] },
    ]);
    const para = proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock;
    expect(para.attrs).toMatchObject({ leftIndent: 36, rightIndent: 12, firstLineIndent: -18, spaceAfter: 0 });
    expect(para.attrs?.lineHeight).toBeCloseTo(1.8); // 1.2 x the layout's single-spacing default of 1.5
  });

  it('scales the keyword spacings and ignores heights it cannot express', () => {
    const lineHeight = (lineSpacing: unknown) =>
      (proseMirrorToFlowBlocks(j([{ type: 'paragraph', attrs: { lineSpacing }, content: [text('x')] }]))[0] as ParagraphBlock).attrs?.lineHeight;
    expect(lineHeight('single')).toBeUndefined();
    expect(lineHeight({ type: 'onePointFive' })).toBe(2.25);
    expect(lineHeight({ type: 'double' })).toBe(3);
    expect(lineHeight({ type: 'exactly', value: 20 })).toBeUndefined();
    expect(lineHeight({ type: 'atLeast', value: 20 })).toBeUndefined();
    expect(lineHeight({ type: 'multiple', value: 0 })).toBeUndefined();
    expect(lineHeight({ type: 'multiple', value: 99 })).toBeUndefined();
  });

  it('carries heading spacing and indents', () => {
    const doc = j([{ type: 'heading', attrs: { level: 2, spaceBefore: 12, spaceAfter: 0, leftIndent: 8 }, content: [text('H')] }]);
    expect(proseMirrorToFlowBlocks(doc)[0]).toMatchObject({ kind: 'heading', attrs: { spaceBefore: 12, spaceAfter: 0, leftIndent: 8 } });
  });

  it('maps small caps, all caps and hidden marks onto the run', () => {
    const doc = j([
      { type: 'paragraph', content: [text('a', [{ type: 'smallCaps' }]), text('b', [{ type: 'allCaps' }]), text('c', [{ type: 'hidden' }]), text('d')] },
    ]);
    const runs = (proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock).runs as TextRun[];
    expect(runs.map((r) => [r.smallCaps, r.allCaps, r.hidden])).toEqual([
      [true, undefined, undefined],
      [undefined, true, undefined],
      [undefined, undefined, true],
      [undefined, undefined, undefined],
    ]);
  });

  it('ignores the named style mark and attr (the formatting is already on the run and block)', () => {
    const doc = j([
      { type: 'paragraph', attrs: { namedStyleId: 'Callout' }, content: [text('a', [{ type: 'namedStyle', attrs: { id: 'AccentWord' } }, { type: 'bold' }])] },
    ]);
    const para = proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock;
    const run = para.runs[0] as TextRun;
    expect(run.text).toBe('a');
    expect(run.bold).toBe(true);
    expect(JSON.stringify(para)).not.toContain('AccentWord');
  });

  it('numbers an ordered list from its start attribute', () => {
    const item = (s: string) => ({ type: 'list_item', attrs: { level: 0 }, content: [{ type: 'paragraph', content: [text(s)] }] });
    const doc = j([{ type: 'ordered_list', attrs: { start: 7 }, content: [item('a'), item('b')] }, { type: 'ordered_list', content: [item('c')] }]);
    const items = proseMirrorToFlowBlocks(doc) as ListItemBlock[];
    expect(items.map((i) => i.listIndex)).toEqual([7, 8, 1]);
  });
});

describe('field nodes', () => {
  const fieldSchema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: { group: 'block', content: 'inline*', toDOM: () => ['p', 0] },
      text: { group: 'inline' },
      field: { inline: true, group: 'inline', atom: true, attrs: { instruction: {}, kind: {}, result: {}, dirty: { default: false } } },
    },
  });

  it('displays the cached result of a field, from a string or from runs', () => {
    const doc = fieldSchema.nodeFromJSON({
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          content: [
            { type: 'text', text: 'see ' },
            { type: 'field', attrs: { instruction: 'REF x', kind: 'REF', result: 'Figure 1' } },
            { type: 'field', attrs: { instruction: 'SEQ a', kind: 'SEQ', result: [{ type: 'text', text: '2', bold: true }, { type: 'text', text: 'b' }] } },
          ],
        },
      ],
    });
    const para = proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock;
    expect(para.runs.map((r) => (r.kind === 'text' ? r.text : ''))).toEqual(['see ', 'Figure 1', '2b']);
  });
});

describe('bookmark nodes', () => {
  const bookmarkSchema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: { group: 'block', content: 'inline*', toDOM: () => ['p', 0] },
      text: { group: 'inline' },
      bookmark: { inline: true, group: 'inline', atom: true, attrs: { name: {} } },
    },
    marks: { link: { attrs: { href: {} }, toDOM: (m) => ['a', { href: m.attrs.href }, 0] } },
  });

  it('keeps a bookmark as one zero-width character, so layout offsets stay in step with positions', () => {
    const doc = bookmarkSchema.nodeFromJSON({
      type: 'doc',
      content: [{ type: 'paragraph', content: [{ type: 'bookmark', attrs: { name: '_Toc1' } }, { type: 'text', text: 'Intro' }] }],
    });
    const para = proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock;
    expect(para.runs.map((r) => (r.kind === 'text' ? r.text : ''))).toEqual(['​', 'Intro']);
  });

  it('keeps an internal link href as written', () => {
    const doc = bookmarkSchema.nodeFromJSON({
      type: 'doc',
      content: [{ type: 'paragraph', content: [{ type: 'text', text: 'go', marks: [{ type: 'link', attrs: { href: '#_Toc1' } }] }] }],
    });
    const para = proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock;
    expect(para.runs[0]).toMatchObject({ kind: 'link', href: '#_Toc1' });
  });
});

describe('section_break nodes in the flow', () => {
  const doc = j([
    { type: 'paragraph', content: [text('one')] },
    { type: 'section_break', attrs: { sectionId: 'a' } },
    { type: 'paragraph', content: [text('two')] },
    { type: 'table', content: [{ type: 'table_row', content: [{ type: 'table_cell', content: [{ type: 'paragraph' }] }] }] },
    { type: 'section_break', attrs: { sectionId: 'b' } },
    { type: 'paragraph', content: [text('three')] },
  ]);

  it('draws nothing for a break and keeps the blocks in order', () => {
    const blocks = proseMirrorToFlowBlocks(doc, { resetIds: true });
    expect(blocks.map((b) => b.kind)).toEqual(['paragraph', 'paragraph', 'table', 'paragraph']);
  });

  it('maps each block to its own position, skipping the breaks', () => {
    const blocks = proseMirrorToFlowBlocks(doc, { resetIds: true });
    const map = createBlockPositionMap(doc, blocks);
    const starts = blocks.map((b) => map.get(b.id)?.start);
    const expected: number[] = [];
    doc.forEach((n, offset) => {
      if (n.type.name !== 'section_break') expected.push(offset);
    });
    expect(starts).toEqual(expected);
  });
});

describe('proseMirrorToSectionedFlowBlocks', () => {
  const para = (t: string) => ({ type: 'paragraph', content: [text(t)] });
  const list = { type: 'bullet_list', content: [{ type: 'list_item', content: [para('a')] }, { type: 'list_item', content: [para('b')] }] };

  it('is one section holding every block when the document has no break', () => {
    const doc = j([para('x'), list]);
    const out = proseMirrorToSectionedFlowBlocks(doc, { resetIds: true });
    expect(out.blocks).toEqual(proseMirrorToFlowBlocks(doc, { resetIds: true }));
    expect(out.sections).toEqual([{ sectionId: null, blockStart: 0, blockEnd: out.blocks.length }]);
  });

  it('splits at breaks, counting the flow blocks a list expands to', () => {
    const doc = j([para('one'), list, { type: 'section_break', attrs: { sectionId: 'a' } }, para('two'), { type: 'section_break', attrs: { sectionId: 'b' } }, para('three')]);
    const out = proseMirrorToSectionedFlowBlocks(doc);
    expect(out.blocks.map((b) => b.kind)).toEqual(['paragraph', 'listItem', 'listItem', 'paragraph', 'paragraph']);
    expect(out.sections).toEqual([
      { sectionId: 'a', blockStart: 0, blockEnd: 3 },
      { sectionId: 'b', blockStart: 3, blockEnd: 4 },
      { sectionId: null, blockStart: 4, blockEnd: 5 },
    ]);
  });

  it('keeps empty sections: adjacent breaks and a break at the end', () => {
    const doc = j([{ type: 'section_break', attrs: { sectionId: 'a' } }, { type: 'section_break', attrs: { sectionId: 'b' } }, para('x'), { type: 'section_break', attrs: { sectionId: 'c' } }]);
    expect(proseMirrorToSectionedFlowBlocks(doc).sections).toEqual([
      { sectionId: 'a', blockStart: 0, blockEnd: 0 },
      { sectionId: 'b', blockStart: 0, blockEnd: 0 },
      { sectionId: 'c', blockStart: 0, blockEnd: 1 },
      { sectionId: null, blockStart: 1, blockEnd: 1 },
    ]);
  });
});

describe('content controls', () => {
  const para = (t: string) => ({ type: 'paragraph', content: [text(t)] });
  const control = (attrs: object, content: unknown[]) => ({ type: 'content_control', attrs: { controlType: 'richText', ...attrs }, content });
  const inline = (attrs: object, content: unknown[]) => ({ type: 'content_control_inline', attrs: { controlType: 'text', ...attrs }, content });

  it('flows the blocks of a control as if it were not there, and titles the first one with the control', () => {
    const doc = j([para('before'), control({ alias: 'Summary', tag: 's' }, [para('in one'), para('in two')]), para('after')]);
    const blocks = proseMirrorToFlowBlocks(doc, { resetIds: true });
    expect(blocks.map((b) => b.kind)).toEqual(['paragraph', 'paragraph', 'paragraph', 'paragraph']);
    expect(blocks.map((b) => b.controlChips)).toEqual([undefined, ['Summary'], undefined, undefined]);
  });

  it('titles a control by alias, then tag, then kind; nested controls give outermost-first chips', () => {
    const titleOf = (attrs: object) => proseMirrorToFlowBlocks(j([control(attrs, [para('x')])]))[0].controlChips;
    expect(titleOf({ alias: 'A', tag: 'T' })).toEqual(['A']);
    expect(titleOf({ tag: 'T' })).toEqual(['T']);
    expect(titleOf({ controlType: 'dropDown' })).toEqual(['Drop-down list']);
    expect(titleOf({ controlType: 'weird' })).toEqual(['Content control']);
    const nested = j([control({ alias: 'Outer' }, [control({ alias: 'Inner' }, [para('deep')])])]);
    expect(proseMirrorToFlowBlocks(nested)[0].controlChips).toEqual(['Outer', 'Inner']);
  });

  it('puts a control around a list, a table and an empty one without blocks of its own', () => {
    const item = { type: 'list_item', attrs: { level: 0 }, content: [para('li')] };
    const table = { type: 'table', content: [{ type: 'table_row', content: [{ type: 'table_cell', content: [para('c')] }] }] };
    const blocks = proseMirrorToFlowBlocks(j([control({ alias: 'L' }, [{ type: 'bullet_list', content: [item] }, table])]));
    expect(blocks.map((b) => b.kind)).toEqual(['listItem', 'table']);
    expect(blocks[0].controlChips).toEqual(['L']);
    expect(blocks[1].controlChips).toBeUndefined();
  });

  it('draws the opening and closing positions of an inline control as zero-width runs, so offsets stay in step with positions', () => {
    const doc = j([{ type: 'paragraph', content: [text('a '), inline({ alias: 'Name' }, [text('Ada')]), text('!')] }]);
    const runs = (proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock).runs as TextRun[];
    expect(runs.map((r) => r.text)).toEqual(['a ', '​', 'Ada', '​', '!']);
    expect(runs.map((r) => r.control?.title)).toEqual([undefined, 'Name', 'Name', 'Name', undefined]);
    // The paragraph holds 2 + 1 + 3 + 1 + 1 positions and the runs 2 + 1 + 3 + 1 + 1 characters.
    const paragraph = doc.firstChild!;
    expect(paragraph.content.size).toBe(runs.reduce((n, r) => n + r.text.length, 0));
  });

  it('tags placeholder text and a check box, and the innermost control wins', () => {
    const doc = j([
      {
        type: 'paragraph',
        content: [
          inline({ alias: 'Outer', showingPlaceholder: true }, [text('click'), inline({ controlType: 'checkbox', alias: 'Box' }, [text('☐')])]),
        ],
      },
    ]);
    const runs = (proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock).runs as TextRun[];
    expect(runs.find((r) => r.text === 'click')?.control).toEqual({ title: 'Outer', placeholder: true });
    expect(runs.find((r) => r.text === '☐')?.control).toEqual({ title: 'Box', checkbox: true });
  });

  it('keeps links, fields and bookmarks inside an inline control in step with positions', () => {
    const doc = j([
      {
        type: 'paragraph',
        content: [
          inline({}, [
            { type: 'bookmark', attrs: { name: 'b' } },
            { type: 'field', attrs: { instruction: 'REF x', kind: 'REF', result: 'Fig' } },
            text('go', [{ type: 'link', attrs: { href: 'https://example.com' } }]),
          ]),
        ],
      },
    ]);
    const runs = (proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock).runs;
    expect(runs.map((r) => (r.kind === 'text' || r.kind === 'link' ? r.text : ''))).toEqual(['​', '​', 'Fig', 'go', '​']);
    expect(runs.map((r) => (r.kind === 'text' || r.kind === 'link' ? r.control?.title : undefined))).toEqual(Array(5).fill('Text'));
    expect(doc.firstChild!.content.size).toBe(1 + 1 + 1 + 'go'.length + 1);
  });

  it('maps blocks inside controls to their own positions, at the top level and in table cells', () => {
    const cell = (children: unknown[]) => ({ type: 'table', content: [{ type: 'table_row', content: [{ type: 'table_cell', content: children }] }] });
    const doc = j([para('one'), control({}, [para('two'), control({}, [para('three')])]), cell([para('c1'), control({}, [para('c2')]), para('c3')]), para('four')]);
    const blocks = proseMirrorToFlowBlocks(doc, { resetIds: true });
    const map = createBlockPositionMap(doc, blocks);
    const paragraphs = blocks.filter((b): b is ParagraphBlock => b.kind === 'paragraph');
    // Each paragraph's mapped range holds its own text.
    const textAt = (start: number, end: number) => doc.textBetween(start, end, ' ');
    expect(paragraphs.map((p) => textAt(map.get(p.id)!.start, map.get(p.id)!.end))).toEqual(['one', 'two', 'three', 'four']);
    const table = blocks.find((b) => b.kind === 'table');
    if (table?.kind !== 'table') throw new Error('table');
    const cellBlocks = table.rows[0].cells[0].blocks as ParagraphBlock[];
    expect(cellBlocks.map((b) => doc.textBetween(b.pmStart!, b.pmEnd!, ' '))).toEqual(['c1', 'c2', 'c3']);
  });

  it('splits sections at breaks after flattening controls', () => {
    const doc = j([control({}, [para('a'), para('b')]), { type: 'section_break', attrs: { sectionId: 'x' } }, para('c')]);
    const out = proseMirrorToSectionedFlowBlocks(doc);
    expect(out.sections).toEqual([
      { sectionId: 'x', blockStart: 0, blockEnd: 2 },
      { sectionId: null, blockStart: 2, blockEnd: 3 },
    ]);
  });
});

describe('notes', () => {
  const para = (...content: unknown[]) => ({ type: 'paragraph', content });
  const note = (kind: 'footnote' | 'endnote', body: unknown[], attrs: object = {}) => ({ type: 'note_ref', attrs: { kind, noteId: '1', ...attrs }, content: body });
  const labelOf = (run: unknown) => (run as TextRun).text;

  it('numbers footnotes and endnotes separately, in document order, across paragraphs, tables and lists', () => {
    const cell = { type: 'table', content: [{ type: 'table_row', content: [{ type: 'table_cell', content: [para(note('footnote', [text('c')]))] }] }] };
    const item = { type: 'list_item', attrs: { level: 0 }, content: [para(note('endnote', [text('l')]))] };
    const doc = j([para(note('footnote', [text('a')]), note('endnote', [text('b')])), cell, { type: 'bullet_list', content: [item] }, para(note('footnote', [text('d')]))]);
    const blocks = proseMirrorToFlowBlocks(doc, { resetIds: true });
    const refs = (blocks.filter((b) => b.kind === 'paragraph' && !b.synthetic)[0] as ParagraphBlock).runs.filter((r): r is TextRun => r.kind === 'text' && !!r.noteRef);
    expect(refs.map((r) => [r.noteRef?.kind, r.noteRef?.ordinal, r.text])).toEqual([
      ['footnote', 1, '1'],
      ['endnote', 1, 'i'],
    ]);
    const entries = blocks.filter((b): b is ParagraphBlock => b.kind === 'paragraph' && !!b.noteTarget);
    expect(entries.map((b) => `${b.noteTarget!.kind}:${b.noteTarget!.ordinal}`)).toEqual(['footnote:1', 'footnote:2', 'footnote:3', 'endnote:1', 'endnote:2']);
    expect(entries.map((b) => labelOf(b.runs[0]))).toEqual(['1', '2', '3', 'i', 'ii']);
  });

  it('draws a reference as a superscript number padded to the positions it takes, so offsets stay in step', () => {
    const body = [text('abcdef')];
    const doc = j([para(text('x'), note('footnote', body), text('y'))]);
    const first = proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock;
    const runs = first.runs as TextRun[];
    expect(runs.map((r) => r.text)).toEqual(['x', '1', '​'.repeat(7), 'y']);
    expect(runs[1]).toMatchObject({ superscript: true, noteRef: { kind: 'footnote', ordinal: 1 } });
    // The paragraph holds 1 + (2 + 6) + 1 positions.
    expect(runs.reduce((n, r) => n + r.text.length, 0)).toBe(doc.firstChild!.content.size);
  });

  it('shows a custom mark instead of the number, and does not pad when the mark is as long as the positions', () => {
    const doc = j([para(note('footnote', [], { customMark: '*' }))]);
    const runs = (proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock).runs as TextRun[];
    expect(runs.map((r) => r.text)).toEqual(['*', '​']);
    const long = j([para(note('footnote', [], { customMark: '***' }))]);
    expect((proseMirrorToFlowBlocks(long)[0] as ParagraphBlock).runs.map(labelOf)).toEqual(['***']);
  });

  it('collects the notes in an area after the last block: a heading per kind, then each note, numbered, in a smaller size', () => {
    const doc = j([para(text('a'), note('footnote', [text('first '), text('bold', [{ type: 'bold' }])])), para(note('endnote', [text('last')]))]);
    const blocks = proseMirrorToFlowBlocks(doc);
    expect(blocks.map((b) => (b as ParagraphBlock).synthetic ?? false)).toEqual([false, false, true, true, true, true]);
    const area = blocks.slice(2) as ParagraphBlock[];
    expect(area.map((b) => b.runs.map(labelOf).join(''))).toEqual(['Footnotes', '1 first bold', 'Endnotes', 'i last']);
    expect(area[1].runs.slice(2).map((r) => (r as TextRun).fontSize)).toEqual([10, 10]);
    expect(area[1].noteTarget).toEqual({ kind: 'footnote', ordinal: 1 });
    expect(area[0].noteTarget).toBeUndefined();
  });

  it('has no area when there are no notes, and lists no note of another section', () => {
    expect(proseMirrorToFlowBlocks(j([para(text('x'))])).some((b) => (b as ParagraphBlock).synthetic)).toBe(false);
    const doc = j([para(note('footnote', [text('one')])), { type: 'section_break', attrs: { sectionId: 'a' } }, para(text('two')), { type: 'section_break', attrs: { sectionId: 'b' } }, para(note('footnote', [text('three')]))]);
    const out = proseMirrorToSectionedFlowBlocks(doc);
    expect(out.sections).toEqual([
      { sectionId: 'a', blockStart: 0, blockEnd: 3 },
      { sectionId: 'b', blockStart: 3, blockEnd: 4 },
      { sectionId: null, blockStart: 4, blockEnd: 7 },
    ]);
    const area = (range: { blockStart: number; blockEnd: number }) => out.blocks.slice(range.blockStart, range.blockEnd).map((b) => (b as ParagraphBlock).runs.map(labelOf).join('').replace(/\u200B/g, ''));
    expect(area(out.sections[0])).toEqual(['1', 'Footnotes', '1 one']);
    expect(area(out.sections[2])).toEqual(['2', 'Footnotes', '2 three']);
  });

  it('maps the blocks of the document to their positions, skipping the notes area', () => {
    const doc = j([para(note('footnote', [text('n')])), { type: 'section_break', attrs: { sectionId: 'a' } }, para(text('after')), para(text('last'))]);
    const { blocks } = proseMirrorToSectionedFlowBlocks(doc);
    const map = createBlockPositionMap(doc, blocks);
    const real = blocks.filter((b) => !(b as ParagraphBlock).synthetic);
    expect(real.map((b) => doc.textBetween(map.get(b.id)!.start, map.get(b.id)!.end, ' ')).join('|')).toBe('n|after|last');
    expect(blocks.filter((b) => (b as ParagraphBlock).synthetic).every((b) => !map.has(b.id))).toBe(true);
  });

  it('writes endnote numbers in roman numerals', () => {
    expect([1, 4, 9, 14, 40, 90, 400, 1994, 3999].map((ordinal) => noteLabel({ kind: 'endnote', ordinal }))).toEqual(['i', 'iv', 'ix', 'xiv', 'xl', 'xc', 'cd', 'mcmxciv', 'mmmcmxcix']);
    expect(noteLabel({ kind: 'footnote', ordinal: 12 })).toBe('12');
  });
});
