// Multi-section collaboration: section_break nodes carry the section structure in the Y.Doc.

import { readFileSync } from 'node:fs';
import path from 'node:path';
import * as Y from 'yjs';
import { EditorState } from 'prosemirror-state';
import { updateYFragment, yXmlFragmentToProseMirrorRootNode } from 'y-prosemirror';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { docsSchema } from '../prosemirror/schema';
import { blocksToPmDoc } from '../prosemirror/sync';
import { DocumentImpl } from '../document';
import { importFromDocx } from '../export/docx';
import { createParagraphFromText } from '../blocks/paragraph';
import { createListItemFromText } from '../blocks/list';
import { createTableFromData } from '../blocks/table';
import { DEFAULT_PAGE_CONFIG } from '../types';
import type { Block, DocumentData, PageConfig, ParagraphBlock, SectionData } from '../types';
import {
  COLLAB_SCHEMA_VERSION,
  CollabSchemaVersionError,
  assertSchemaVersion,
  getYDocFields,
  hydrateYDocFromData,
  readSchemaVersion,
  serializeYDocToData,
} from './y-schema';
import { CollabBlockedError, getCollabBlockers } from './blockers';

const identity = (n: string): CollabIdentity => ({ userId: n, displayName: n, color: '#111' });
const fixtures = path.resolve(__dirname, '../../../../fixtures/docx');

const landscape: PageConfig = {
  size: { w: 816, h: 1056 },
  margins: { top: 48, right: 48, bottom: 48, left: 48 },
  orientation: 'landscape',
  columns: { count: 2, spacing: 24 },
};
const a4: PageConfig = {
  size: { w: 794, h: 1123 },
  margins: { top: 96, right: 72, bottom: 96, left: 72 },
  orientation: 'portrait',
};

function para(text: string): ParagraphBlock {
  return createParagraphFromText(text);
}

function fourSections(): DocumentData {
  const sections: SectionData[] = [
    {
      id: 's1',
      pageConfig: DEFAULT_PAGE_CONFIG,
      blocks: [para('one a'), para('one b')],
      header: { blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'H1' }] }], differentFirstPage: true, firstPageBlocks: [] },
    },
    {
      id: 's2',
      pageConfig: landscape,
      blocks: [para('two'), createListItemFromText('bullet', 'item'), createTableFromData([['a', 'b']])],
      footer: { blocks: [{ type: 'paragraph', content: [{ type: 'text', text: 'F2' }] }] },
      pageNumbering: { start: 5, format: 'lowerRoman' },
    },
    { id: 's3', pageConfig: a4, blocks: [], verticalAlign: 'center' },
    { id: 's4', pageConfig: DEFAULT_PAGE_CONFIG, blocks: [para('four')] },
  ];
  return {
    id: 'doc',
    title: 'Four',
    sections,
    defaultPageConfig: DEFAULT_PAGE_CONFIG,
    textStylePool: {},
    paragraphStylePool: {},
  };
}

function text(blocks: Block[]): string[] {
  return blocks.map((b) => ('content' in b ? b.content.map((c) => ('text' in c ? c.text : '')).join('') : b.type));
}

function breakIds(ydoc: Y.Doc): string[] {
  const pm = yXmlFragmentToProseMirrorRootNode(getYDocFields(ydoc).content, docsSchema);
  const ids: string[] = [];
  pm.forEach((n) => {
    if (n.type.name === 'section_break') ids.push(n.attrs.sectionId as string);
  });
  return ids;
}

/** Edit the body of a Y.Doc the way ySyncPlugin does: read it as PM, transact, diff back. */
function edit(ydoc: Y.Doc, fn: (state: EditorState) => ReturnType<EditorState['tr']['replaceWith']>): void {
  const fragment = getYDocFields(ydoc).content;
  const state = EditorState.create({ doc: yXmlFragmentToProseMirrorRootNode(fragment, docsSchema) });
  const tr = fn(state);
  updateYFragment(ydoc, fragment, tr.doc, { mapping: new Map(), isOMark: new Map() });
}

function syncDocs(a: Y.Doc, b: Y.Doc): void {
  const ua = Y.encodeStateAsUpdate(a, Y.encodeStateVector(b));
  const ub = Y.encodeStateAsUpdate(b, Y.encodeStateVector(a));
  Y.applyUpdate(b, ua);
  Y.applyUpdate(a, ub);
}

afterEach(() => __resetInMemoryRooms());

describe('hydrate and serialize a four-section document', () => {
  it('writes one section_break per section but the last, and the schema version', () => {
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, fourSections());
    expect(breakIds(ydoc)).toEqual(['s1', 's2', 's3']);
    expect(getYDocFields(ydoc).sectionsMeta.length).toBe(4);
    expect(getYDocFields(ydoc).meta.get('schemaVersion')).toBe(COLLAB_SCHEMA_VERSION);
  });

  it('reads back every section with its page setup, header, footer, numbering and blocks', () => {
    const ydoc = new Y.Doc();
    const data = fourSections();
    hydrateYDocFromData(ydoc, data);
    const back = serializeYDocToData(ydoc);

    expect(back.sections.map((s) => s.id)).toEqual(['s1', 's2', 's3', 's4']);
    expect(back.sections.map((s) => s.pageConfig)).toEqual(data.sections.map((s) => s.pageConfig));
    expect(back.sections[0].header).toEqual(data.sections[0].header);
    expect(back.sections[1].footer).toEqual(data.sections[1].footer);
    expect(back.sections[1].pageNumbering).toEqual({ start: 5, format: 'lowerRoman' });
    expect(back.sections[2].verticalAlign).toBe('center');
    expect(back.sections.map((s) => text(s.blocks))).toEqual([
      ['one a', 'one b'],
      ['two', 'item', 'table'],
      [], // an empty section stays empty
      ['four'],
    ]);
    expect(back.sections[1].blocks.map((b) => b.type)).toEqual(['paragraph', 'list-item', 'table']);
  });

  it('keeps a section that is empty at the very start and one that is empty at the end', () => {
    const data = fourSections();
    data.sections = [
      { id: 'a', pageConfig: DEFAULT_PAGE_CONFIG, blocks: [] },
      { id: 'b', pageConfig: a4, blocks: [para('mid')] },
      { id: 'c', pageConfig: landscape, blocks: [] },
    ];
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, data);
    const back = serializeYDocToData(ydoc);
    expect(back.sections.map((s) => [s.id, text(s.blocks)])).toEqual([['a', []], ['b', ['mid']], ['c', []]]);
    expect(back.sections[2].pageConfig).toEqual(landscape);
  });

  it('reads the last section by the id no break names, so a deleted break leaves the later section standing', () => {
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, fourSections());
    // Delete the break between s2 and s3 (the third top-level node counted from the s2 break).
    edit(ydoc, (state) => {
      let pos = -1;
      let seen = 0;
      state.doc.forEach((n, offset) => {
        if (n.type.name === 'section_break' && ++seen === 2) pos = offset;
      });
      return state.tr.delete(pos, pos + 1);
    });
    const back = serializeYDocToData(ydoc);
    // s2 merged into s3: the surviving section is s3 (the one after the break) with its own page setup.
    expect(back.sections.map((s) => s.id)).toEqual(['s1', 's3', 's4']);
    expect(back.sections[1].pageConfig).toEqual(a4);
    expect(text(back.sections[1].blocks)).toEqual(['two', 'item', 'table']);
  });

  it('reads a body with no break as the last section when others were left behind', () => {
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, fourSections());
    edit(ydoc, (state) => {
      const tr = state.tr;
      const positions: number[] = [];
      state.doc.forEach((n, offset) => {
        if (n.type.name === 'section_break') positions.push(offset);
      });
      for (const p of positions.reverse()) tr.delete(p, p + 1);
      return tr;
    });
    const back = serializeYDocToData(ydoc);
    expect(back.sections.map((s) => s.id)).toEqual(['s4']);
    expect(text(back.sections[0].blocks)).toContain('four');
  });

  it('renames a repeated break id instead of producing two sections with one id', () => {
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, fourSections());
    edit(ydoc, (state) => {
      const brk = state.doc.firstChild!.type.schema.nodes.section_break.create({ sectionId: 's1' });
      return state.tr.insert(state.doc.content.size, brk);
    });
    const ids = serializeYDocToData(ydoc).sections.map((s) => s.id);
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids).toContain('s1_2');
  });

  it('a break naming an unknown id still yields a section, with the first section’s page setup', () => {
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, fourSections());
    edit(ydoc, (state) => {
      const brk = state.doc.type.schema.nodes.section_break.create({ sectionId: 'ghost' });
      return state.tr.insert(0, brk);
    });
    const back = serializeYDocToData(ydoc);
    expect(back.sections[0].id).toBe('ghost');
    expect(back.sections[0].pageConfig).toEqual(DEFAULT_PAGE_CONFIG);
    expect(back.sections.length).toBe(5);
  });

  it('hydrates and reads back the multi-section fixtures', async () => {
    for (const [file, count] of [['sections-headers.docx', 4], ['kitchen-sink.docx', 4], ['long-document.docx', 15]] as const) {
      const doc = new DocumentImpl('x', 'x');
      await importFromDocx(new Uint8Array(readFileSync(path.join(fixtures, file))), doc);
      const data = doc.getData();
      expect(data.sections.length).toBe(count);
      const ydoc = new Y.Doc();
      hydrateYDocFromData(ydoc, data);
      const back = serializeYDocToData(ydoc);
      expect(back.sections.map((s) => s.id)).toEqual(data.sections.map((s) => s.id));
      expect(back.sections.map((s) => s.pageConfig)).toEqual(data.sections.map((s) => s.pageConfig));
      expect(back.sections.map((s) => s.header)).toEqual(data.sections.map((s) => s.header));
      expect(back.sections.map((s) => s.footer)).toEqual(data.sections.map((s) => s.footer));
      expect(back.sections.map((s) => s.blocks.map((b) => b.type))).toEqual(data.sections.map((s) => s.blocks.map((b) => b.type)));
    }
  });
});

describe('degenerate input', () => {
  it('refuses to hydrate a document with no sections', () => {
    const data = fourSections();
    data.sections = [];
    expect(() => hydrateYDocFromData(new Y.Doc(), data)).toThrow('no sections');
  });

  it('an empty Y.Doc reads back as one empty section', () => {
    const back = serializeYDocToData(new Y.Doc());
    expect(back.sections.length).toBe(1);
    expect(back.sections[0].blocks.length).toBe(1);
  });
});

describe('two clients on a multi-section document', () => {
  it('converge after concurrent edits in different sections', () => {
    const a = new Y.Doc();
    hydrateYDocFromData(a, fourSections());
    const b = new Y.Doc();
    Y.applyUpdate(b, Y.encodeStateAsUpdate(a));

    // A types into section 1, B into section 4, while disconnected.
    edit(a, (state) => state.tr.insertText('A:', 1));
    edit(b, (state) => state.tr.insertText(':B', state.doc.content.size - 1));
    syncDocs(a, b);

    const da = serializeYDocToData(a);
    const db = serializeYDocToData(b);
    expect(da.sections.map((s) => text(s.blocks))).toEqual(db.sections.map((s) => text(s.blocks)));
    expect(text(da.sections[0].blocks)[0]).toBe('A:one a');
    expect(text(da.sections[3].blocks)).toEqual(['four:B']);
    expect(breakIds(a)).toEqual(['s1', 's2', 's3']);
    expect(breakIds(b)).toEqual(['s1', 's2', 's3']);
  });

  it('converge when one inserts a break while the other edits text', () => {
    const a = new Y.Doc();
    hydrateYDocFromData(a, fourSections());
    const b = new Y.Doc();
    Y.applyUpdate(b, Y.encodeStateAsUpdate(a));

    edit(a, (state) => {
      const brk = state.doc.type.schema.nodes.section_break.create({ sectionId: 'new1' });
      return state.tr.insert(state.doc.child(0).nodeSize, brk);
    });
    edit(b, (state) => state.tr.insertText('!', 1));
    syncDocs(a, b);

    expect(breakIds(a)).toEqual(breakIds(b));
    expect(serializeYDocToData(a).sections.map((s) => s.id)).toEqual(serializeYDocToData(b).sections.map((s) => s.id));
    expect(text(serializeYDocToData(b).sections[0].blocks)[0]).toBe('!one a');
  });

  it('attach through the collab binding: a creator seeds four sections and a joiner reads them', async () => {
    __resetInMemoryRooms();
    const creator = new DocumentImpl('room', 'Doc');
    creator.setData(fourSections());
    const ha = await creator.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const joiner = new DocumentImpl('room', 'Doc');
    const hb = await joiner.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });

    const fromJoiner = serializeYDocToData(hb.ydoc);
    expect(fromJoiner.sections.map((s) => s.id)).toEqual(['s1', 's2', 's3', 's4']);
    expect(fromJoiner.sections[1].pageConfig).toEqual(landscape);

    // An edit by the joiner reaches the creator, in the last section.
    edit(hb.ydoc, (state) => state.tr.insertText('B:', state.doc.content.size - 1 - 'four'.length));
    expect(text(serializeYDocToData(ha.ydoc).sections[3].blocks)).toEqual(['B:four']);
    ha.detach();
    hb.detach();
  });
});

describe('a single-section Y state written before section breaks existed', () => {
  /** What an older client wrote: no schemaVersion, one sectionsMeta entry, a body without section_break nodes. */
  function v1State(): Y.Doc {
    const ydoc = new Y.Doc();
    const fields = getYDocFields(ydoc);
    ydoc.transact(() => {
      fields.meta.set('id', 'old');
      fields.meta.set('title', 'Old');
      fields.meta.set('defaultPageConfig', DEFAULT_PAGE_CONFIG);
      const section = new Y.Map<unknown>();
      section.set('id', 'only');
      section.set('pageConfig', a4);
      fields.sectionsMeta.push([section]);
      updateYFragment(ydoc, fields.content, blocksToPmDoc([para('legacy one'), para('legacy two')]), {
        mapping: new Map(),
        isOMark: new Map(),
      });
    });
    return ydoc;
  }

  it('still loads as one section with its page setup and blocks', () => {
    const ydoc = new Y.Doc();
    Y.applyUpdate(ydoc, Y.encodeStateAsUpdate(v1State()));
    expect(readSchemaVersion(ydoc)).toBe(1);
    expect(() => assertSchemaVersion(ydoc)).not.toThrow();
    const data = serializeYDocToData(ydoc);
    expect(data.sections.map((s) => s.id)).toEqual(['only']);
    expect(data.sections[0].pageConfig).toEqual(a4);
    expect(text(data.sections[0].blocks)).toEqual(['legacy one', 'legacy two']);
  });

  it('a new client joins the room, sees the content, and does not reseed or rewrite it', async () => {
    __resetInMemoryRooms();
    const old = v1State();
    const oldProvider = new InMemoryProvider();
    await oldProvider.connect('room');
    // Publish the old client's state to the room the way a client's update would reach the relay.
    const { writeUpdate } = await import('y-protocols/sync');
    const enc = await import('lib0/encoding');
    const encoder = enc.createEncoder();
    writeUpdate(encoder, Y.encodeStateAsUpdate(old));
    oldProvider.send('doc', enc.toUint8Array(encoder));

    const joiner = new DocumentImpl('other', 'Other');
    joiner.setSectionBlocks(joiner.getSections()[0].id, [para('should not be seeded')]);
    const handle = await joiner.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    const data = serializeYDocToData(handle.ydoc);
    expect(text(data.sections[0].blocks)).toEqual(['legacy one', 'legacy two']);
    expect(data.sections.length).toBe(1);
    expect(handle.ydoc.getMap('meta').has('schemaVersion')).toBe(false);
    handle.detach();
    oldProvider.disconnect();
  });
});

describe('schemaVersion gate', () => {
  it('reads a missing, invalid or valid version', () => {
    const ydoc = new Y.Doc();
    expect(readSchemaVersion(ydoc)).toBe(1);
    const meta = ydoc.getMap('meta');
    meta.set('schemaVersion', 'two');
    expect(readSchemaVersion(ydoc)).toBe(1);
    meta.set('schemaVersion', 0);
    expect(readSchemaVersion(ydoc)).toBe(1);
    meta.set('schemaVersion', COLLAB_SCHEMA_VERSION);
    expect(readSchemaVersion(ydoc)).toBe(COLLAB_SCHEMA_VERSION);
    expect(() => assertSchemaVersion(ydoc)).not.toThrow();
  });

  it('refuses a version newer than the client supports, with a clear error', () => {
    const ydoc = new Y.Doc();
    ydoc.getMap('meta').set('schemaVersion', COLLAB_SCHEMA_VERSION + 1);
    expect(() => assertSchemaVersion(ydoc)).toThrow(CollabSchemaVersionError);
    try {
      assertSchemaVersion(ydoc);
    } catch (e) {
      const err = e as CollabSchemaVersionError;
      expect(err.roomVersion).toBe(COLLAB_SCHEMA_VERSION + 1);
      expect(err.supportedVersion).toBe(COLLAB_SCHEMA_VERSION);
      expect(err.message).toContain(`version ${COLLAB_SCHEMA_VERSION + 1}`);
      expect(err.message).toContain('Update the editor');
    }
  });

  it('attachCollab to a newer room rejects, disconnects, seeds nothing and leaves the document detached', async () => {
    __resetInMemoryRooms();
    const creator = new DocumentImpl('room', 'Doc');
    creator.setSectionBlocks(creator.getSections()[0].id, [para('hello')]);
    const ha = await creator.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    // A newer client (or a later release) raises the room's version.
    ha.ydoc.getMap('meta').set('schemaVersion', COLLAB_SCHEMA_VERSION + 1);

    class Spy extends InMemoryProvider {
      disconnected = false;
      override disconnect(): void {
        this.disconnected = true;
        super.disconnect();
      }
    }
    const spy = new Spy();
    const joiner = new DocumentImpl('room', 'Doc');
    await expect(joiner.attachCollab(spy, identity('B'), { roomId: 'room' })).rejects.toBeInstanceOf(CollabSchemaVersionError);
    expect(spy.disconnected).toBe(true);
    expect(joiner.getCollabHandle()).toBeNull();
    // The failed attach can be retried (it is not stuck "already being attached").
    await expect(joiner.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' })).rejects.toBeInstanceOf(CollabSchemaVersionError);
    ha.detach();
  });
});

describe('getCollabBlockers', () => {
  it('is empty for a plain document and for a multi-section one', () => {
    expect(getCollabBlockers(new DocumentImpl('a', 'a').getData())).toEqual([]);
    expect(getCollabBlockers(fourSections())).toEqual([]);
  });

  it('reports a document with no sections', () => {
    expect(getCollabBlockers({ sections: [] })).toEqual([
      { id: 'collab.noSections', message: 'The document has no sections', count: 1 },
    ]);
  });

  it('reports sections that share an id', () => {
    const data = fourSections();
    data.sections[2].id = 's1';
    data.sections[3].id = 's1';
    const blockers = getCollabBlockers(data);
    expect(blockers.map((b) => [b.id, b.count])).toEqual([['collab.duplicateSectionIds', 2]]);
  });

  it('reports block and inline types this version does not know, also inside table cells', () => {
    const data = fourSections();
    const table = data.sections[1].blocks[2];
    if (table.type !== 'table') throw new Error('fixture');
    table.rows[0].cells[0].content.push({ type: 'smartArt' } as never);
    table.rows[0].cells[1].blocks = [{ id: 'x', type: 'chart' } as never];
    data.sections[0].blocks.push({ id: 'y', type: 'equation' } as never);
    (data.sections[0].blocks[0] as ParagraphBlock).content.push({ type: 'mention' } as never);
    expect(getCollabBlockers(data).map((b) => [b.id, b.count])).toEqual([
      ['collab.unsupportedBlock', 2],
      ['collab.unsupportedInline', 2],
    ]);
  });

  it('attachCollab refuses a blocked document with the list, and attaches nothing', async () => {
    __resetInMemoryRooms();
    const doc = new DocumentImpl('room', 'Doc');
    const data = fourSections();
    data.sections[3].id = 's1';
    doc.setData(data);
    const err = await doc.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' }).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(CollabBlockedError);
    expect((err as CollabBlockedError).blockers.map((b) => b.id)).toEqual(['collab.duplicateSectionIds']);
    expect((err as CollabBlockedError).message).toContain('share the same id');
    expect(doc.getCollabHandle()).toBeNull();
  });
});

describe('splitSection under collaboration', () => {
  it('adds the new section to the shared document before the later one, and a peer that already wrote it is followed', async () => {
    __resetInMemoryRooms();
    const a = new DocumentImpl('room', 'Doc');
    a.setData(fourSections());
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    a.splitSection('s2', 'mid', 'continuous');
    const ids = getYDocFields(ha.ydoc).sectionsMeta.toArray().map((m) => m.get('id'));
    expect(ids).toEqual(['s1', 'mid', 's2', 's3', 's4']);
    const later = getYDocFields(ha.ydoc).sectionsMeta.get(2);
    expect(later.get('breakType')).toBe('continuous');
    expect(later.has('footer')).toBe(false);
    expect(getYDocFields(ha.ydoc).sectionsMeta.get(1).get('footer')).toEqual(fourSections().sections[1].footer);

    // A second client's model learns of the section from the room and follows the writer's start type.
    const b = new DocumentImpl('room', 'Doc');
    b.setData(fourSections());
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    const split = b.splitSection('s2', 'mid', 'evenPage');
    expect(split?.id).toBe('mid');
    expect(b.getSection('s2')?.breakType).toBe('continuous');
    expect(getYDocFields(hb.ydoc).sectionsMeta.length).toBe(5);
    ha.detach();
    hb.detach();
  });
});
