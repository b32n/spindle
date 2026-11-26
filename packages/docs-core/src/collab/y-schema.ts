// Y.Doc schema for collaborative documents.
//
// Layout:
//   meta:               Y.Map<string, JSON>  — id, title, defaultPageConfig, timestamps, schemaVersion
//   content:            Y.XmlFragment        — body text (ProseMirror-compatible, bound via y-prosemirror)
//   sectionsMeta:       Y.Array<Y.Map>       — per-section metadata (id, pageConfig (with columns), header, footer,
//                                              pageNumbering, verticalAlign, breakType)
//   namedStyles:        Y.Map<string, JSON>  — DOCX named style id → NamedStyle (additive)
//   textStylePool:      Y.Map<string, JSON>  — styleId → TextStyle
//   paragraphStylePool: Y.Map<string, JSON>  — styleId → ParagraphStyle
//
// Sections. The blocks of all sections live in the one Y.XmlFragment; a
// `section_break` node ends every section but the last and names it by id
// (attrs.sectionId). sectionsMeta holds the setup of each section, looked up by
// that id; the last section has no break, and is the last sectionsMeta entry
// that no break names. A document with no break is a single section, so every
// document written before sections were co-edited is still valid unchanged.
// Entries of deleted breaks stay in sectionsMeta unused and are not read back.
//
// schemaVersion (meta): the version of this layout the room was written with.
// A missing key means 1 (no section breaks). A client refuses to join a room
// whose version is newer than COLLAB_SCHEMA_VERSION (see assertSchemaVersion),
// because y-prosemirror drops nodes it cannot instantiate.

import * as Y from 'yjs';
import {
  prosemirrorToYXmlFragment,
  yXmlFragmentToProseMirrorRootNode,
} from 'y-prosemirror';
import { docsSchema } from '../prosemirror/schema';
import { TextStylePoolImpl } from '../style-pool';
import {
  sectionsToPmDoc,
  proseMirrorToDocument,
} from '../prosemirror/sync';
import type {
  Document,
  DocumentData,
  HeaderFooterContent,
  NamedStyle,
  PageConfig,
  PageNumbering,
  ParagraphStyle,
  SectionBreakType,
  SectionVerticalAlign,
  SectionData,
  TextStyle,
} from '../types';
import type { DocsCommentThread } from '../comments';

/**
 * The Y document layout this build reads and writes. 1 is the layout without section breaks (the key is then absent),
 * 2 adds `section_break` nodes. Raised by any later change an older client could not read.
 */
export const COLLAB_SCHEMA_VERSION = 2;

/** Thrown when a room was written with a newer Y document layout than this client understands. */
export class CollabSchemaVersionError extends Error {
  readonly roomVersion: number;
  readonly supportedVersion: number;
  constructor(roomVersion: number, supportedVersion: number = COLLAB_SCHEMA_VERSION) {
    super(
      `This collaboration room uses document schema version ${roomVersion}, but this editor only supports up to ` +
        `version ${supportedVersion}. Update the editor to join the room.`,
    );
    this.name = 'CollabSchemaVersionError';
    this.roomVersion = roomVersion;
    this.supportedVersion = supportedVersion;
  }
}

/** The schema version a Y.Doc was written with (1 when the key is absent or not a number). */
export function readSchemaVersion(ydoc: Y.Doc): number {
  const v = ydoc.getMap<unknown>('meta').get('schemaVersion');
  return typeof v === 'number' && Number.isFinite(v) && v >= 1 ? v : 1;
}

/** Throw a CollabSchemaVersionError when the Y.Doc was written with a newer schema than this client supports. */
export function assertSchemaVersion(ydoc: Y.Doc): void {
  const version = readSchemaVersion(ydoc);
  if (version > COLLAB_SCHEMA_VERSION) throw new CollabSchemaVersionError(version);
}

export interface YDocFields {
  meta: Y.Map<unknown>;
  content: Y.XmlFragment;
  sectionsMeta: Y.Array<Y.Map<unknown>>;
  textStylePool: Y.Map<unknown>;
  paragraphStylePool: Y.Map<unknown>;
  /**
   * Comment threads, keyed by thread id (plain JSON). The `comment` mark in
   * the body content is the position anchor; this map holds thread content.
   */
  threads: Y.Map<unknown>;
  /** Named styles by style id (plain JSON; styles are not co-edited text). */
  namedStyles: Y.Map<unknown>;
}

/** Return the named Y types used by the schema. Idempotent. */
/** Reject object keys that would poison a plain object's prototype when a key
 *  comes from peer-controlled collab data. */
function isUnsafeKey(k: string): boolean {
  return k === '__proto__' || k === 'constructor' || k === 'prototype';
}

export function getYDocFields(ydoc: Y.Doc): YDocFields {
  return {
    meta: ydoc.getMap<unknown>('meta'),
    content: ydoc.getXmlFragment('content'),
    sectionsMeta: ydoc.getArray<Y.Map<unknown>>('sectionsMeta'),
    textStylePool: ydoc.getMap<unknown>('textStylePool'),
    paragraphStylePool: ydoc.getMap<unknown>('paragraphStylePool'),
    threads: ydoc.getMap<unknown>('threads'),
    namedStyles: ydoc.getMap<unknown>('namedStyles'),
  };
}

/**
 * Initialize an empty Y.Doc from a DocumentData snapshot. Intended for
 * first-time hydration only — once collab has started, updates must come
 * through y-prosemirror / Y.applyUpdate rather than re-running this.
 */
export function hydrateYDocFromData(ydoc: Y.Doc, data: DocumentData): void {
  if (data.sections.length === 0) {
    throw new Error('Cannot hydrate Y.Doc from a document with no sections');
  }
  const fields = getYDocFields(ydoc);

  ydoc.transact(() => {
    // Document-level meta
    fields.meta.set('schemaVersion', COLLAB_SCHEMA_VERSION);
    fields.meta.set('id', data.id);
    fields.meta.set('title', data.title);
    fields.meta.set('defaultPageConfig', data.defaultPageConfig);
    if (data.createdAt) fields.meta.set('createdAt', data.createdAt);
    if (data.updatedAt) fields.meta.set('updatedAt', data.updatedAt);

    // Style pools (Y.Map of plain JSON values — style records aren't edited
    // text, just keyed objects).
    for (const [id, style] of Object.entries(data.textStylePool)) {
      fields.textStylePool.set(id, style);
    }
    for (const [id, style] of Object.entries(data.paragraphStylePool)) {
      fields.paragraphStylePool.set(id, style);
    }

    // Section metadata (one entry per section, in document order). Page config /
    // header / footer are stored as plain JSON on the Y.Map; they aren't co-edited.
    for (const section of data.sections) {
      const sectionMap = new Y.Map<unknown>();
      sectionMap.set('id', section.id);
      sectionMap.set('pageConfig', section.pageConfig);
      if (section.header) sectionMap.set('header', section.header);
      if (section.footer) sectionMap.set('footer', section.footer);
      if (section.pageNumbering) sectionMap.set('pageNumbering', section.pageNumbering);
      if (section.verticalAlign) sectionMap.set('verticalAlign', section.verticalAlign);
      if (section.breakType) sectionMap.set('breakType', section.breakType);
      fields.sectionsMeta.push([sectionMap]);
    }

    // Named styles — keyed by style id, stored as plain JSON.
    for (const [id, style] of Object.entries(data.namedStyles ?? {})) {
      if (!isUnsafeKey(id)) fields.namedStyles.set(id, style);
    }

    // Comment threads — keyed by thread id, stored as plain JSON.
    if (data.threads) {
      for (const thread of data.threads) {
        fields.threads.set(thread.id, thread);
      }
    }

    // Content: hydrate the Y.XmlFragment from a PM document built by
    // blocksToPmDoc — which performs the list-item grouping pass that the
    // raw documentToProseMirror converter skips (the latter emits an
    // internal "__list_item__" sentinel meant for that grouping step).
    // Runs may reference the text style pool by id; resolve them to marks here
    // or the formatting is silently dropped when the body is seeded. Sections
    // after the first are separated by section_break nodes.
    const textPool = new TextStylePoolImpl();
    textPool.setFromData(data.textStylePool ?? {});
    const pmDoc = sectionsToPmDoc(data.sections, textPool);
    prosemirrorToYXmlFragment(pmDoc, fields.content);
  });
}

/**
 * Read the Y.Doc's current state back into a plain DocumentData object.
 * Used by DocumentImpl.getData() when collab is attached, so consumers see
 * the same JSON shape with or without collab.
 */
export function serializeYDocToData(ydoc: Y.Doc): DocumentData {
  const fields = getYDocFields(ydoc);

  const id = (fields.meta.get('id') as string | undefined) ?? '';
  const title = (fields.meta.get('title') as string | undefined) ?? '';
  const defaultPageConfig = fields.meta.get('defaultPageConfig') as PageConfig;
  const createdAt = fields.meta.get('createdAt') as string | undefined;
  const updatedAt = fields.meta.get('updatedAt') as string | undefined;

  const textStylePool: Record<string, TextStyle> = {};
  for (const [k, v] of fields.textStylePool.entries()) {
    if (isUnsafeKey(k)) continue; // guard against __proto__/constructor from a peer
    textStylePool[k] = v as TextStyle;
  }
  const paragraphStylePool: Record<string, ParagraphStyle> = {};
  for (const [k, v] of fields.paragraphStylePool.entries()) {
    if (isUnsafeKey(k)) continue;
    paragraphStylePool[k] = v as ParagraphStyle;
  }

  // Reconstruct the sections' blocks by routing the Y.XmlFragment through
  // the existing PM → block converter. existingDoc provides the metadata
  // that proseMirrorToDocument needs to preserve.
  const pmDoc = yXmlFragmentToProseMirrorRootNode(fields.content, docsSchema);

  const sections: SectionData[] = [];
  fields.sectionsMeta.forEach((sectionMap) => {
    const sectionId = sectionMap.get('id') as string;
    const pageConfig = sectionMap.get('pageConfig') as PageConfig;
    const header = sectionMap.get('header') as HeaderFooterContent | undefined;
    const footer = sectionMap.get('footer') as HeaderFooterContent | undefined;
    const pageNumbering = sectionMap.get('pageNumbering') as PageNumbering | undefined;
    const verticalAlign = sectionMap.get('verticalAlign') as SectionVerticalAlign | undefined;
    const breakType = sectionMap.get('breakType') as SectionBreakType | undefined;
    sections.push({
      id: sectionId,
      pageConfig,
      blocks: [], // filled in below from the PM tree
      header,
      footer,
      ...(pageNumbering && { pageNumbering }),
      ...(verticalAlign && { verticalAlign }),
      ...(breakType && { breakType }),
    });
  });

  if (sections.length === 0) {
    // Shouldn't happen for a hydrated Y.Doc, but stay defensive.
    sections.push({
      id: `section_${Date.now()}`,
      pageConfig: defaultPageConfig,
      blocks: [],
    });
  }

  // Each section_break names the section it ends; the blocks after the last one form the last section. A body
  // with no break is one section: the last entry of sectionsMeta (the only one in a document written before sections
  // were co-edited).
  let hasBreak = false;
  pmDoc.forEach((node) => {
    if (node.type.name === 'section_break') hasBreak = true;
  });
  const seedDoc: Document = {
    id,
    title,
    defaultPageConfig,
    sections: hasBreak ? sections : [sections[sections.length - 1]],
    createdAt,
    updatedAt,
  };
  const rebuilt = proseMirrorToDocument(pmDoc, seedDoc);
  const orderedSections: SectionData[] = rebuilt.sections;

  const namedStyles = readNamedStyles(fields.namedStyles);

  const threads: DocsCommentThread[] = [];
  for (const value of fields.threads.values()) {
    threads.push(value as DocsCommentThread);
  }

  return {
    id,
    title,
    sections: orderedSections,
    defaultPageConfig,
    textStylePool,
    paragraphStylePool,
    createdAt,
    updatedAt,
    threads: threads.length > 0 ? threads : undefined,
    namedStyles,
  };
}

/** The named styles of a Y map as plain JSON (undefined when there are none). */
export function readNamedStyles(map: Y.Map<unknown>): Record<string, NamedStyle> | undefined {
  const out: Record<string, NamedStyle> = {};
  for (const [k, v] of map.entries()) {
    if (isUnsafeKey(k)) continue; // guard against __proto__/constructor from a peer
    out[k] = v as NamedStyle;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}
