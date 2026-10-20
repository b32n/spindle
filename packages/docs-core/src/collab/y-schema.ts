// Y.Doc schema for collaborative documents.
//
// Layout:
//   meta:               Y.Map<string, JSON>  — id, title, defaultPageConfig, timestamps
//   content:            Y.XmlFragment        — body text (ProseMirror-compatible, bound via y-prosemirror)
//   sectionsMeta:       Y.Array<Y.Map>       — per-section metadata (id, pageConfig, header, footer)
//   textStylePool:      Y.Map<string, JSON>  — styleId → TextStyle
//   paragraphStylePool: Y.Map<string, JSON>  — styleId → ParagraphStyle
//
// v1 supports single-section documents in collaboration mode. The blocks of
// all sections live in a single Y.XmlFragment; sectionsMeta holds page
// configuration and header/footer for that one section. Multi-section
// collab will need block-level section markers in the PM schema and is
// deferred. hydrateYDocFromData throws if asked to load a multi-section doc.

import * as Y from 'yjs';
import {
  prosemirrorToYXmlFragment,
  yXmlFragmentToProseMirrorRootNode,
} from 'y-prosemirror';
import { docsSchema } from '../prosemirror/schema';
import { TextStylePoolImpl } from '../style-pool';
import {
  blocksToPmDoc,
  proseMirrorToDocument,
} from '../prosemirror/sync';
import type {
  Document,
  DocumentData,
  HeaderFooterContent,
  PageConfig,
  ParagraphStyle,
  SectionData,
  TextStyle,
} from '../types';
import type { DocsCommentThread } from '../comments';

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
  if (data.sections.length > 1) {
    throw new Error(
      'Collaboration v1 supports only single-section documents; ' +
        `received ${data.sections.length} sections`,
    );
  }

  const fields = getYDocFields(ydoc);

  ydoc.transact(() => {
    // Document-level meta
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

    // Section metadata (one entry per section). Page config / header /
    // footer are stored as plain JSON on the Y.Map; they aren't co-edited
    // in v1.
    for (const section of data.sections) {
      const sectionMap = new Y.Map<unknown>();
      sectionMap.set('id', section.id);
      sectionMap.set('pageConfig', section.pageConfig);
      if (section.header) sectionMap.set('header', section.header);
      if (section.footer) sectionMap.set('footer', section.footer);
      fields.sectionsMeta.push([sectionMap]);
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
    // or the formatting is silently dropped when the body is seeded.
    const textPool = new TextStylePoolImpl();
    textPool.setFromData(data.textStylePool ?? {});
    const pmDoc = blocksToPmDoc(data.sections[0].blocks, textPool);
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

  // Reconstruct the section's blocks by routing the Y.XmlFragment through
  // the existing PM → block converter. existingDoc provides the metadata
  // that proseMirrorToDocument needs to preserve.
  const pmDoc = yXmlFragmentToProseMirrorRootNode(fields.content, docsSchema);

  const sections: SectionData[] = [];
  fields.sectionsMeta.forEach((sectionMap) => {
    const sectionId = sectionMap.get('id') as string;
    const pageConfig = sectionMap.get('pageConfig') as PageConfig;
    const header = sectionMap.get('header') as HeaderFooterContent | undefined;
    const footer = sectionMap.get('footer') as HeaderFooterContent | undefined;
    sections.push({
      id: sectionId,
      pageConfig,
      blocks: [], // filled in below from the PM tree
      header,
      footer,
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

  // Single-section v1 — drop all PM blocks into sections[0].
  const seedDoc: Document = {
    id,
    title,
    defaultPageConfig,
    sections: [
      {
        id: sections[0].id,
        pageConfig: sections[0].pageConfig,
        blocks: [],
      },
    ],
    createdAt,
    updatedAt,
  };
  const rebuilt = proseMirrorToDocument(pmDoc, seedDoc);
  sections[0].blocks = rebuilt.sections[0].blocks;

  const threads: DocsCommentThread[] = [];
  for (const value of fields.threads.values()) {
    threads.push(value as DocsCommentThread);
  }

  return {
    id,
    title,
    sections,
    defaultPageConfig,
    textStylePool,
    paragraphStylePool,
    createdAt,
    updatedAt,
    threads: threads.length > 0 ? threads : undefined,
  };
}
