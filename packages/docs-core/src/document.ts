// Document model and operations

import type {
  Document,
  DocumentData,
  Section,
  SectionData,
  Block,
  PageConfig,
  TextSelection,
  DocumentEventHandler,
  DocumentEventType,
} from './types';
import { DEFAULT_PAGE_CONFIG, HeaderFooterContent } from './types';
import { DocumentEventEmitter } from './event-emitter';
import { DocumentHistory } from './history';
import { TextStylePoolImpl, ParagraphStylePoolImpl } from './style-pool';
import { createParagraphFromText } from './blocks/paragraph';
import { attachCollabToYDoc } from './collab/binding';
import { getYDocFields } from './collab/y-schema';
import { DocsCommentStore, type DocsCommentEvent, type DocsCommentThread } from './comments';

function generateDocumentId(): string {
  return `doc_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
}

function generateSectionId(): string {
  return `section_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
}

/**
 * Create a new section with default configuration
 */
export function createSection(
  pageConfig?: PageConfig,
  blocks?: Block[]
): Section {
  return {
    id: generateSectionId(),
    pageConfig: pageConfig || { ...DEFAULT_PAGE_CONFIG },
    blocks: blocks || [createParagraphFromText('')],
  };
}

/**
 * Create a new empty document
 */
export function createDocument(
  title: string = 'Untitled Document',
  pageConfig?: PageConfig
): Document {
  const config = pageConfig || { ...DEFAULT_PAGE_CONFIG };
  return {
    id: generateDocumentId(),
    title,
    sections: [createSection(config)],
    defaultPageConfig: config,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}

/**
 * Document implementation class
 */
export class DocumentImpl {
  private document: Document;
  private eventEmitter: DocumentEventEmitter;
  private history: DocumentHistory;
  private textStylePool: TextStylePoolImpl;
  private paragraphStylePool: ParagraphStylePoolImpl;
  private selection: TextSelection | null = null;
  // Set when attachCollab is called. The Y.Doc inside the handle is the
  // source of truth for body content while attached; local mutations
  // (history snapshots, addBlock, etc.) keep working on this.document but
  // the React editor binds to handle.xmlFragment via ySyncPlugin.
  private collabHandle: import('./collab/binding').CollabHandle | null = null;
  // Comment threads. The `comment` mark in the body is the position anchor;
  // this store holds thread content (comments, replies, status).
  private commentStore: DocsCommentStore = new DocsCommentStore();
  // True while mirrorThreads is writing to the Y.Doc — lets the threads
  // observer ignore our own writes.
  private mirroringThreads = false;
  private threadsObserver: (() => void) | null = null;

  constructor(id?: string, title?: string) {
    this.document = createDocument(title || 'Untitled Document');
    if (id) {
      this.document.id = id;
    }
    this.eventEmitter = new DocumentEventEmitter();
    this.history = new DocumentHistory();
    this.textStylePool = new TextStylePoolImpl();
    this.paragraphStylePool = new ParagraphStylePoolImpl();
    this.commentStore.__setChangeListener((event) => this.onCommentChange(event));

    // Record initial state
    this.recordHistory('Initial state');
  }

  /** The document's comment-thread store. */
  getComments(): DocsCommentStore {
    return this.commentStore;
  }

  /**
   * Comment-mutation handler: mirror threads to the Y.Doc (when collab is
   * attached), trigger a UI re-render, and surface a semantic `commentEvent`.
   */
  private onCommentChange(event: DocsCommentEvent): void {
    this.mirrorThreads();
    this.emit('commentChange', {});
    this.emit('commentEvent', event);
  }

  /** Re-sync the comment store's threads into the Y.Doc's threads map. */
  private mirrorThreads(): void {
    const handle = this.collabHandle;
    if (!handle) return;
    const threadsY = getYDocFields(handle.ydoc).threads;
    this.mirroringThreads = true;
    try {
      handle.ydoc.transact(() => {
        threadsY.clear();
        for (const thread of this.commentStore.toJSON()) {
          threadsY.set(thread.id, thread);
        }
      });
    } finally {
      this.mirroringThreads = false;
    }
  }

  /** Load the comment store from the Y.Doc's threads map (remote changes). */
  private loadThreadsFromY(): void {
    const handle = this.collabHandle;
    if (!handle) return;
    const threadsY = getYDocFields(handle.ydoc).threads;
    const threads: DocsCommentThread[] = [];
    for (const value of threadsY.values()) {
      threads.push(value as DocsCommentThread);
    }
    this.commentStore.loadJSON(threads);
    this.emit('commentChange', {});
  }

  // ============================================================================
  // Getters
  // ============================================================================

  getId(): string {
    return this.document.id;
  }

  getTitle(): string {
    return this.document.title;
  }

  getSections(): Section[] {
    return this.document.sections;
  }

  getSection(sectionId: string): Section | undefined {
    return this.document.sections.find(s => s.id === sectionId);
  }

  getDefaultPageConfig(): PageConfig {
    return this.document.defaultPageConfig;
  }

  getSelection(): TextSelection | null {
    return this.selection;
  }

  getTextStylePool(): TextStylePoolImpl {
    return this.textStylePool;
  }

  getParagraphStylePool(): ParagraphStylePoolImpl {
    return this.paragraphStylePool;
  }

  // ============================================================================
  // Document Operations
  // ============================================================================

  setTitle(title: string): void {
    this.document.title = title;
    this.document.updatedAt = new Date().toISOString();
    this.emit('documentChange', { title });
  }

  setDefaultPageConfig(config: PageConfig): void {
    this.document.defaultPageConfig = { ...config };
    this.document.updatedAt = new Date().toISOString();
    this.emit('pageConfigChange', { config });
  }

  // ============================================================================
  // Section Operations
  // ============================================================================

  addSection(pageConfig?: PageConfig): Section {
    const section = createSection(pageConfig || this.document.defaultPageConfig);
    this.document.sections.push(section);
    this.document.updatedAt = new Date().toISOString();
    this.emit('documentChange', { action: 'addSection', sectionId: section.id });
    return section;
  }

  insertSection(index: number, pageConfig?: PageConfig): Section {
    const section = createSection(pageConfig || this.document.defaultPageConfig);
    this.document.sections.splice(index, 0, section);
    this.document.updatedAt = new Date().toISOString();
    this.emit('documentChange', { action: 'insertSection', sectionId: section.id, index });
    return section;
  }

  deleteSection(sectionId: string): boolean {
    if (this.document.sections.length <= 1) {
      return false; // Keep at least one section
    }
    
    const index = this.document.sections.findIndex(s => s.id === sectionId);
    if (index === -1) return false;
    
    this.document.sections.splice(index, 1);
    this.document.updatedAt = new Date().toISOString();
    this.emit('documentChange', { action: 'deleteSection', sectionId });
    return true;
  }

  setSectionPageConfig(sectionId: string, config: PageConfig): void {
    const section = this.getSection(sectionId);
    if (!section) return;
    
    section.pageConfig = { ...config };
    this.document.updatedAt = new Date().toISOString();
    this.emit('pageConfigChange', { sectionId, config });
  }

  /**
   * Replace all blocks in a section
   */
  setSectionBlocks(sectionId: string, blocks: Block[]): void {
    const section = this.getSection(sectionId);
    if (!section) return;
    
    section.blocks = blocks;
    this.document.updatedAt = new Date().toISOString();
    this.emit('documentChange', { action: 'setSectionBlocks', sectionId });
  }

  // ============================================================================
  // Header/Footer Operations
  // ============================================================================

  /**
   * Get the header content for a section
   */
  getSectionHeader(sectionId: string): HeaderFooterContent | undefined {
    const section = this.getSection(sectionId);
    return section?.header;
  }

  /**
   * Set the header content for a section
   */
  setSectionHeader(sectionId: string, header: HeaderFooterContent | undefined): void {
    const section = this.getSection(sectionId);
    if (!section) return;
    
    section.header = header;
    this.document.updatedAt = new Date().toISOString();
    this.emit('documentChange', { action: 'setSectionHeader', sectionId });
  }

  /**
   * Get the footer content for a section
   */
  getSectionFooter(sectionId: string): HeaderFooterContent | undefined {
    const section = this.getSection(sectionId);
    return section?.footer;
  }

  /**
   * Set the footer content for a section
   */
  setSectionFooter(sectionId: string, footer: HeaderFooterContent | undefined): void {
    const section = this.getSection(sectionId);
    if (!section) return;
    
    section.footer = footer;
    this.document.updatedAt = new Date().toISOString();
    this.emit('documentChange', { action: 'setSectionFooter', sectionId });
  }

  // ============================================================================
  // Block Operations
  // ============================================================================

  getBlock(blockId: string): Block | undefined {
    for (const section of this.document.sections) {
      const block = section.blocks.find(b => b.id === blockId);
      if (block) return block;
    }
    return undefined;
  }

  getBlockSection(blockId: string): Section | undefined {
    for (const section of this.document.sections) {
      if (section.blocks.some(b => b.id === blockId)) {
        return section;
      }
    }
    return undefined;
  }

  addBlock(sectionId: string, block: Block): void {
    const section = this.getSection(sectionId);
    if (!section) return;
    
    section.blocks.push(block);
    this.document.updatedAt = new Date().toISOString();
    this.emit('blockAdd', { sectionId, blockId: block.id });
  }

  insertBlock(sectionId: string, index: number, block: Block): void {
    const section = this.getSection(sectionId);
    if (!section) return;
    
    section.blocks.splice(index, 0, block);
    this.document.updatedAt = new Date().toISOString();
    this.emit('blockAdd', { sectionId, blockId: block.id, index });
  }

  updateBlock(blockId: string, updates: Partial<Block>): void {
    for (const section of this.document.sections) {
      const index = section.blocks.findIndex(b => b.id === blockId);
      if (index !== -1) {
        section.blocks[index] = { ...section.blocks[index], ...updates } as Block;
        this.document.updatedAt = new Date().toISOString();
        this.emit('blockChange', { sectionId: section.id, blockId });
        return;
      }
    }
  }

  replaceBlock(blockId: string, newBlock: Block): void {
    for (const section of this.document.sections) {
      const index = section.blocks.findIndex(b => b.id === blockId);
      if (index !== -1) {
        section.blocks[index] = newBlock;
        this.document.updatedAt = new Date().toISOString();
        this.emit('blockChange', { sectionId: section.id, blockId: newBlock.id });
        return;
      }
    }
  }

  deleteBlock(blockId: string): boolean {
    for (const section of this.document.sections) {
      const index = section.blocks.findIndex(b => b.id === blockId);
      if (index !== -1) {
        // Don't delete the last block in a section
        if (section.blocks.length <= 1) {
          // Replace with empty paragraph instead
          section.blocks[0] = createParagraphFromText('');
          this.emit('blockChange', { sectionId: section.id, blockId: section.blocks[0].id });
          return true;
        }
        
        section.blocks.splice(index, 1);
        this.document.updatedAt = new Date().toISOString();
        this.emit('blockDelete', { sectionId: section.id, blockId });
        return true;
      }
    }
    return false;
  }

  getBlockIndex(blockId: string): { sectionIndex: number; blockIndex: number } | null {
    for (let si = 0; si < this.document.sections.length; si++) {
      const section = this.document.sections[si];
      const bi = section.blocks.findIndex(b => b.id === blockId);
      if (bi !== -1) {
        return { sectionIndex: si, blockIndex: bi };
      }
    }
    return null;
  }

  // ============================================================================
  // Selection
  // ============================================================================

  setSelection(selection: TextSelection | null): void {
    this.selection = selection;
    this.emit('selectionChange', { selection });
  }

  // ============================================================================
  // History (Undo/Redo)
  // ============================================================================

  recordHistory(description?: string): void {
    this.history.record(this.document.sections, this.selection || undefined, description);
    this.emit('historyChange', { canUndo: this.canUndo(), canRedo: this.canRedo() });
  }

  undo(): boolean {
    const snapshot = this.history.undo(this.document.sections, this.selection || undefined);
    if (snapshot) {
      this.document.sections = snapshot.sections;
      this.selection = snapshot.selection || null;
      this.document.updatedAt = new Date().toISOString();
      this.emit('documentChange', { action: 'undo' });
      this.emit('historyChange', { canUndo: this.canUndo(), canRedo: this.canRedo() });
      return true;
    }
    return false;
  }

  redo(): boolean {
    const snapshot = this.history.redo(this.document.sections, this.selection || undefined);
    if (snapshot) {
      this.document.sections = snapshot.sections;
      this.selection = snapshot.selection || null;
      this.document.updatedAt = new Date().toISOString();
      this.emit('documentChange', { action: 'redo' });
      this.emit('historyChange', { canUndo: this.canUndo(), canRedo: this.canRedo() });
      return true;
    }
    return false;
  }

  canUndo(): boolean {
    return this.history.canUndo();
  }

  canRedo(): boolean {
    return this.history.canRedo();
  }

  // ============================================================================
  // Events
  // ============================================================================

  on(event: DocumentEventType, handler: DocumentEventHandler): () => void {
    return this.eventEmitter.on(event, handler);
  }

  off(event: DocumentEventType, handler: DocumentEventHandler): void {
    this.eventEmitter.off(event, handler);
  }

  private emit(event: DocumentEventType, payload: unknown): void {
    this.eventEmitter.emit(event, payload);
  }

  batch(operations: () => void): void {
    this.eventEmitter.batch(operations);
  }

  // ============================================================================
  // Serialization
  // ============================================================================

  getData(): DocumentData {
    const sectionData: SectionData[] = this.document.sections.map(section => ({
      id: section.id,
      pageConfig: section.pageConfig,
      blocks: section.blocks,
      header: section.header,
      footer: section.footer,
    }));

    return {
      id: this.document.id,
      title: this.document.title,
      sections: sectionData,
      defaultPageConfig: this.document.defaultPageConfig,
      textStylePool: this.textStylePool.toData(),
      paragraphStylePool: this.paragraphStylePool.toData(),
      createdAt: this.document.createdAt,
      updatedAt: this.document.updatedAt,
      threads: this.commentStore.toJSON(),
    };
  }

  setData(data: DocumentData): void {
    if (this.collabHandle) {
      throw new Error(
        'Cannot setData while collaboration is attached — detach first.',
      );
    }
    this.document = {
      id: data.id,
      title: data.title,
      sections: data.sections.map(s => ({
        id: s.id,
        pageConfig: s.pageConfig,
        blocks: s.blocks,
        header: s.header,
        footer: s.footer,
      })),
      defaultPageConfig: data.defaultPageConfig,
      createdAt: data.createdAt,
      updatedAt: data.updatedAt,
    };

    this.textStylePool.setFromData(data.textStylePool || {});
    this.paragraphStylePool.setFromData(data.paragraphStylePool || {});
    this.commentStore.loadJSON(data.threads);

    this.history.clear();
    this.recordHistory('Loaded document');

    this.emit('documentChange', { action: 'setData' });
  }

  // ============================================================================
  // Collaboration
  // ============================================================================

  /**
   * Attach a CollabProvider so this document's body content syncs with peers
   * via Yjs. The returned handle exposes the Y.Doc + Y.XmlFragment + Awareness
   * needed by the React editor's ySyncPlugin / yCursorPlugin. Idempotent
   * guard: throws if already attached.
   *
   * v1 restrictions: documents must be single-section; setData is disallowed
   * while attached; local DocumentHistory undo is shadowed by Yjs UndoManager
   * once we wire that in (phase 1.6).
   */
  async attachCollab(
    provider: import('@b32nio/spindle-shared').CollabProvider,
    identity: import('@b32nio/spindle-shared').CollabIdentity,
    options?: import('./collab/binding').AttachCollabOptions,
  ): Promise<import('./collab/binding').CollabHandle> {
    if (this.collabHandle) {
      throw new Error('Collaboration is already attached to this document.');
    }
    const handle = await attachCollabToYDoc(
      this.getData(),
      provider,
      identity,
      options,
    );
    this.collabHandle = handle;

    // Reconcile the comment store with the Y.Doc's threads map, then keep it
    // live: on first attach the map was hydrated from getData(); on a late
    // join it already holds peers' threads — loadThreadsFromY covers both.
    this.loadThreadsFromY();
    const threadsY = getYDocFields(handle.ydoc).threads;
    this.threadsObserver = () => {
      if (!this.mirroringThreads) this.loadThreadsFromY();
    };
    threadsY.observe(this.threadsObserver);

    this.emit('documentChange', { action: 'attachCollab' });
    return handle;
  }

  /** Return the live collab handle, or null if not in collab mode. */
  getCollabHandle(): import('./collab/binding').CollabHandle | null {
    return this.collabHandle;
  }

  /** Detach from the current collaboration session. */
  detachCollab(): void {
    if (!this.collabHandle) return;
    if (this.threadsObserver) {
      getYDocFields(this.collabHandle.ydoc).threads.unobserve(this.threadsObserver);
      this.threadsObserver = null;
    }
    this.collabHandle.detach();
    this.collabHandle = null;
    this.emit('documentChange', { action: 'detachCollab' });
  }
}

