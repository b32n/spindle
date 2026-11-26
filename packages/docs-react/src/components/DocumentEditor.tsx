/**
 * DocumentEditor - Main document editor component
 * 
 * This component provides a Google Docs-like editing experience using a single
 * ProseMirror instance architecture with a true layout engine.
 * 
 * Architecture:
 * - Single hidden ProseMirror instance handles all editing
 * - Separate visual rendering layer displays paginated content
 * - InputBridge forwards keyboard events from visible surface to hidden editor
 * - SelectionOverlay renders caret and selection as a separate layer
 */

import { memo, useState, useCallback, useRef, useEffect, useMemo } from 'react';
import type { PageConfig, PageMargins, Block, HeaderFooterContent, RevisionAuthor, ReviewChange } from '@b32nio/spindle-docs-core';
import { activeMarksPluginKey, createCommands, docsSchema, documentRevisionAuthors, listChanges, resolveAll, resolveChange, revisionAuthors, sameRevisionAuthors } from '@b32nio/spindle-docs-core';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { useDocument, useSections } from '../context/DocumentContext';
import { Toolbar } from './Toolbar';
import { CommentsSidebar } from './CommentsSidebar';
import { RevisionViewBanner, isEditLocked } from './RevisionViewBanner';
import { RevisionLegend } from './RevisionLegend';
import { ReviewPanel } from './ReviewPanel';
import { Ruler } from './Ruler';
import { VerticalRuler } from './VerticalRuler';
import { PageSetupModal } from './PageSetupModal';
import { TrueLayoutEditor, CellSelection, ActivePageInfo, TrueLayoutEditorHandle, type RevisionView, type TrueLayoutSection } from '../core';
import type { ActiveMarks } from './ProseMirrorEditor';
import type { ImportLossReport } from '@b32nio/spindle-shared';
import { downloadFile, useImportConfirm, ImageHooksContext, type FontSource, type ImageHooks, type ImportConfirmation } from '@b32nio/spindle-shared/react';
import {
  DOCS_EXPORT_FORMATS,
  exportDocumentFile,
  importDocumentFile,
  inspectDocumentFile,
  type DocsExportFormat,
} from '../file-io';

/** How long typing pauses before the authors of the tracked changes are collected again. */
const AUTHORS_DELAY_MS = 150;

interface DocumentEditorProps {
  width?: number;
  height?: number;
  showToolbar?: boolean;
  showRuler?: boolean;
  /** Disables editing: content, margins/header/footer, and comments become view-only. */
  readOnly?: boolean;
  /**
   * Show Open (.docx, read from a local file input) and Download (.docx) in
   * the toolbar. Opening a file replaces the whole document. Default: false.
   * Requires the toolbar (`showToolbar`) and is hidden in read-only mode;
   * use `exportDocumentFile` for a host-side download button.
   */
  fileActions?: boolean;
  /** Formats offered by Download. Default: every format in `DOCS_EXPORT_FORMATS`. */
  exportFormats?: readonly DocsExportFormat[];
  /** Base name for downloaded files (no extension). Default: the document title. */
  fileName?: string;
  /** Called after a file was opened; `warnings` lists features the importer had to drop. */
  onFileOpened?: (info: { fileName: string; warnings: string[]; loss: ImportLossReport }) => void;
  /** Called when opening or downloading fails (unsupported type, corrupt file, collab attached…). */
  onFileError?: (error: Error, action: 'open' | 'download') => void;
  /**
   * Where web fonts offered by the font menu are loaded from. Default: the
   * Spindle font mirror (`https://fonts.b32n.io/css2`), falling back to system
   * fonts if unreachable. `{ kind: 'local' }` makes no font requests.
   */
  fontSource?: FontSource;
  /**
   * What Open does when the file contains things that cannot be imported (see `inspectDocumentFile`):
   * `'dialog'` (default) asks "Import anyway / Cancel", `'never'` imports without asking, or pass
   * `(report, file) => boolean | Promise<boolean>` to ask with your own UI. Nothing is asked for a
   * file that loses nothing.
   */
  importConfirmation?: ImportConfirmation;
  /**
   * Fetches an image for "insert by URL" (return a Blob or a `data:image/…` URL).
   * Without it the editor never fetches image URLs: the By-URL tab is hidden
   * and only local files can be inserted.
   */
  resolveImage?: ImageHooks['resolveImage'];
  /** Stores a local image file and returns the URL to put in the document (default: embed as a data URL). */
  uploadImage?: ImageHooks['uploadImage'];
  /** Largest accepted image in bytes (default 10 MB). */
  maxImageBytes?: number;
  /**
   * How tracked changes show to begin with: `markup` (default, every change marked), `final` (as if all were accepted) or
   * `original` (as if all were rejected). The review panel changes it; the document is not changed by it.
   */
  revisionView?: RevisionView;
}

const defaultActiveMarks: ActiveMarks = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  superscript: false,
  subscript: false,
  link: null,
  textStyle: null,
  blockType: 'paragraph',
  headingLevel: 1,
  listType: null,
  alignment: 'left',
};

/**
 * Main document editor component.
 * 
 * Uses a single ProseMirror instance for the entire document:
 * - Better performance for large documents
 * - Unified undo/redo history
 * - Simpler toolbar integration
 * - Proper selection handling across blocks
 */
export const DocumentEditor = memo(function DocumentEditor({
  width,
  height,
  showToolbar = true,
  showRuler = true,
  readOnly: readOnlyProp = false,
  fileActions = false,
  exportFormats,
  fileName,
  onFileOpened,
  onFileError,
  fontSource,
  importConfirmation,
  resolveImage,
  uploadImage,
  maxImageBytes,
  revisionView: initialRevisionView = 'markup',
}: DocumentEditorProps) {
  const { document: docModel, zoom, setZoom, currentUser } = useDocument();
  const sections = useSections();

  const [showPageSetup, setShowPageSetup] = useState(false);
  const [editorView, setEditorView] = useState<EditorView | null>(null);
  const [activeMarks, setActiveMarks] = useState<ActiveMarks>(defaultActiveMarks);
  const [selectedCell, setSelectedCell] = useState<CellSelection | null>(null);
  const [activePageInfo, setActivePageInfo] = useState<ActivePageInfo | null>(null);
  const [scrollContainer, setScrollContainer] = useState<HTMLElement | null>(null);
  const [showComments, setShowComments] = useState(false);
  const [authors, setAuthors] = useState<RevisionAuthor[]>([]);
  const [changes, setChanges] = useState<ReviewChange[]>([]);
  const [palette, setPalette] = useState<string[]>([]);
  const [showReview, setShowReview] = useState(false);
  const showReviewRef = useRef(false);
  const [revisionView, setRevisionView] = useState<RevisionView>(initialRevisionView);
  // The Final and Original views hide text, so they cannot be edited (the banner offers the way back to All markup).
  const readOnly = isEditLocked(readOnlyProp, revisionView);
  const [pendingComment, setPendingComment] = useState<{ from: number; to: number; quote: string } | null>(null);
  const [activeThreadId, setActiveThreadId] = useState<string | null>(null);
  
  // Bumped after a file is opened: the editor only seeds from the model on
  // mount, so remounting it is how a replaced document reaches the screen.
  const [loadVersion, setLoadVersion] = useState(0);

  const containerRef = useRef<HTMLDivElement>(null);
  const editorRef = useRef<TrueLayoutEditorHandle>(null);
  
  // The section the active page belongs to (the first one until the editor reports a page): rulers and page setup act on it.
  const activeSection = sections.find((s) => s.id === activePageInfo?.sectionId) ?? sections[0];
  const currentPageConfig = activeSection?.pageConfig || docModel.getDefaultPageConfig();
  const documentTitle = docModel.getTitle();
  
  // ---- File open / download (toolbar, opt-in via `fileActions`) ----
  const downloadFormats = useMemo(
    () => (exportFormats ? DOCS_EXPORT_FORMATS.filter((f) => exportFormats.includes(f.id)) : DOCS_EXPORT_FORMATS),
    [exportFormats]
  );

  const { confirm: confirmImport, dialog: importDialog } = useImportConfirm(importConfirmation);

  const handleOpenFile = useCallback(async (file: File) => {
    try {
      // Ask before replacing the document if the file has things we can't import.
      if (!(await confirmImport(file, await inspectDocumentFile(file)))) return;
      const { warnings, loss } = await importDocumentFile(docModel, file);
      setPendingComment(null);
      setActiveThreadId(null);
      setLoadVersion((v) => v + 1);
      onFileOpened?.({ fileName: file.name, warnings, loss });
    } catch (err) {
      onFileError?.(err instanceof Error ? err : new Error(String(err)), 'open');
    }
  }, [docModel, onFileOpened, onFileError, confirmImport]);

  const handleDownload = useCallback(async (format: DocsExportFormat) => {
    try {
      const out = await exportDocumentFile(docModel, format, fileName);
      downloadFile(out.data, out.fileName, out.mimeType);
    } catch (err) {
      onFileError?.(err instanceof Error ? err : new Error(String(err)), 'download');
    }
  }, [docModel, fileName, onFileError]);

  // Handle header changes - update document model
  const handleHeaderChange = useCallback((content: HeaderFooterContent, sectionId?: string) => {
    if (readOnly) return;
    const target = sections.find((s) => s.id === sectionId) ?? sections[0];
    if (target) {
      docModel.setSectionHeader(target.id, content);
      docModel.recordHistory('Update header');
    }
  }, [docModel, sections, readOnly]);

  // Handle footer changes - update document model
  const handleFooterChange = useCallback((content: HeaderFooterContent, sectionId?: string) => {
    if (readOnly) return;
    const target = sections.find((s) => s.id === sectionId) ?? sections[0];
    if (target) {
      docModel.setSectionFooter(target.id, content);
      docModel.recordHistory('Update footer');
    }
  }, [docModel, sections, readOnly]);

  const handleInsertSectionBreak = useCallback((type: 'nextPage' | 'continuous') => {
    if (readOnly || !editorView) return;
    const id = `section_${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
    pendingBreaks.current.set(id, type);
    createCommands(docsSchema).insertSectionBreak(id)(editorView.state, editorView.dispatch);
    editorRef.current?.focus();
  }, [readOnly, editorView]);

  const handleInsertNote = useCallback((kind: 'footnote' | 'endnote') => {
    if (readOnly) return;
    editorRef.current?.insertNote(kind);
  }, [readOnly]);

  const handleTableOfContents = useCallback((action: 'insert' | 'update' | 'remove') => {
    if (readOnly) return;
    editorRef.current?.tableOfContents(action);
  }, [readOnly]);

  const handleResolve = useCallback((action: 'accept' | 'reject', target: { change: ReviewChange } | { author?: string }) => {
    // An explicit command by the reader: allowed in the Final and Original views too, not when the host made the editor read-only.
    if (readOnlyProp || !editorView) return;
    const command = 'change' in target ? resolveChange(target.change.id, action) : resolveAll(action, target.author);
    command(editorView.state, editorView.dispatch);
    // The list is refreshed now (not after the typing pause), so a second click straight after the first never uses an id that is gone.
    if (showReviewRef.current) setChanges(listChanges(editorView.state.doc));
  }, [readOnlyProp, editorView]);

  const handleGoToChange = useCallback((change: ReviewChange) => {
    editorRef.current?.goTo?.(change.from);
  }, []);

  const handleRemoveSectionBreak = useCallback(() => {
    if (readOnly || !editorView) return;
    createCommands(docsSchema).removeSectionBreak()(editorView.state, editorView.dispatch);
    editorRef.current?.focus();
  }, [readOnly, editorView]);

  const handleHeaderFooterLinkChange = useCallback((type: 'header' | 'footer', sectionId: string, linked: boolean, content: HeaderFooterContent) => {
    if (readOnly) return;
    const value = linked ? undefined : content;
    if (type === 'header') docModel.setSectionHeader(sectionId, value);
    else docModel.setSectionFooter(sectionId, value);
    docModel.recordHistory(linked ? 'Link to previous' : 'Unlink from previous');
  }, [docModel, readOnly]);

  // Handle margins change from ruler
  const handleMarginsChange = useCallback((margins: PageMargins) => {
    if (readOnly) return;
    if (activeSection) {
      docModel.setSectionPageConfig(activeSection.id, {
        ...activeSection.pageConfig,
        margins,
      });
      docModel.recordHistory('Update margins');
    }
  }, [docModel, activeSection, readOnly]);

  // Handle page setup confirmation
  const handlePageSetupConfirm = useCallback((config: PageConfig, scope: 'section' | 'document') => {
    if (readOnly) return;
    if (scope === 'document') {
      // Every section gets the size, orientation and margins (each keeps its own columns).
      for (const s of sections) {
        docModel.setSectionPageConfig(s.id, { ...config, columns: s.pageConfig.columns });
      }
      docModel.recordHistory('Update page setup');
    } else if (activeSection) {
      docModel.setSectionPageConfig(activeSection.id, config);
      docModel.recordHistory('Update page setup');
    }
    docModel.setDefaultPageConfig(config);
  }, [docModel, activeSection, sections, readOnly]);

  // Handle document changes from the editor
  // Break types chosen in the toolbar, by the id of the section the new break ends (read when the model learns of it).
  const pendingBreaks = useRef(new Map<string, 'nextPage' | 'continuous'>());

  const handleSectionsChange = useCallback((next: Array<{ id: string; blocks: Block[] }>) => {
    if (readOnly) return;
    const ids = new Set(next.map((s) => s.id));
    // A section the model does not know yet was split off by a new section break (here or in a peer's editor).
    next.forEach((s, i) => {
      if (docModel.getSection(s.id)) return;
      const later = next[i + 1];
      if (later && docModel.getSection(later.id)) {
        docModel.splitSection(later.id, s.id, pendingBreaks.current.get(s.id));
        pendingBreaks.current.delete(s.id);
      }
    });
    for (const s of next) docModel.setSectionBlocks(s.id, s.blocks);
    // A section whose break was deleted has merged into the one after it.
    for (const s of [...docModel.getSections()]) if (!ids.has(s.id)) docModel.deleteSection(s.id);
    docModel.recordHistory('Edit content');
  }, [docModel, readOnly]);
  
  // Handle selection changes from the editor
  const authorsDoc = useRef<EditorState['doc'] | null>(null);
  const authorsTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  useEffect(() => () => clearTimeout(authorsTimer.current), []);
  const updateAuthors = useCallback((state: EditorState) => {
    if (authorsDoc.current === state.doc) return;
    authorsDoc.current = state.doc;
    // The walk over the whole document waits until typing pauses, so a long document is not scanned on every key.
    clearTimeout(authorsTimer.current);
    authorsTimer.current = setTimeout(() => {
      const nextChanges = listChanges(state.doc);
      // The colours are shared out among every author of the document (a move's other end can have another author than the change lists).
      const everyone = documentRevisionAuthors(state.doc);
      const next = revisionAuthors(nextChanges, everyone);
      setPalette((prev) => (prev.length === everyone.length && prev.every((name, i) => name === everyone[i]) ? prev : everyone));
      // Keep the old list (and skip a render) when nothing in it changed.
      setAuthors((prev) => (sameRevisionAuthors(prev, next) ? prev : next));
      // The list of changes (new ids with every edit) is only kept while the panel shows it, so typing does not render the editor.
      if (showReviewRef.current) setChanges(nextChanges);
    }, AUTHORS_DELAY_MS);
  }, []);

  // Opening the panel lists the changes as they are now; closing it stops keeping the list.
  useEffect(() => {
    showReviewRef.current = showReview;
    if (showReview && editorView) setChanges(listChanges(editorView.state.doc));
  }, [showReview, editorView]);

  const handleSelectionChange = useCallback((state: EditorState) => {
    const marks = activeMarksPluginKey.getState(state) as ActiveMarks | undefined;
    if (marks) {
      setActiveMarks(marks);
    }
    updateAuthors(state);
  }, [updateAuthors]);
  
  // Handle editor ready
  const handleEditorReady = useCallback((view: EditorView) => {
    setEditorView(view);
    updateAuthors(view.state);
    
    // Get initial active marks
    const marks = activeMarksPluginKey.getState(view.state) as ActiveMarks | undefined;
    if (marks) {
      setActiveMarks(marks);
    }
    
    // Get scroll container for vertical ruler
    setTimeout(() => {
      if (editorRef.current) {
        setScrollContainer(editorRef.current.getScrollContainer());
      }
    }, 0);
  }, [updateAuthors]);
  
  // Handle cell selection changes
  const handleCellSelectionChange = useCallback((selection: CellSelection | null) => {
    setSelectedCell(selection);
  }, []);
  
  // Handle active page changes
  const handleActivePageChange = useCallback((info: ActivePageInfo) => {
    setActivePageInfo(info);
  }, []);

  // --- Comments ------------------------------------------------------------

  // Begin a comment on the current selection: capture the range + quoted text
  // and open the sidebar with a composer. The thread is created on submit.
  const startComment = useCallback(() => {
    if (!editorView) return;
    const { from, to } = editorView.state.selection;
    if (from === to) return;
    const quote = editorView.state.doc.textBetween(from, to, ' ');
    setPendingComment({ from, to, quote });
    setActiveThreadId(null);
    setShowComments(true);
    setShowReview(false);
  }, [editorView]);

  // Create the thread and apply the comment mark over the captured range.
  const createComment = useCallback((body: string, mentions: string[]) => {
    if (!editorView || !pendingComment) return;
    const thread = docModel.getComments().addThread(
      { quote: pendingComment.quote },
      body,
      currentUser,
      mentions,
    );
    const commentMark = editorView.state.schema.marks.comment;
    editorView.dispatch(
      editorView.state.tr.addMark(
        pendingComment.from,
        pendingComment.to,
        commentMark.create({ threadId: thread.id }),
      ),
    );
    setPendingComment(null);
    setActiveThreadId(thread.id);
  }, [editorView, pendingComment, docModel, currentUser]);

  const replyToThread = useCallback((threadId: string, body: string, mentions: string[]) => {
    docModel.getComments().addReply(threadId, body, currentUser, mentions);
  }, [docModel, currentUser]);

  const resolveThread = useCallback((threadId: string) => {
    docModel.getComments().resolveThread(threadId, currentUser);
  }, [docModel, currentUser]);

  const reopenThread = useCallback((threadId: string) => {
    docModel.getComments().reopenThread(threadId, currentUser);
  }, [docModel, currentUser]);

  // Strip a thread's comment mark from the document body.
  const stripCommentMark = useCallback((threadId: string) => {
    if (!editorView) return;
    createCommands(docsSchema).removeComment(threadId)(editorView.state, editorView.dispatch);
  }, [editorView]);

  const deleteThread = useCallback((threadId: string) => {
    docModel.getComments().deleteThread(threadId);
    stripCommentMark(threadId);
  }, [docModel, stripCommentMark]);

  const deleteComment = useCallback((threadId: string, commentId: string) => {
    const store = docModel.getComments();
    store.deleteComment(threadId, commentId);
    // Deleting the last comment removes the thread — drop its highlight too.
    if (!store.getThread(threadId)) {
      stripCommentMark(threadId);
    }
  }, [docModel, stripCommentMark]);

  // Clicking a comment highlight opens the sidebar focused on that thread.
  useEffect(() => {
    const root = containerRef.current;
    if (!root) return;
    const onClick = (e: MouseEvent) => {
      const el = (e.target as HTMLElement).closest('[data-comment-thread]');
      const threadId = el?.getAttribute('data-comment-thread');
      if (threadId) {
        setActiveThreadId(threadId);
        setPendingComment(null);
        setShowComments(true);
        setShowReview(false);
      }
    };
    root.addEventListener('click', onClick);
    return () => root.removeEventListener('click', onClick);
  }, []);
  
  // Keyboard shortcuts for document-level undo/redo
  useEffect(() => {
    if (readOnly) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      // While attached, the editor's Yjs undo manager handles Mod-z/y (local edits only).
      if (docModel.getCollabHandle()) return;
      if ((e.metaKey || e.ctrlKey) && e.key === 'z') {
        e.preventDefault();
        if (e.shiftKey) {
          docModel.redo();
        } else {
          docModel.undo();
        }
      }
      if ((e.metaKey || e.ctrlKey) && e.key === 'y') {
        e.preventDefault();
        docModel.redo();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [docModel, readOnly]);
  
  // Convert a PageConfig to the format TrueLayoutEditor expects: for landscape, swap width and height.
  const toPresentation = (config: PageConfig) => ({
    width: config.orientation === 'landscape' ? config.size.h : config.size.w,
    height: config.orientation === 'landscape' ? config.size.w : config.size.h,
    margins: config.margins,
  });
  const layoutSections: TrueLayoutSection[] = sections.map((s) => ({
    id: s.id,
    pageConfig: toPresentation(s.pageConfig),
    blocks: s.blocks,
    header: s.header,
    footer: s.footer,
    ...(s.breakType && { breakType: s.breakType }),
    ...(s.pageNumbering && { pageNumbering: s.pageNumbering }),
  }));
  const presentationPageConfig = layoutSections[0]?.pageConfig ?? toPresentation(currentPageConfig);

  // Fit the page to the viewport width on narrow screens. The page is a fixed
  // pixel width (Letter ≈ 816px); on a phone it's wider than the screen, and
  // because the viewport centres it, the leading edge becomes unreachable —
  // content spills off both sides. When it can't fit at 100% we scale zoom down
  // to fit; when it fits (desktop) we leave the user's chosen zoom alone.
  const pageWidthPx = Math.max(presentationPageConfig.width, ...layoutSections.map((s) => s.pageConfig.width));
  // True while we're actively overriding zoom to fit a narrow screen — so when
  // the screen widens again we can restore 100%, but a manual zoom on a screen
  // that always fit (desktop) is never touched.
  const autoFitRef = useRef(false);
  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    // Viewport has 24px padding each side; leave a little slack for the ruler
    // gutter and scrollbar so the page never kisses the edge.
    const CHROME = 24 * 2 + 16;
    const fit = () => {
      const avail = el.clientWidth - CHROME;
      if (avail <= 0) return;
      if (avail >= pageWidthPx) {
        if (autoFitRef.current) { autoFitRef.current = false; setZoom(100); }
        return; // fits at 100% → leave the user's zoom alone
      }
      autoFitRef.current = true;
      setZoom(Math.max(25, Math.floor((avail / pageWidthPx) * 100)));
    };
    fit();
    const ro = new ResizeObserver(fit);
    ro.observe(el);
    return () => ro.disconnect();
  }, [pageWidthPx, setZoom]);

  const imageHooks = useMemo<ImageHooks>(
    () => ({ resolveImage, uploadImage, maxImageBytes }),
    [resolveImage, uploadImage, maxImageBytes]
  );

  return (
    <ImageHooksContext.Provider value={imageHooks}>
    <div
      ref={containerRef}
      className="document-editor"
      style={{
        width: width || '100%',
        height: height || '100%',
        display: 'flex',
        flexDirection: 'column',
        backgroundColor: '#f8f9fa',
        overflow: 'hidden',
      }}
    >
      {/* Floating Toolbar — hidden in read-only mode, since every action it exposes mutates the document */}
      {showToolbar && !readOnly && (
        <div style={{ 
          position: 'relative', 
          zIndex: 200, 
          flexShrink: 0,
          display: 'flex',
          justifyContent: 'center',
          padding: '8px 16px',
          background: 'linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%)',
        }}>
          <Toolbar
            fontSource={fontSource}
            onPageSetup={() => setShowPageSetup(true)}
            editorView={editorView}
            activeMarks={activeMarks}
            selectedCell={selectedCell}
            onAddComment={startComment}
            onInsertNote={handleInsertNote}
            onTableOfContents={handleTableOfContents}
            onToggleReview={() => {
              setShowReview((v) => !v);
              setShowComments(false);
            }}
            reviewOpen={showReview}
            onInsertSectionBreak={handleInsertSectionBreak}
            onRemoveSectionBreak={handleRemoveSectionBreak}
            canRemoveSectionBreak={sections.length > 1}
            onOpenFile={fileActions ? handleOpenFile : undefined}
            onDownload={fileActions ? handleDownload : undefined}
            downloadFormats={downloadFormats}
          />
        </div>
      )}
      
      {/* Ruler container — hidden in read-only mode; dragging its margin guides mutates page config */}
      {showRuler && !readOnly && (
        <div
          style={{
            display: 'flex',
            justifyContent: 'center',
            backgroundColor: '#ffffff',
            borderBottom: '1px solid #e8eaed',
            padding: '8px 0',
          }}
        >
          <Ruler
            pageConfig={currentPageConfig}
            onMarginsChange={handleMarginsChange}
          />
        </div>
      )}
      
      <RevisionViewBanner view={revisionView} onSwitchToMarkup={() => setRevisionView('markup')} />
      <RevisionLegend authors={authors} />

      {/* Editor area with vertical ruler */}
      <div style={{ flex: 1, display: 'flex', position: 'relative', overflow: 'hidden' }}>
        {/* Vertical Ruler - positioned to the left */}
        {showRuler && !readOnly && activePageInfo && scrollContainer && (
          <div
            style={{
              position: 'relative',
              width: 24 * (zoom / 100),
              flexShrink: 0,
              backgroundColor: '#f8f9fa',
              borderRight: '1px solid #e8eaed',
              overflow: 'hidden',
            }}
          >
            <VerticalRuler
              pageConfig={currentPageConfig}
              onMarginsChange={handleMarginsChange}
              pageIndex={activePageInfo.pageIndex}
              pageHeight={activePageInfo.pageHeight}
              pageGap={24}
              scrollContainerRef={{ current: scrollContainer }}
            />
          </div>
        )}
        
        {/* Editor - using TrueLayoutEditor (line-level pagination) */}
        <TrueLayoutEditor
          key={loadVersion}
          ref={editorRef}
          initialBlocks={[]}
          sections={layoutSections}
          onSectionsChange={handleSectionsChange}
          pageConfig={presentationPageConfig}
          zoom={zoom / 100}
          editable={!readOnly}
          onSelectionChange={handleSelectionChange}
          onCellSelectionChange={handleCellSelectionChange}
          onActivePageChange={handleActivePageChange}
          onReady={handleEditorReady}
          pageGap={24}
          minLinesAtBreak={2}
          textStylePool={docModel.getTextStylePool()}
          documentTitle={documentTitle}
          onHeaderChange={handleHeaderChange}
          onFooterChange={handleFooterChange}
          onHeaderFooterLinkChange={handleHeaderFooterLinkChange}
          collabHandle={docModel.getCollabHandle()}
          revisionView={revisionView}
        />

        {/* Review panel — tracked changes, in the place of the comments sidebar */}
        {showReview && (
          <div
            style={{
              position: 'absolute',
              top: 0,
              right: 0,
              bottom: 0,
              width: 320,
              zIndex: 30,
              boxShadow: '-8px 0 24px -12px rgba(15, 23, 42, 0.15)',
            }}
          >
            <ReviewPanel
              changes={changes}
              palette={palette}
              view={revisionView}
              onViewChange={setRevisionView}
              onResolve={readOnlyProp ? undefined : handleResolve}
              onGoTo={handleGoToChange}
              onClose={() => setShowReview(false)}
            />
          </div>
        )}

        {/* Comments sidebar — a floating overlay so the page + ruler stay put */}
        {!readOnly && showComments && (
          <div
            style={{
              position: 'absolute',
              top: 0,
              right: 0,
              bottom: 0,
              width: 320,
              zIndex: 30,
              boxShadow: '-8px 0 24px -12px rgba(15, 23, 42, 0.15)',
            }}
          >
            <CommentsSidebar
              threads={docModel.getComments().getThreads()}
              activeThreadId={activeThreadId}
              pendingQuote={pendingComment?.quote ?? null}
              onClose={() => {
                setShowComments(false);
                setPendingComment(null);
              }}
              onCreate={createComment}
              onCancelPending={() => setPendingComment(null)}
              onReply={replyToThread}
              onResolve={resolveThread}
              onReopen={reopenThread}
              onDeleteThread={deleteThread}
              onDeleteComment={deleteComment}
            />
          </div>
        )}
      </div>

      {/* Page Setup Modal */}
      <PageSetupModal
        isOpen={showPageSetup}
        pageConfig={currentPageConfig}
        onClose={() => setShowPageSetup(false)}
        onConfirm={handlePageSetupConfirm}
        sectionCount={sections.length}
      />
      {importDialog}
    </div>
    </ImageHooksContext.Provider>
  );
});
