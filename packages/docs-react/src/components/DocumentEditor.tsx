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
import type { PageConfig, PageMargins, Block, HeaderFooterContent } from '@b32nio/spindle-docs-core';
import { activeMarksPluginKey, createCommands, docsSchema } from '@b32nio/spindle-docs-core';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { useDocument, useSections } from '../context/DocumentContext';
import { Toolbar } from './Toolbar';
import { CommentsSidebar } from './CommentsSidebar';
import { Ruler } from './Ruler';
import { VerticalRuler } from './VerticalRuler';
import { PageSetupModal } from './PageSetupModal';
import { TrueLayoutEditor, CellSelection, ActivePageInfo, TrueLayoutEditorHandle } from '../core';
import type { ActiveMarks } from './ProseMirrorEditor';

interface DocumentEditorProps {
  width?: number;
  height?: number;
  showToolbar?: boolean;
  showRuler?: boolean;
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
  const [pendingComment, setPendingComment] = useState<{ from: number; to: number; quote: string } | null>(null);
  const [activeThreadId, setActiveThreadId] = useState<string | null>(null);
  
  const containerRef = useRef<HTMLDivElement>(null);
  const editorRef = useRef<TrueLayoutEditorHandle>(null);
  
  // Get the first section's page config for the ruler
  const currentPageConfig = sections[0]?.pageConfig || docModel.getDefaultPageConfig();
  const initialBlocks = sections[0]?.blocks || [];
  
  // Get header/footer from the document model (first section)
  const header = sections[0]?.header;
  const footer = sections[0]?.footer;
  const documentTitle = docModel.getTitle();
  
  // Handle header changes - update document model
  const handleHeaderChange = useCallback((content: HeaderFooterContent) => {
    if (sections[0]) {
      docModel.setSectionHeader(sections[0].id, content);
      docModel.recordHistory('Update header');
    }
  }, [docModel, sections]);
  
  // Handle footer changes - update document model
  const handleFooterChange = useCallback((content: HeaderFooterContent) => {
    if (sections[0]) {
      docModel.setSectionFooter(sections[0].id, content);
      docModel.recordHistory('Update footer');
    }
  }, [docModel, sections]);
  
  // Handle margins change from ruler
  const handleMarginsChange = useCallback((margins: PageMargins) => {
    if (sections[0]) {
      docModel.setSectionPageConfig(sections[0].id, {
        ...sections[0].pageConfig,
        margins,
      });
      docModel.recordHistory('Update margins');
    }
  }, [docModel, sections]);
  
  // Handle page setup confirmation
  const handlePageSetupConfirm = useCallback((config: PageConfig) => {
    if (sections[0]) {
      docModel.setSectionPageConfig(sections[0].id, config);
      docModel.recordHistory('Update page setup');
    }
    docModel.setDefaultPageConfig(config);
  }, [docModel, sections]);
  
  // Handle document changes from the editor
  const handleDocChange = useCallback((blocks: Block[]) => {
    if (sections[0]) {
      docModel.setSectionBlocks(sections[0].id, blocks);
      docModel.recordHistory('Edit content');
    }
  }, [docModel, sections]);
  
  // Handle selection changes from the editor
  const handleSelectionChange = useCallback((state: EditorState) => {
    const marks = activeMarksPluginKey.getState(state) as ActiveMarks | undefined;
    if (marks) {
      setActiveMarks(marks);
    }
  }, []);
  
  // Handle editor ready
  const handleEditorReady = useCallback((view: EditorView) => {
    setEditorView(view);
    
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
  }, []);
  
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
      }
    };
    root.addEventListener('click', onClick);
    return () => root.removeEventListener('click', onClick);
  }, []);
  
  // Keyboard shortcuts for document-level undo/redo
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
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
  }, [docModel]);
  
  // Convert PageConfig to the format expected by TrueLayoutEditor
  // When orientation is landscape, swap width and height
  // Memoize to prevent unnecessary re-renders of TrueLayoutEditor
  // Note: We list individual margin properties to ensure proper reactivity
  const presentationPageConfig = useMemo(() => ({
    width: currentPageConfig.orientation === 'landscape' 
      ? currentPageConfig.size.h 
      : currentPageConfig.size.w,
    height: currentPageConfig.orientation === 'landscape' 
      ? currentPageConfig.size.w 
      : currentPageConfig.size.h,
    margins: currentPageConfig.margins,
  }), [
    currentPageConfig.orientation,
    currentPageConfig.size.w,
    currentPageConfig.size.h,
    currentPageConfig.margins.top,
    currentPageConfig.margins.bottom,
    currentPageConfig.margins.left,
    currentPageConfig.margins.right,
    currentPageConfig.margins.header,
    currentPageConfig.margins.footer,
  ]);
  
  // Fit the page to the viewport width on narrow screens. The page is a fixed
  // pixel width (Letter ≈ 816px); on a phone it's wider than the screen, and
  // because the viewport centres it, the leading edge becomes unreachable —
  // content spills off both sides. When it can't fit at 100% we scale zoom down
  // to fit; when it fits (desktop) we leave the user's chosen zoom alone.
  const pageWidthPx = presentationPageConfig.width;
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

  return (
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
      {/* Floating Toolbar */}
      {showToolbar && (
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
            onPageSetup={() => setShowPageSetup(true)}
            editorView={editorView}
            activeMarks={activeMarks}
            selectedCell={selectedCell}
            onAddComment={startComment}
          />
        </div>
      )}
      
      {/* Ruler container */}
      {showRuler && (
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
      
      {/* Editor area with vertical ruler */}
      <div style={{ flex: 1, display: 'flex', position: 'relative', overflow: 'hidden' }}>
        {/* Vertical Ruler - positioned to the left */}
        {showRuler && activePageInfo && scrollContainer && (
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
          ref={editorRef}
          initialBlocks={initialBlocks}
          pageConfig={presentationPageConfig}
          zoom={zoom / 100}
          editable={true}
          onDocChange={handleDocChange}
          onSelectionChange={handleSelectionChange}
          onCellSelectionChange={handleCellSelectionChange}
          onActivePageChange={handleActivePageChange}
          onReady={handleEditorReady}
          pageGap={24}
          minLinesAtBreak={2}
          textStylePool={docModel.getTextStylePool()}
          header={header}
          footer={footer}
          documentTitle={documentTitle}
          onHeaderChange={handleHeaderChange}
          onFooterChange={handleFooterChange}
          collabHandle={docModel.getCollabHandle()}
        />

        {/* Comments sidebar — a floating overlay so the page + ruler stay put */}
        {showComments && (
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
      />
    </div>
  );
});
