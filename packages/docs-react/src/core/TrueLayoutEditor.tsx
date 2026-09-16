/**
 * TrueLayoutEditor - Main orchestrator for the true layout engine
 * 
 * This component brings together all the pieces of the true layout engine:
 * 1. Hidden ProseMirror editor for editing
 * 2. FlowBlocks converter
 * 3. DOM Measurer
 * 4. Layout Engine
 * 5. DOM Painter
 * 6. Input Bridge
 * 7. Selection Overlay
 * 
 * The result is a document editor with true line-level pagination,
 * where each page renders only its assigned content.
 */

import {
  useRef,
  useState,
  useCallback,
  useEffect,
  forwardRef,
  useImperativeHandle,
} from 'react';
import { EditorState, Transaction } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { Node as PmNode } from 'prosemirror-model';
import { docsSchema, createPlugins, blocksToPmDoc, proseMirrorToDocument, Block, DEFAULT_PAGE_CONFIG } from '@b32nio/spindle-docs-core';
import type { CollabHandle } from '@b32nio/spindle-docs-core/collab';
import { ySyncPlugin, yCursorPlugin } from 'y-prosemirror';
import { ensureCollabCursorStyles } from './collab-cursor-styles';

import { FlowBlock } from './flow-blocks';
import { proseMirrorToFlowBlocks, createBlockPositionMap } from './pm-to-blocks';
import { DomMeasurer, Measure } from './measurer';
import { computeTrueLayout, DocumentLayout, PageConfig } from './true-layout-engine';
import { DomPainter, HeaderFooterContent } from './dom-painter';
import { InputBridge, createInputBridge, CellSelection } from './input-bridge';
import { SelectionOverlayManager, getSelectionOverlayStyles } from './selection-overlay';
import { RemoteCursorOverlay } from './remote-cursor-overlay';
import { TableInteractionManager, createTableInteractionManager } from './table-interactions';
import { ImageInteractionManager, createImageInteractionManager } from './image-interactions';
import { HeaderFooterEditor } from '../components/HeaderFooterEditor';

// ============================================================================
// Types
// ============================================================================

export interface ActivePageInfo {
  /** Index of the active page (0-based) */
  pageIndex: number;
  /** Total number of pages */
  totalPages: number;
  /** Height of the active page */
  pageHeight: number;
}

export interface TrueLayoutEditorProps {
  /** Initial blocks to render */
  initialBlocks: Block[];
  /** Page configuration */
  pageConfig: PageConfig;
  /** Zoom level (1 = 100%) */
  zoom?: number;
  /** Whether the editor is editable */
  editable?: boolean;
  /** Called when the document changes */
  onDocChange?: (blocks: Block[]) => void;
  /** Called when selection changes */
  onSelectionChange?: (state: EditorState) => void;
  /** Called when cell selection changes (for table cell formatting) */
  onCellSelectionChange?: (selection: CellSelection | null) => void;
  /** Called when the active page changes (for vertical ruler positioning) */
  onActivePageChange?: (info: ActivePageInfo) => void;
  /** Called when the editor is ready */
  onReady?: (view: EditorView) => void;
  /** Container width */
  width?: number | string;
  /** Container height */
  height?: number | string;
  /** Gap between pages in pixels */
  pageGap?: number;
  /** Minimum lines to keep together at page break */
  minLinesAtBreak?: number;
  /** Text style pool for efficient style storage (recommended for large documents) */
  textStylePool?: import('@b32nio/spindle-docs-core').TextStylePool;
  /** Header content configuration */
  header?: HeaderFooterContent;
  /** Footer content configuration */
  footer?: HeaderFooterContent;
  /** Document title (used for dynamic field resolution in headers/footers) */
  documentTitle?: string;
  /** Called when header content changes */
  onHeaderChange?: (content: HeaderFooterContent) => void;
  /** Called when footer content changes */
  onFooterChange?: (content: HeaderFooterContent) => void;
  /** Whether headers/footers are editable (default: true when editable is true) */
  headerFooterEditable?: boolean;
  /**
   * If present, the editor binds the hidden ProseMirror state to the Y.Doc's
   * Y.XmlFragment via ySyncPlugin. Local edits propagate to peers and remote
   * edits land in the editor through normal PM transactions, so the layout
   * engine sees them just like local typing. initialBlocks is ignored when
   * collabHandle is present — the Y.Doc is the source of truth.
   */
  collabHandle?: CollabHandle | null;
}

export interface TrueLayoutEditorHandle {
  /** Get the ProseMirror EditorView */
  getView: () => EditorView | null;
  /** Get the current EditorState */
  getState: () => EditorState | null;
  /** Focus the editor */
  focus: () => void;
  /** Check if the editor has focus */
  hasFocus: () => boolean;
  /** Get the current layout */
  getLayout: () => DocumentLayout | null;
  /** Force a re-layout */
  reflow: () => void;
  /** Get the scroll container element (for syncing vertical ruler) */
  getScrollContainer: () => HTMLElement | null;
}

// ============================================================================
// Component
// ============================================================================

export const TrueLayoutEditor = forwardRef<TrueLayoutEditorHandle, TrueLayoutEditorProps>(
  function TrueLayoutEditor(
    {
      initialBlocks,
      pageConfig,
      zoom = 1,
      editable = true,
      onDocChange,
      onSelectionChange,
      onCellSelectionChange,
      onActivePageChange,
      onReady,
      width = '100%',
      height = '100%',
      pageGap = 24,
      minLinesAtBreak = 2,
      textStylePool,
      header,
      footer,
      documentTitle,
      onHeaderChange,
      onFooterChange,
      headerFooterEditable,
      collabHandle = null,
    },
    ref
  ) {
    // Refs
    const containerRef = useRef<HTMLDivElement>(null);
    const viewportRef = useRef<HTMLDivElement>(null);
    const pagesContainerRef = useRef<HTMLDivElement>(null);
    const hiddenEditorRef = useRef<HTMLDivElement>(null);
    const overlayRef = useRef<HTMLDivElement>(null);
    
    // State
    const [editorView, setEditorView] = useState<EditorView | null>(null);
    const [layout, setLayout] = useState<DocumentLayout | null>(null);
    const layoutRef = useRef<DocumentLayout | null>(null);
    const activePageIndexRef = useRef<number>(0);
    
    // Header/footer editing state
    const [hfEditState, setHfEditState] = useState<{
      isEditing: boolean;
      type: 'header' | 'footer';
      pageIndex: number;
    } | null>(null);
    
    // Keep layoutRef in sync with layout state
    useEffect(() => {
      layoutRef.current = layout;
    }, [layout]);
    
    // Memoized instances
    const measurerRef = useRef<DomMeasurer | null>(null);
    const painterRef = useRef<DomPainter | null>(null);
    const inputBridgeRef = useRef<InputBridge | null>(null);
    const selectionOverlayRef = useRef<SelectionOverlayManager | null>(null);
    const remoteCursorOverlayRef = useRef<RemoteCursorOverlay | null>(null);
    const tableInteractionRef = useRef<TableInteractionManager | null>(null);
    const imageInteractionRef = useRef<ImageInteractionManager | null>(null);
    const blocksRef = useRef<FlowBlock[]>([]);
    const performLayoutRef = useRef<((doc: PmNode, view: EditorView) => void) | null>(null);
    const measuresRef = useRef<Measure[]>([]);
    const blockPositionMapRef = useRef<Map<string, { start: number; end: number }>>(new Map());
    
    // Track currently hovered table/image for showing resize handles
    const hoveredTableRef = useRef<HTMLTableElement | null>(null);
    const hoveredImageRef = useRef<HTMLElement | null>(null);
    
    // Track if editor has been initialized - we only want to create it ONCE
    const editorInitializedRef = useRef(false);
    
    // Refs for callbacks to avoid stale closures in dispatchTransaction
    const onDocChangeRef = useRef(onDocChange);
    const onSelectionChangeRef = useRef(onSelectionChange);
    const onActivePageChangeRef = useRef(onActivePageChange);
    const textStylePoolRef = useRef(textStylePool);
    
    // Keep callback refs up to date
    useEffect(() => {
      onDocChangeRef.current = onDocChange;
    }, [onDocChange]);
    useEffect(() => {
      onSelectionChangeRef.current = onSelectionChange;
    }, [onSelectionChange]);
    useEffect(() => {
      onActivePageChangeRef.current = onActivePageChange;
    }, [onActivePageChange]);
    useEffect(() => {
      textStylePoolRef.current = textStylePool;
    }, [textStylePool]);
    
    // Create initial PM document - only used for FIRST initialization
    // After that, the editor manages its own state via transactions
    const initialDocRef = useRef<PmNode | null>(null);
    if (!initialDocRef.current) {
      initialDocRef.current = blocksToPmDoc(initialBlocks, textStylePool);
    }
    
    // Calculate dimensions
    const pageWidth = Math.floor(pageConfig.width * zoom);
    const pageHeight = Math.floor(pageConfig.height * zoom);
    const contentWidth = Math.floor(
      (pageConfig.width - pageConfig.margins.left - pageConfig.margins.right) * zoom
    );
    // The measurer measures at NATURAL size (natural font, natural width); the
    // layout engine and painter apply `scale: zoom` afterwards. Feeding it the
    // zoomed width while it renders an unscaled font made its chars-per-line
    // ratio wrong at any zoom != 1, mis-wrapping lines and mis-paginating.
    const measureContentWidth = Math.floor(
      pageConfig.width - pageConfig.margins.left - pageConfig.margins.right
    );

    // Initialize measurer
    useEffect(() => {
      measurerRef.current = new DomMeasurer({ contentWidth: measureContentWidth });

      return () => {
        measurerRef.current?.destroy();
        measurerRef.current = null;
      };
    }, [measureContentWidth]);
    
    // Header/footer click handler
    const handleHeaderFooterClick = useCallback((type: 'header' | 'footer', pageIndex: number) => {
      setHfEditState({ isEditing: true, type, pageIndex });
    }, []);
    
    // Determine if header/footer editing is enabled
    const isHfEditable = headerFooterEditable ?? editable;
    
    // Initialize painter once on mount
    useEffect(() => {
      painterRef.current = new DomPainter({
        pageConfig,
        scale: zoom,
        pageGap,
        header,
        footer,
        documentTitle,
        onHeaderFooterClick: handleHeaderFooterClick,
        headerFooterEditable: isHfEditable,
      });
      
      // Cleanup on unmount
      return () => {
        painterRef.current?.destroy();
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);
    
    // Update painter config when settings change (without recreating the instance)
    useEffect(() => {
      painterRef.current?.updateConfig({ 
        pageConfig, 
        scale: zoom, 
        pageGap, 
        header, 
        footer, 
        documentTitle,
        onHeaderFooterClick: handleHeaderFooterClick,
        headerFooterEditable: isHfEditable,
      });
    }, [pageConfig, zoom, pageGap, header, footer, documentTitle, handleHeaderFooterClick, isHfEditable]);
    
    // Configure virtualization with scroll container
    useEffect(() => {
      if (painterRef.current && viewportRef.current) {
        painterRef.current.setVirtualization({
          enabled: true,
          windowSize: 5,  // Render 5 pages at a time
          overscan: 1,    // Plus 1 page buffer on each side
          scrollContainer: viewportRef.current,
          // Invalidate selection cache when pages change during virtualization
          onPagesChange: () => {
            selectionOverlayRef.current?.invalidateSpanCache();
          },
        });
      }
    }, []);
    
    // Initialize selection overlay
    useEffect(() => {
      if (overlayRef.current) {
        selectionOverlayRef.current = new SelectionOverlayManager();
        selectionOverlayRef.current.initialize(overlayRef.current, pagesContainerRef.current || undefined);
        // Set scroll container for contentRectCache invalidation on scroll
        if (viewportRef.current) {
          selectionOverlayRef.current.setScrollContainer(viewportRef.current);
        }
      }

      return () => {
        selectionOverlayRef.current?.destroy();
        selectionOverlayRef.current = null;
      };
    }, []);

    // Initialize remote-cursor overlay when collab is attached.
    useEffect(() => {
      if (!collabHandle) return;
      if (!overlayRef.current || !selectionOverlayRef.current) return;
      const overlay = new RemoteCursorOverlay(
        collabHandle.awareness,
        selectionOverlayRef.current,
      );
      overlay.initialize(overlayRef.current);
      if (editorView) overlay.setEditorView(editorView);
      remoteCursorOverlayRef.current = overlay;
      return () => {
        overlay.destroy();
        remoteCursorOverlayRef.current = null;
      };
      // editorView is set after the PM editor mounts; we re-bind below.
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [collabHandle]);

    // Hand the live EditorView to the remote overlay once PM is up.
    useEffect(() => {
      remoteCursorOverlayRef.current?.setEditorView(editorView);
    }, [editorView]);
    
    // Initialize ProseMirror - only ONCE on mount
    useEffect(() => {
      // Skip if already initialized or no container
      if (editorInitializedRef.current || !hiddenEditorRef.current) return;
      // initialDocRef is only required for non-collab init; ySyncPlugin
      // seeds the editor from the Y.XmlFragment instead.
      if (!collabHandle && !initialDocRef.current) return;
      editorInitializedRef.current = true;

      const plugins = createPlugins(docsSchema);
      let state: EditorState;
      if (collabHandle) {
        ensureCollabCursorStyles();
        plugins.unshift(
          ySyncPlugin(collabHandle.xmlFragment),
          yCursorPlugin(collabHandle.awareness),
        );
        state = EditorState.create({ schema: docsSchema, plugins });
      } else {
        state = EditorState.create({
          doc: initialDocRef.current!,
          plugins,
        });
      }
      
      // `let view` (not const) avoids the temporal-dead-zone trap when
      // ySyncPlugin dispatches a synchronous transaction inside the
      // EditorView constructor — at that moment the outer `view` binding
      // hasn't been assigned yet. The `view ?? this` fallback uses the
      // EditorView that ProseMirror binds to `this` inside dispatchTransaction.
      let view: EditorView | undefined;
      // eslint-disable-next-line prefer-const -- `const` would TDZ-crash; see above
      view = new EditorView(hiddenEditorRef.current, {
        state,
        editable: () => editable,
        dispatchTransaction(transaction: Transaction) {
          // eslint-disable-next-line @typescript-eslint/no-this-alias
          const v = (view ?? (this as unknown as EditorView));
          const newState = v.state.apply(transaction);
          v.updateState(newState);

          if (transaction.docChanged) {
            // Trigger re-layout - this will also update selection overlay
            requestAnimationFrame(() => {
              if (view) performLayout(newState.doc, view);
            });
            
            // Notify parent of doc change (use ref to avoid stale closure)
            if (onDocChangeRef.current) {
              const docModel = proseMirrorToDocument(
                newState.doc,
                // Minimal stand-in document — only its blocks are read back.
                {
                  id: '',
                  title: '',
                  sections: [{ id: 'main', pageConfig: DEFAULT_PAGE_CONFIG, blocks: [] }],
                  defaultPageConfig: DEFAULT_PAGE_CONFIG,
                },
                textStylePoolRef.current  // Pass style pool for efficient style storage
              );
              onDocChangeRef.current(docModel.sections[0].blocks);
            }
            
            // Also update active marks when doc changes (e.g., font size applied)
            onSelectionChangeRef.current?.(newState);
          } else if (transaction.selectionSet) {
            // Only update selection overlay if doc didn't change (layout handles it otherwise)
            // Check for immediate flag (used during drag for responsiveness)
            const isImmediate = transaction.getMeta('immediateSelection') === true;
            selectionOverlayRef.current?.setFocused(v.hasFocus());
            selectionOverlayRef.current?.updateSelection(newState, v, isImmediate);
            onSelectionChangeRef.current?.(newState);
            
            // Update active page index - only check if we actually have the callback
            // to avoid unnecessary computation during drag selection
            if (onActivePageChangeRef.current && layoutRef.current && layoutRef.current.pages.length > 0) {
              const { from } = newState.selection;
              let newActivePageIndex = activePageIndexRef.current;
              
              // Binary search would be faster but pages are usually few
              // Only search if selection might have moved to a different page
              pageLoop: for (let i = 0; i < layoutRef.current.pages.length; i++) {
                const page = layoutRef.current.pages[i];
                for (const fragment of page.fragments) {
                  // Look up PM positions from the block position map
                  const posInfo = blockPositionMapRef.current.get(fragment.blockId);
                  if (posInfo && from >= posInfo.start && from <= posInfo.end) {
                    newActivePageIndex = i;
                    break pageLoop;
                  }
                }
              }
              
              // If page changed, update ref and notify
              if (newActivePageIndex !== activePageIndexRef.current) {
                activePageIndexRef.current = newActivePageIndex;
                onActivePageChangeRef.current?.({
                  pageIndex: newActivePageIndex,
                  totalPages: layoutRef.current.pages.length,
                  pageHeight,
                });
              }
            }
          }
        },
        attributes: {
          class: 'true-layout-hidden-editor',
          role: 'textbox',
          'aria-multiline': 'true',
        },
      });
      
      // Track focus changes on the hidden editor
      const handleEditorFocus = () => {
        selectionOverlayRef.current?.setFocused(true);
      };
      const handleEditorBlur = () => {
        // Only unfocus if we're not clicking on the pages container
        setTimeout(() => {
          if (!view.hasFocus() && document.activeElement !== pagesContainerRef.current) {
            selectionOverlayRef.current?.setFocused(false);
          }
        }, 0);
      };
      
      view.dom.addEventListener('focus', handleEditorFocus);
      view.dom.addEventListener('blur', handleEditorBlur);
      
      setEditorView(view);
      onReady?.(view);
      
      // Initial layout
      requestAnimationFrame(() => {
        performLayout(view.state.doc, view);
      });
      
      return () => {
        view.dom.removeEventListener('focus', handleEditorFocus);
        view.dom.removeEventListener('blur', handleEditorBlur);
        view.destroy();
        setEditorView(null);
        editorInitializedRef.current = false;  // Allow re-init if component remounts
      };
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [editable]);  // Only depend on editable, NOT on initialDoc
    
    // Initialize input bridge
    useEffect(() => {
      if (!editorView || !pagesContainerRef.current || !viewportRef.current) return;
      
      inputBridgeRef.current = createInputBridge();
      inputBridgeRef.current.initialize(editorView, pagesContainerRef.current, {
        onSelectionChange: (_from, _to) => {
          // Selection is handled by the editor's dispatch
        },
        onFocusChange: (focused) => {
          selectionOverlayRef.current?.setFocused(focused);
        },
        onCellSelectionChange: (selection) => {
          // Hide caret when a cell is selected
          selectionOverlayRef.current?.setCellSelectionActive(selection !== null);
          onCellSelectionChange?.(selection);
        },
        scrollContainer: viewportRef.current,
      });
      
      // Set initial focus state - the hidden editor might already be focused
      selectionOverlayRef.current?.setFocused(editorView.hasFocus());
      
      return () => {
        inputBridgeRef.current?.destroy();
        inputBridgeRef.current = null;
      };
    }, [editorView, onCellSelectionChange]);
    
    // Perform layout calculation
    const performLayout = useCallback((doc: PmNode, view: EditorView) => {
      // Removed logging for production
      if (!measurerRef.current || !painterRef.current || !pagesContainerRef.current) return;
      
      // 1. Convert ProseMirror doc to FlowBlocks
      const blocks = proseMirrorToFlowBlocks(doc, { resetIds: true });
      blocksRef.current = blocks;
      
      // 2. Create PM position map for blocks
      const pmPositions = createBlockPositionMap(doc, blocks);
      blockPositionMapRef.current = pmPositions;
      
      // 3. Measure all blocks (clear cache first to ensure fresh measurements)
      measurerRef.current.clearCache();
      const measures = measurerRef.current.measureBlocks(blocks);
      measuresRef.current = measures;
      
      // 4. Compute layout
      const newLayout = computeTrueLayout(blocks, measures, {
        pageConfig,
        scale: zoom,
        pageGap,
        minLinesAtBreak,
      });
      
      setLayout(newLayout);
      
      // 5. Paint pages with PM positions for click-to-position mapping
      painterRef.current.setData(blocks, measures, pmPositions);
      painterRef.current.paint(newLayout, pagesContainerRef.current);
      
      // 6. Update input bridge with new layout
      inputBridgeRef.current?.updateLayout(newLayout, blocks, measures, doc);
      
      // 6b. Refresh cell selection visual after repaint
      inputBridgeRef.current?.refreshCellSelectionVisual();
      
      // 7. Update selection overlay with new layout and render cursor IMMEDIATELY
      // Using immediate=true avoids the extra RAF delay for smoother cursor updates
      selectionOverlayRef.current?.setPagesMount(pagesContainerRef.current);
      selectionOverlayRef.current?.updateLayout(newLayout, blocks, measures, doc);
      selectionOverlayRef.current?.setFocused(view.hasFocus());
      selectionOverlayRef.current?.updateSelection(view.state, view, true);  // immediate=true

      // 8. Refresh remote-cursor overlay so peer carets/selections re-place on
      //    the freshly painted pages.
      remoteCursorOverlayRef.current?.refresh();
    }, [pageConfig, zoom, pageGap, minLinesAtBreak]);
    
    // Keep performLayoutRef in sync to avoid stale closures
    performLayoutRef.current = performLayout;
    
    // Trigger re-layout when page config, zoom, layout settings, or header/footer change
    useEffect(() => {
      if (editorView && pagesContainerRef.current) {
        // Use RAF to avoid multiple rapid re-layouts
        requestAnimationFrame(() => {
          performLayoutRef.current?.(editorView.state.doc, editorView);
        });
      }
      // Note: We use performLayoutRef to avoid adding performLayout to deps
      // which would cause infinite loops. The ref is always up-to-date.
    }, [pageConfig, zoom, pageGap, minLinesAtBreak, editorView, header, footer]);
    
    // Initialize table interaction manager
    useEffect(() => {
      if (!editorView || !pagesContainerRef.current) return;
      
      tableInteractionRef.current = createTableInteractionManager();
      tableInteractionRef.current.initialize(pagesContainerRef.current, editorView, {
        scale: zoom,
        onTableUpdate: () => {
          // Trigger re-layout after table modifications (use ref to avoid stale closure)
          if (editorView) {
            performLayoutRef.current?.(editorView.state.doc, editorView);
          }
        },
      });
      
      return () => {
        tableInteractionRef.current?.destroy();
        tableInteractionRef.current = null;
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [editorView]);
    
    // Initialize image interaction manager
    useEffect(() => {
      if (!editorView || !pagesContainerRef.current) return;
      
      imageInteractionRef.current = createImageInteractionManager();
      imageInteractionRef.current.initialize(pagesContainerRef.current, editorView, {
        onImageUpdate: () => {
          // Trigger re-layout after image modifications (use ref to avoid stale closure)
          if (editorView) {
            performLayoutRef.current?.(editorView.state.doc, editorView);
          }
        },
      });
      
      return () => {
        imageInteractionRef.current?.destroy();
        imageInteractionRef.current = null;
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [editorView]);
    
    // Scroll state detection - pause mouse processing during scroll
    const isScrollingRef = useRef<boolean>(false);
    const scrollTimeoutRef = useRef<number | null>(null);
    
    // FPS diagnostics for scroll performance debugging
    const fpsRef = useRef<{ frames: number; lastTime: number; fps: number }>({ frames: 0, lastTime: 0, fps: 0 });
    const [showDiagnostics, setShowDiagnostics] = useState(false);
    const [diagnosticData, setDiagnosticData] = useState({ fps: 0, pageCount: 0, scrolling: false });
    
    // Toggle diagnostics with Ctrl+Shift+D
    useEffect(() => {
      const handleKeyDown = (e: KeyboardEvent) => {
        if (e.ctrlKey && e.shiftKey && e.key === 'D') {
          setShowDiagnostics(prev => !prev);
        }
      };
      window.addEventListener('keydown', handleKeyDown);
      return () => window.removeEventListener('keydown', handleKeyDown);
    }, []);
    
    // FPS counter effect
    useEffect(() => {
      if (!showDiagnostics) return;
      
      let rafId: number;
      const measureFps = (now: number) => {
        fpsRef.current.frames++;
        const elapsed = now - fpsRef.current.lastTime;
        if (elapsed >= 1000) {
          fpsRef.current.fps = Math.round((fpsRef.current.frames * 1000) / elapsed);
          fpsRef.current.frames = 0;
          fpsRef.current.lastTime = now;
          
          const pageCount = pagesContainerRef.current?.querySelectorAll('.page').length || 0;
          setDiagnosticData({
            fps: fpsRef.current.fps,
            pageCount,
            scrolling: isScrollingRef.current,
          });
        }
        rafId = requestAnimationFrame(measureFps);
      };
      
      fpsRef.current.lastTime = performance.now();
      rafId = requestAnimationFrame(measureFps);
      
      return () => cancelAnimationFrame(rafId);
    }, [showDiagnostics]);
    
    // Detect scroll start/end to pause mouse move processing
    useEffect(() => {
      const viewport = viewportRef.current;
      if (!viewport) return;
      
      const handleScroll = () => {
        isScrollingRef.current = true;
        
        // Clear any pending timeout
        if (scrollTimeoutRef.current !== null) {
          clearTimeout(scrollTimeoutRef.current);
        }
        
        // Mark scroll as ended after 150ms of no scroll events
        scrollTimeoutRef.current = window.setTimeout(() => {
          isScrollingRef.current = false;
          scrollTimeoutRef.current = null;
        }, 150);
      };
      
      viewport.addEventListener('scroll', handleScroll, { passive: true });
      return () => {
        viewport.removeEventListener('scroll', handleScroll);
        if (scrollTimeoutRef.current !== null) {
          clearTimeout(scrollTimeoutRef.current);
        }
      };
    }, []);
    
    // Handle table/image hover for showing drag handles (paused during scroll)
    const handleMouseMove = useCallback((e: React.MouseEvent) => {
      // Skip processing entirely during scroll for maximum performance
      if (isScrollingRef.current || !pagesContainerRef.current) return;
      
      processMouseMove(e);
    }, []);
    
    const processMouseMove = useCallback((e: React.MouseEvent) => {
      const target = e.target as HTMLElement;
      
      // Handle table hover
      if (tableInteractionRef.current) {
        const tableEl = target.closest('.table-block') as HTMLTableElement | null;
        
        if (tableEl && tableEl !== hoveredTableRef.current) {
          hoveredTableRef.current = tableEl;
          tableInteractionRef.current.cancelScheduledHide();
          tableInteractionRef.current.showHandlesForTable(tableEl);
        } else if (tableEl && tableEl === hoveredTableRef.current) {
          tableInteractionRef.current.cancelScheduledHide();
        } else if (!tableEl && hoveredTableRef.current) {
          hoveredTableRef.current = null;
          tableInteractionRef.current.scheduleHideHandles();
        }
      }
      
      // Handle image hover
      if (imageInteractionRef.current) {
        const imageEl = target.closest('.image-block') as HTMLElement | null;
        
        if (imageEl && imageEl !== hoveredImageRef.current) {
          hoveredImageRef.current = imageEl;
          imageInteractionRef.current.cancelScheduledHide();
          imageInteractionRef.current.showHandlesForImage(imageEl);
        } else if (imageEl && imageEl === hoveredImageRef.current) {
          imageInteractionRef.current.cancelScheduledHide();
        } else if (!imageEl && hoveredImageRef.current) {
          hoveredImageRef.current = null;
          imageInteractionRef.current.scheduleHideHandles();
        }
      }
    }, []);
    
    const handleMouseLeave = useCallback(() => {
      // Schedule hide for tables
      if (tableInteractionRef.current) {
        hoveredTableRef.current = null;
        tableInteractionRef.current.scheduleHideHandles();
      }
      
      // Schedule hide for images
      if (imageInteractionRef.current) {
        hoveredImageRef.current = null;
        imageInteractionRef.current.scheduleHideHandles();
      }
    }, []);
    
    // Update layout when zoom changes
    useEffect(() => {
      if (editorView) {
        measurerRef.current?.updateContentWidth(measureContentWidth);
        painterRef.current?.updateConfig({ scale: zoom, pageGap });
        tableInteractionRef.current?.setScale(zoom);
        imageInteractionRef.current?.setScale(zoom);
        // Use ref to avoid dependency on performLayout identity
        performLayoutRef.current?.(editorView.state.doc, editorView);
      }
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [zoom, contentWidth, pageGap, editorView]);
    
    // Notify about active page changes (only when page index or total changes, not on scroll)
    const lastNotifiedRef = useRef<{ pageIndex: number; totalPages: number } | null>(null);
    
    useEffect(() => {
      if (!layoutRef.current) return;
      
      const activePageIndex = activePageIndexRef.current;
      const totalPages = layoutRef.current.pages.length;
      
      // Only notify if page index or total pages changed
      const last = lastNotifiedRef.current;
      if (!last || last.pageIndex !== activePageIndex || last.totalPages !== totalPages) {
        lastNotifiedRef.current = { pageIndex: activePageIndex, totalPages };
        onActivePageChange?.({
          pageIndex: activePageIndex,
          totalPages,
          pageHeight,
        });
      }
    }, [layout, pageHeight, onActivePageChange]);
    
    // Expose methods via ref
    useImperativeHandle(ref, () => ({
      getView: () => editorView,
      getState: () => editorView?.state ?? null,
      focus: () => {
        inputBridgeRef.current?.focus();
      },
      hasFocus: () => inputBridgeRef.current?.hasFocus() ?? false,
      getLayout: () => layout,
      reflow: () => {
        if (editorView) {
          performLayoutRef.current?.(editorView.state.doc, editorView);
        }
      },
      getScrollContainer: () => viewportRef.current,
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }), [editorView, layout]);
    
    // Calculate total height
    const totalHeight = layout?.totalHeight ?? pageHeight;
    
    return (
      <div
        ref={containerRef}
        className="true-layout-editor"
        style={{
          width,
          height,
          display: 'flex',
          flexDirection: 'column',
          backgroundColor: '#f8f9fa',
          overflow: 'hidden',
          position: 'relative',
        }}
      >
        {/* Hidden ProseMirror editor */}
        <div
          ref={hiddenEditorRef}
          className="true-layout-hidden-editor-container"
          style={{
            position: 'fixed',
            left: -9999,
            top: 0,
            width: contentWidth,
            opacity: 0,
            zIndex: -1,
            pointerEvents: 'none',
          }}
        />
        
        {/* Scrollable viewport */}
        <div
          ref={viewportRef}
          className="true-layout-viewport"
          style={{
            flex: 1,
            overflow: 'auto',
            padding: 24,
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            willChange: 'scroll-position',
            WebkitOverflowScrolling: 'touch',
          } as React.CSSProperties}
        >
          {/* Pages wrapper - contains both painted pages and overlay */}
          <div
            className="true-layout-pages-wrapper"
            style={{
              position: 'relative',
              width: pageWidth,
              height: totalHeight,
              flexShrink: 0,
              contain: 'layout style',
            }}
          >
            {/* Pages container - where DomPainter renders, receives input events */}
            <div
              ref={pagesContainerRef}
              className="true-layout-pages"
              style={{
                position: 'absolute',
                top: 0,
                left: 0,
                width: '100%',
                height: '100%',
                outline: 'none',
              }}
              tabIndex={0}
              onMouseMove={handleMouseMove}
              onMouseLeave={handleMouseLeave}
            />
            
            {/* Selection overlay - sibling to pages, scrolls together */}
            <div
              ref={overlayRef}
              className="selection-overlay"
              style={{
                position: 'absolute',
                top: 0,
                left: 0,
                width: '100%',
                height: '100%',
                pointerEvents: 'none',
                zIndex: 100,
              }}
            />
          </div>
        </div>
        
        {/* Diagnostics overlay - toggle with Ctrl+Shift+D */}
        {showDiagnostics && (
          <div
            style={{
              position: 'fixed',
              top: 100,
              right: 20,
              backgroundColor: 'rgba(0,0,0,0.85)',
              color: '#00ff00',
              padding: '12px 16px',
              borderRadius: 8,
              fontFamily: 'monospace',
              fontSize: 12,
              zIndex: 99999,
              minWidth: 180,
            }}
          >
            <div style={{ fontWeight: 'bold', marginBottom: 8, color: '#fff' }}>
              📊 Scroll Diagnostics
            </div>
            <div style={{ color: diagnosticData.fps < 30 ? '#ff4444' : diagnosticData.fps < 55 ? '#ffaa00' : '#00ff00' }}>
              FPS: {diagnosticData.fps}
            </div>
            <div style={{ color: diagnosticData.pageCount > 5 ? '#ffaa00' : '#00ff00' }}>
              Pages in DOM: {diagnosticData.pageCount}
            </div>
            <div>
              Scrolling: {diagnosticData.scrolling ? '✓' : '—'}
            </div>
            <div style={{ marginTop: 8, fontSize: 10, color: '#888' }}>
              Press Ctrl+Shift+D to close
            </div>
          </div>
        )}
        
        {/* Styles */}
        <style>{`
          ${getSelectionOverlayStyles()}
          
          .true-layout-hidden-editor-container .ProseMirror {
            outline: none;
            white-space: pre-wrap;
            word-wrap: break-word;
          }
          
          .true-layout-pages:focus {
            outline: none;
          }
          
          .true-layout-pages .page {
            font-family: Arial, sans-serif;
          }
          
          .true-layout-pages .page-content {
            white-space: pre-wrap;
            word-wrap: break-word;
          }
          
          .true-layout-pages .fragment p {
            white-space: pre-wrap;
            word-wrap: break-word;
          }
          
          .true-layout-pages .fragment h1,
          .true-layout-pages .fragment h2,
          .true-layout-pages .fragment h3,
          .true-layout-pages .fragment h4,
          .true-layout-pages .fragment h5,
          .true-layout-pages .fragment h6 {
            white-space: pre-wrap;
            word-wrap: break-word;
          }
          
          .true-layout-pages a {
            color: #1a73e8;
            text-decoration: underline;
          }
          
          .true-layout-pages table {
            border-collapse: collapse;
          }
          
          .true-layout-pages th,
          .true-layout-pages td {
            border: 1px solid #dadce0;
          }
          
          .true-layout-pages hr {
            border: none;
            border-top: 1px solid #dadce0;
          }
          
          .true-layout-pages img {
            max-width: 100%;
          }
          
          /* Cell selection styles - border only, no background overlay */
          .true-layout-pages .table-cell.cell-selected {
            outline: 2px solid #1a73e8 !important;
            outline-offset: -2px !important;
          }
          
          /* Additional selector for better specificity */
          .true-layout-pages table td.table-cell.cell-selected,
          .true-layout-pages table th.table-cell.cell-selected {
            outline: 2px solid #1a73e8 !important;
            outline-offset: -2px !important;
          }
          
          /* Table resize handle styles */
          .table-col-resize-handle:hover,
          .table-row-resize-handle:hover {
            background: rgba(26, 115, 232, 0.3) !important;
          }
          
          .table-drag-handle:hover {
            background: #f1f3f4 !important;
          }
          
          .table-drag-handle:active {
            cursor: grabbing !important;
            background: #e8eaed !important;
          }
          
          /* Context menu styles */
          .table-context-menu .context-menu-item:hover {
            background: #f1f3f4 !important;
          }
        `}</style>
        
        {/* Header/Footer Editor - Inline Google Docs-style */}
        {hfEditState && pagesContainerRef.current && viewportRef.current && (
          <HeaderFooterEditor
            type={hfEditState.type}
            content={hfEditState.type === 'header' 
              ? (header ?? { blocks: [] }) 
              : (footer ?? { blocks: [] })}
            onChange={(newContent) => {
              if (hfEditState.type === 'header') {
                onHeaderChange?.(newContent);
              } else {
                onFooterChange?.(newContent);
              }
            }}
            isEditing={hfEditState.isEditing}
            onClose={() => setHfEditState(null)}
            pageIndex={hfEditState.pageIndex}
            totalPages={layout?.pages.length ?? 1}
            documentTitle={documentTitle}
            contentWidth={contentWidth}
            scale={zoom}
            position={(() => {
              // Calculate the position of the header/footer area
              const containerRect = pagesContainerRef.current!.getBoundingClientRect();
              const viewportRect = viewportRef.current!.getBoundingClientRect();
              const scrollTop = viewportRef.current!.scrollTop;
              const pageGapPx = pageGap * zoom;
              const pageHeightPx = pageHeight;
              
              // Calculate page Y position in the document
              const pageY = hfEditState.pageIndex * (pageHeightPx + pageGapPx);
              // Convert to viewport position
              const pageTop = containerRect.top + pageY - scrollTop + viewportRect.top - containerRect.top;
              
              const margins = {
                top: pageConfig.margins.top * zoom,
                bottom: pageConfig.margins.bottom * zoom,
                left: pageConfig.margins.left * zoom,
                right: pageConfig.margins.right * zoom,
                header: (pageConfig.margins.header ?? 48) * zoom,
                footer: (pageConfig.margins.footer ?? 48) * zoom,
              };
              
              return {
                top: hfEditState.type === 'header' ? margins.header : undefined,
                bottom: hfEditState.type === 'footer' ? margins.footer : undefined,
                left: containerRect.left + margins.left,
                pageY: pageTop,
              };
            })()}
            mainEditorView={editorView}
          />
        )}
      </div>
    );
  }
);

// Re-export types
export type { PageConfig } from './true-layout-engine';
export type { FlowBlock } from './flow-blocks';

