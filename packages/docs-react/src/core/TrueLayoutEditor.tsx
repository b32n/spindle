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
  useMemo,
  forwardRef,
  useImperativeHandle,
} from 'react';
import { EditorState, TextSelection, Transaction } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { Node as PmNode } from 'prosemirror-model';
import { docsSchema, createPlugins, blocksToPmDoc, sectionsToPmDoc, proseMirrorToDocument, toggleCheckboxControl, findNoteRef, insertNoteReference, noteOrdinalAt, noteTextFragment, insertTableOfContents, updateTableOfContents, removeTableOfContents, Block, DEFAULT_PAGE_CONFIG } from '@b32nio/spindle-docs-core';
import type { CollabHandle } from '@b32nio/spindle-docs-core/collab';
import { ySyncPlugin, yUndoPlugin, yCursorPlugin } from 'y-prosemirror';
import { ensureCollabCursorStyles } from './collab-cursor-styles';

import { FlowBlock, NoteKey, noteKeyString } from './flow-blocks';
import { proseMirrorToSectionedFlowBlocks, createBlockPositionMap, type RevisionView } from './pm-to-blocks';
import { applyPageNotes, createFootnoteHooks, footnoteRefs, type FootnotePlan } from './footnote-layout';
import { DomMeasurer, Measure } from './measurer';
import { findBookmarkPos, findElementAtPos, internalLinkTarget } from './internal-links';
import { computeTrueLayout, DocumentLayout, PageConfig, getMaxPageWidth, getPageConfig, getPageX, getPageY } from './true-layout-engine';
import { computeSectionedLayout, formatPageNumber, matchSectionsToBreaks, resolveInheritedContent, type LayoutSection, type LayoutSectionBreak } from './section-layout';
import { DomPainter, HeaderFooterContent } from './dom-painter';
import { InputBridge, createInputBridge, CellSelection } from './input-bridge';
import { SelectionOverlayManager, getSelectionOverlayStyles } from './selection-overlay';
import { RemoteCursorOverlay } from './remote-cursor-overlay';
import { TableInteractionManager, createTableInteractionManager } from './table-interactions';
import { ImageInteractionManager, createImageInteractionManager } from './image-interactions';
import { HeaderFooterEditor } from '../components/HeaderFooterEditor';
import { NoteEditor } from '../components/NoteEditor';

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
  /** Index of the section the active page belongs to (0 for a document of one section). */
  sectionIndex?: number;
  /** Id of that section, when the editor was given `sections`. */
  sectionId?: string;
}

/** One section of a document shown by the editor: its page setup, content and what it shows in the margins. */
export interface TrueLayoutSection {
  /** Section id: the `section_break` that ends the section names it; the last section has no break. */
  id: string;
  /** Page setup as rendered (width and height already swapped for landscape). */
  pageConfig: PageConfig;
  /** The blocks the editor starts with (read once, like `initialBlocks`). */
  blocks: Block[];
  header?: HeaderFooterContent;
  footer?: HeaderFooterContent;
  /** How the section starts; default a new page. */
  breakType?: LayoutSectionBreak;
  pageNumbering?: { start?: number; format?: string };
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
  /**
   * The sections of the document, in order. When given, the editor shows all of them (each with its own page setup,
   * header and footer), `initialBlocks`, `pageConfig`, `header` and `footer` are ignored, and changes are reported to
   * `onSectionsChange` instead of `onDocChange`. Page setup, headers and footers are read on every render; the blocks
   * only when the editor mounts.
   */
  sections?: TrueLayoutSection[];
  /**
   * Called when the content changes while `sections` is set: one entry per section of the document now, in order. A
   * section whose break was deleted is no longer listed.
   */
  onSectionsChange?: (sections: Array<{ id: string; blocks: Block[] }>) => void;
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
  /**
   * Lay footnotes out at the bottom of the page their reference is on (default: true). When false they are collected in
   * the notes area at the end of each section, with the endnotes.
   */
  footnotesAtPageBottom?: boolean;
  /**
   * How tracked changes show (default `markup`): every change marked, or the text as if every change were accepted
   * (`final`) or rejected (`original`). The document itself is not changed.
   */
  revisionView?: RevisionView;
  /** Text style pool for efficient style storage (recommended for large documents) */
  textStylePool?: import('@b32nio/spindle-docs-core').TextStylePool;
  /** Header content configuration */
  header?: HeaderFooterContent;
  /** Footer content configuration */
  footer?: HeaderFooterContent;
  /** Document title (used for dynamic field resolution in headers/footers) */
  documentTitle?: string;
  /**
   * Called when header content changes. With `sections`, `sectionId` is the section whose header it is: the one that
   * has the header shown on the edited page (a section with none edits the nearest earlier one that has).
   */
  onHeaderChange?: (content: HeaderFooterContent, sectionId?: string) => void;
  /** Called when footer content changes (`sectionId` as for `onHeaderChange`) */
  onFooterChange?: (content: HeaderFooterContent, sectionId?: string) => void;
  /**
   * Called when the user links a section's header or footer to the previous section (`linked`: the section shows the
   * previous one's and keeps none of its own) or unlinks it (it gets its own copy, `content`, to edit).
   */
  onHeaderFooterLinkChange?: (type: 'header' | 'footer', sectionId: string, linked: boolean, content: HeaderFooterContent) => void;
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
  /** Insert a footnote or endnote reference at the caret and start editing the new note's text. False when it cannot go there. */
  insertNote: (kind: 'footnote' | 'endnote') => boolean;
  /**
   * Insert a table of contents after the block the caret is in, update every table of contents from the headings (and the
   * pages they are on), or remove the one at the caret. False when there is nothing to do or it cannot be done there.
   */
  tableOfContents: (action: 'insert' | 'update' | 'remove') => boolean;
  /** Put the caret at a position of the document and scroll it into view (the pages outside the window are brought in first). */
  goTo: (pos: number) => void;
}

// ============================================================================
// Component
// ============================================================================

/** For each section, the index of the section whose content it shows: itself, or the nearest earlier one that has some. */
function ownersOf(contents: ReadonlyArray<unknown>): number[] {
  let owner = 0;
  return contents.map((content, i) => {
    if (content !== undefined) owner = i;
    return owner;
  });
}

export const TrueLayoutEditor = forwardRef<TrueLayoutEditorHandle, TrueLayoutEditorProps>(
  function TrueLayoutEditor(
    {
      initialBlocks,
      pageConfig: pageConfigProp,
      zoom = 1,
      editable = true,
      onDocChange,
      sections,
      onSectionsChange,
      onSelectionChange,
      onCellSelectionChange,
      onActivePageChange,
      onReady,
      width = '100%',
      height = '100%',
      pageGap = 24,
      minLinesAtBreak = 2,
      footnotesAtPageBottom = true,
      revisionView = 'markup',
      textStylePool,
      header: headerProp,
      footer: footerProp,
      documentTitle,
      onHeaderChange,
      onFooterChange,
      onHeaderFooterLinkChange,
      headerFooterEditable,
      collabHandle = null,
    },
    ref
  ) {
    // With several sections the first one's setup stands for the document where one setup is needed (the painter's
    // default, the hidden editor's width); each page uses its own section's setup.
    // (Kept referentially stable: a new object on every render of the host would re-lay the document out and repaint
    // the pages between a mouse press and its release, losing the click.)
    const pageConfigKey = JSON.stringify(sections?.[0]?.pageConfig ?? pageConfigProp);
     
    const pageConfig = useMemo<PageConfig>(() => JSON.parse(pageConfigKey), [pageConfigKey]);
    const header = sections ? undefined : headerProp;
    const footer = sections ? undefined : footerProp;
    const sectionsRef = useRef(sections);
    sectionsRef.current = sections;
    const onSectionsChangeRef = useRef(onSectionsChange);
    onSectionsChangeRef.current = onSectionsChange;
    // What re-lays the document out when it changes: the page setup, header, footer, break type and numbering of every
    // section (not the blocks, which the editor owns once mounted).
    const sectionsKey = useMemo(
      () => (sections ? JSON.stringify(sections.map((s) => [s.id, s.pageConfig, s.header, s.footer, s.breakType, s.pageNumbering])) : ''),
      [sections],
    );

    // Refs
    const containerRef = useRef<HTMLDivElement>(null);
    const viewportRef = useRef<HTMLDivElement>(null);
    const pagesContainerRef = useRef<HTMLDivElement>(null);
    const pagesWrapperRef = useRef<HTMLDivElement>(null);
    const hiddenEditorRef = useRef<HTMLDivElement>(null);
    const overlayRef = useRef<HTMLDivElement>(null);
    
    // State
    const [editorView, setEditorView] = useState<EditorView | null>(null);
    const [layout, setLayout] = useState<DocumentLayout | null>(null);
    /** The note whose text is being edited in place, and where its editor goes (in the pages wrapper). */
    const [noteEdit, setNoteEdit] = useState<NoteKey | null>(null);
    const [noteEditPos, setNoteEditPos] = useState<{ top: number; left: number; width: number; minHeight: number } | null>(null);
    const layoutRef = useRef<DocumentLayout | null>(null);
    /** Set after a table of contents was written: the layout it caused may move headings to other pages, so it is updated once more when laid out. */
    const tocRefreshRef = useRef(false);
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
    // For each section of the last layout: its id, and the index of the section that owns the header and the footer it shows.
    const sectionOwnersRef = useRef<{
      ids: string[];
      headers: Array<HeaderFooterContent | undefined>;
      footers: Array<HeaderFooterContent | undefined>;
      headerOwners: number[];
      footerOwners: number[];
    } | null>(null);
    // The active page as reported to the host: its height and the section it belongs to.
    const activePageInfo = (current: DocumentLayout, pageIndex: number): ActivePageInfo => {
      const sectionIndex = current.pages[pageIndex]?.sectionIndex;
      return {
        pageIndex,
        totalPages: current.pages.length,
        pageHeight: Math.floor(getPageConfig(current, pageIndex).height * zoom),
        ...(sectionIndex !== undefined && { sectionIndex, sectionId: sectionOwnersRef.current?.ids[sectionIndex] }),
      };
    };
    const activePageInfoRef = useRef(activePageInfo);
    activePageInfoRef.current = activePageInfo;
    const blockPositionMapRef = useRef<Map<string, { start: number; end: number }>>(new Map());
    
    // Track currently hovered table/image for showing resize handles
    const hoveredTableRef = useRef<HTMLTableElement | null>(null);
    const hoveredImageRef = useRef<HTMLElement | null>(null);
    
    // Track if editor has been initialized - we only want to create it ONCE
    const editorInitializedRef = useRef(false);
    
    // Refs for callbacks to avoid stale closures in dispatchTransaction
    // `editable` is read live by the view: switching to and from read-only (Viewing mode, the Final and Original views) must
    // not rebuild the editor, which would start again from the document the editor was mounted with and lose every edit.
    const editableRef = useRef(editable);
    editableRef.current = editable;
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
      initialDocRef.current = sections ? sectionsToPmDoc(sections, textStylePool) : blocksToPmDoc(initialBlocks, textStylePool);
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

      const plugins = createPlugins(docsSchema, { collab: !!collabHandle });
      let state: EditorState;
      if (collabHandle) {
        ensureCollabCursorStyles();
        plugins.unshift(
          ySyncPlugin(collabHandle.xmlFragment),
          yCursorPlugin(collabHandle.awareness),
          yUndoPlugin(), // undo/redo only the local user's edits
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
        editable: () => editableRef.current,
        dispatchTransaction(transaction: Transaction) {
           
          const v = (view ?? (this as unknown as EditorView));
          const newState = v.state.apply(transaction);
          v.updateState(newState);

          if (transaction.docChanged) {
            // Trigger re-layout - this will also update selection overlay
            requestAnimationFrame(() => {
              if (view) performLayout(newState.doc, view);
            });
            
            // Notify parent of doc change (use ref to avoid stale closure)
            if (sectionsRef.current && onSectionsChangeRef.current) {
              const known = sectionsRef.current;
              const docModel = proseMirrorToDocument(
                newState.doc,
                // Stand-in document: only the section ids (to find the last section) and the blocks are used.
                {
                  id: '',
                  title: '',
                  sections: known.map((s) => ({ id: s.id, pageConfig: DEFAULT_PAGE_CONFIG, blocks: [] })),
                  defaultPageConfig: DEFAULT_PAGE_CONFIG,
                },
                textStylePoolRef.current
              );
              onSectionsChangeRef.current(docModel.sections.map((s) => ({ id: s.id, blocks: s.blocks })));
            } else if (onDocChangeRef.current) {
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
              // The editor shows one flow: a shared (collab) document with several sections hands over all their
              // blocks, as before section breaks existed in the editor's document.
              onDocChangeRef.current(docModel.sections.flatMap((section) => section.blocks));
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
                onActivePageChangeRef.current?.(activePageInfoRef.current(layoutRef.current, newActivePageIndex));
              }
            }
          }
        },
        attributes: {
          class: 'true-layout-hidden-editor',
          role: 'textbox',
          'aria-multiline': 'true',
          'aria-label': 'Document',
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
    }, []);  // Not on initialDoc (the document lives in the editor after the first render), nor on editable (read through a ref)
    
    // Switching between editable and read-only: the view reads `editable` again.
    useEffect(() => {
      editorView?.setProps({ editable: () => editableRef.current });
    }, [editable, editorView]);

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
    
    // Clicking an internal link (`#name`) scrolls to the bookmark of that name.
    useEffect(() => {
      const container = pagesContainerRef.current;
      const viewport = viewportRef.current;
      if (!editorView || !container || !viewport) return;
      const onClick = (e: MouseEvent) => {
        const link = (e.target as Element | null)?.closest?.('a[href^="#"]');
        const name = internalLinkTarget(link?.getAttribute('href'));
        if (!link || name === null) return;
        e.preventDefault();
        const pos = findBookmarkPos(editorView.state.doc, name);
        if (pos === null) return;
        // Bring the page into view (pages outside the window are not painted), then the line itself.
        const current = layoutRef.current;
        const page = current?.pages.findIndex((p) =>
          p.fragments.some((f) => {
            const info = blockPositionMapRef.current.get(f.blockId);
            return !!info && pos >= info.start && pos <= info.end;
          })
        );
        if (page !== undefined && page >= 0 && current) viewport.scrollTop = getPageY(current, page);
        const { state } = editorView;
        editorView.dispatch(state.tr.setSelection(TextSelection.near(state.doc.resolve(pos))));
        let tries = 0;
        const reveal = () => {
          const el = findElementAtPos(container, pos);
          if (el) el.scrollIntoView({ block: 'center' });
          else if (++tries < 10) requestAnimationFrame(reveal);
        };
        requestAnimationFrame(reveal);
      };
      container.addEventListener('click', onClick);
      return () => container.removeEventListener('click', onClick);
    }, [editorView]);

    // Going to a position (the review panel's next and previous change): the caret there, and the line in view.
    const goToRef = useRef<((pos: number) => void) | null>(null);
    useEffect(() => {
      const container = pagesContainerRef.current;
      const viewport = viewportRef.current;
      if (!editorView || !container || !viewport) return;
      let frame = 0;
      goToRef.current = (target) => {
        // A newer go-to takes over from the retries of an older one.
        cancelAnimationFrame(frame);
        const { state } = editorView;
        const pos = Math.max(0, Math.min(target, state.doc.content.size));
        const current = layoutRef.current;
        const page = current?.pages.findIndex((p) =>
          p.fragments.some((f) => {
            const info = blockPositionMapRef.current.get(f.blockId);
            return !!info && pos >= info.start && pos <= info.end;
          })
        );
        if (page !== undefined && page >= 0 && current) viewport.scrollTop = getPageY(current, page);
        editorView.dispatch(state.tr.setSelection(TextSelection.near(state.doc.resolve(pos))));
        let tries = 0;
        const reveal = () => {
          const el = findElementAtPos(container, pos);
          if (el) el.scrollIntoView({ block: 'center' });
          else if (++tries < 10) frame = requestAnimationFrame(reveal);
        };
        frame = requestAnimationFrame(reveal);
      };
      return () => {
        cancelAnimationFrame(frame);
        goToRef.current = null;
      };
    }, [editorView]);

    // Clicking the number of a note reference scrolls to the note in the notes area; clicking a note there (its number or
    // its text) goes back to the reference.
    useEffect(() => {
      const container = pagesContainerRef.current;
      const viewport = viewportRef.current;
      if (!editorView || !container || !viewport) return;
      /** Scroll to the page a block is laid out on (pages outside the window are not painted), then bring `find`'s element into view once painted. */
      const reveal = (pageOf: (fragmentBlockId: string) => boolean, find: () => Element | null) => {
        const current = layoutRef.current;
        const page = current?.pages.findIndex((p) => p.fragments.some((f) => pageOf(f.blockId)));
        if (page !== undefined && page >= 0 && current) viewport.scrollTop = getPageY(current, page);
        let tries = 0;
        const scroll = () => {
          const el = find();
          if (el) el.scrollIntoView({ block: 'center' });
          else if (++tries < 10) requestAnimationFrame(scroll);
        };
        requestAnimationFrame(scroll);
      };
      const onClick = (e: MouseEvent) => {
        const target = e.target as Element | null;
        // A click on a note's number in the notes area goes to its reference, one on its text edits it (read-only: goes to the
        // reference too); a click on a reference's number goes to the note.
        const entry = target?.closest?.<HTMLElement>('[data-note-target]');
        const ref = entry ? undefined : target?.closest?.<HTMLElement>('[data-note-ref]');
        const key = entry ? entry.dataset.noteTarget : ref?.dataset.noteRef;
        if (!key) return;
        e.preventDefault();
        if (entry && editable && !target?.closest?.('[data-note-ref]')) {
          const [kind, ordinal] = key.split(':');
          setNoteEdit({ kind: kind === 'endnote' ? 'endnote' : 'footnote', ordinal: Number(ordinal) });
          return;
        }
        if (ref) {
          // From the reference to its note.
          const block = blocksRef.current.find((b) => b.kind === 'paragraph' && !!b.noteTarget && noteKeyString(b.noteTarget) === key);
          if (!block) return;
          reveal(
            (id) => id === block.id,
            () => container.querySelector(`[data-note-target="${key}"]`)
          );
          return;
        }
        // From the note to its reference: the nth reference of its kind in the document.
        const [kind, ordinal] = key.split(':');
        const found = findNoteRef(editorView.state.doc, kind === 'endnote' ? 'endnote' : 'footnote', Number(ordinal));
        if (!found) return;
        const at = found.pos;
        editorView.dispatch(editorView.state.tr.setSelection(TextSelection.near(editorView.state.doc.resolve(at))));
        reveal(
          (id) => {
            const info = blockPositionMapRef.current.get(id);
            return !!info && at >= info.start && at <= info.end;
          },
          () => container.querySelector(`[data-note-ref="${key}"]:not([data-note-target] *)`)
        );
      };
      container.addEventListener('click', onClick);
      return () => container.removeEventListener('click', onClick);
    }, [editorView, editable]);

    // Where the editor of the note being edited goes: over its entry in the notes area, once that is painted.
    useEffect(() => {
      if (!noteEdit) {
        setNoteEditPos(null);
        return;
      }
      const container = pagesContainerRef.current;
      const viewport = viewportRef.current;
      if (!container || !viewport) return;
      const key = noteKeyString(noteEdit);
      const block = blocksRef.current.find((b) => b.kind === 'paragraph' && !!b.noteTarget && noteKeyString(b.noteTarget) === key);
      const current = layoutRef.current;
      const page = block && current ? current.pages.findIndex((p) => p.fragments.some((f) => f.blockId === block.id)) : -1;
      if (page >= 0 && current) viewport.scrollTop = getPageY(current, page);
      let cancelled = false;
      let tries = 0;
      const place = () => {
        if (cancelled) return;
        const entry = container.querySelector<HTMLElement>(`[data-note-target="${key}"]`);
        const wrapper = pagesWrapperRef.current;
        if (entry && wrapper) {
          const at = entry.getBoundingClientRect();
          const origin = wrapper.getBoundingClientRect();
          setNoteEditPos({ top: at.top - origin.top, left: at.left - origin.left, width: at.width, minHeight: at.height });
        } else if (++tries < 30) {
          requestAnimationFrame(place);
        }
      };
      requestAnimationFrame(place);
      return () => {
        cancelled = true;
      };
    }, [noteEdit, layout]);

    // Clicking the glyph of a check box content control toggles it.
    useEffect(() => {
      const container = pagesContainerRef.current;
      if (!editorView || !container) return;
      const onClick = (e: MouseEvent) => {
        const glyph = (e.target as Element | null)?.closest?.<HTMLElement>('[data-control-checkbox]');
        if (!glyph || !editorView.editable) return;
        const pos = Number(glyph.dataset.pmStart);
        if (Number.isNaN(pos)) return;
        if (toggleCheckboxControl(pos)(editorView.state, editorView.dispatch)) e.preventDefault();
      };
      container.addEventListener('click', onClick);
      return () => container.removeEventListener('click', onClick);
    }, [editorView]);

    /** The page number (as shown) of the page a position of the document is on, from the last layout. */
    const pageLabelAtPos = (): ((pos: number) => string | undefined) => {
      const current = layoutRef.current;
      const ranges: Array<{ start: number; end: number; page: number }> = [];
      current?.pages.forEach((page, index) => {
        for (const fragment of page.fragments) {
          const info = blockPositionMapRef.current.get(fragment.blockId);
          if (info) ranges.push({ ...info, page: index });
        }
      });
      return (pos) => {
        const hit = ranges.find((r) => pos >= r.start && pos < r.end);
        const page = hit && current ? current.pages[hit.page] : undefined;
        return page ? formatPageNumber(page.pageNumber ?? page.pageIndex + 1, page.pageNumberFormat) : undefined;
      };
    };

    /**
     * Insert, update or remove tables of contents (see `TrueLayoutEditorHandle.tableOfContents`). `refresh`: write the tables
     * again once the layout they cause is known (their page numbers can change what is on which page).
     */
    const runTableOfContents = (view: EditorView, action: 'insert' | 'update' | 'remove', refresh = true): boolean => {
      const labels = pageLabelAtPos();
      const command =
        action === 'insert' ? insertTableOfContents(labels) : action === 'update' ? updateTableOfContents(labels) : removeTableOfContents;
      let changed = false;
      const done = command(view.state, (tr) => {
        // An update that changes nothing is not a step in the history.
        if (action !== 'update' || !tr.doc.eq(view.state.doc)) {
          view.dispatch(tr);
          changed = true;
        }
      });
      if (changed && refresh && action !== 'remove') tocRefreshRef.current = true;
      return done;
    };

    const runTableOfContentsRef = useRef<typeof runTableOfContents | null>(null);
    runTableOfContentsRef.current = runTableOfContents;

    // Perform layout calculation
    const performLayout = useCallback((doc: PmNode, view: EditorView) => {
      // Removed logging for production
      if (!measurerRef.current || !painterRef.current || !pagesContainerRef.current) return;
      
      // 1. Convert ProseMirror doc to FlowBlocks (and find where each section's blocks are). Footnotes at the bottom of their
      //    pages come back separately; the blocks that draw them follow the body's blocks in the arrays given to the layout.
      const known = sectionsRef.current;
      const sectioned = proseMirrorToSectionedFlowBlocks(doc, { resetIds: true, footnotesAtPageBottom, revisionView });
      const bodyBlocks = sectioned.blocks;
      const noteBlocks = sectioned.footnotes.map((f) => f.block);
      const blocks: FlowBlock[] = [...bodyBlocks, ...noteBlocks];
      blocksRef.current = blocks;
      
      // 2. Create PM position map for blocks (the notes' blocks have no node)
      const pmPositions = createBlockPositionMap(doc, bodyBlocks);
      blockPositionMapRef.current = pmPositions;
      
      // 3. Measure all blocks (clear cache first to ensure fresh measurements)
      measurerRef.current.clearCache();
      let measures: Measure[];
      let newLayout: DocumentLayout;
      /** Measure the footnote blocks of each section at that section's width; they follow the body's measures. */
      const measureNotes = (widthOf: (sectionIndex: number) => number): Measure[] => {
        const out: Measure[] = [];
        for (let i = 0; i < sectioned.sections.length; i++) {
          const own = sectioned.footnotes.filter((f) => f.sectionIndex === i).map((f) => f.block);
          if (own.length === 0) continue;
          measurerRef.current!.updateContentWidth(Math.floor(widthOf(i)));
          out.push(...measurerRef.current!.measureBlocks(own));
        }
        return out;
      };
      const plan = (allMeasures: Measure[]): FootnotePlan | undefined => {
        if (noteBlocks.length === 0) return undefined;
        const noteIndex = new Map<string, { index: number; id: string }>();
        noteBlocks.forEach((b, k) => b.noteTarget && noteIndex.set(noteKeyString(b.noteTarget), { index: bodyBlocks.length + k, id: b.id }));
        return { refs: footnoteRefs(blocks, allMeasures, bodyBlocks.length), noteBlocks: noteIndex };
      };
      if (known) {
        // Each section is measured at its own content width and laid out with its own page setup.
        const matched = matchSectionsToBreaks(sectioned.sections.map((s) => s.sectionId), known);
        const headers = resolveInheritedContent(matched.map((s) => s.header));
        const footers = resolveInheritedContent(matched.map((s) => s.footer));
        const sectionWidth = (i: number) => matched[i].pageConfig.width - matched[i].pageConfig.margins.left - matched[i].pageConfig.margins.right;
        measures = [];
        const layoutSections: LayoutSection[] = sectioned.sections.map((range, i) => {
          const setup = matched[i];
          measurerRef.current!.updateContentWidth(Math.floor(sectionWidth(i)));
          measures.push(...measurerRef.current!.measureBlocks(bodyBlocks.slice(range.blockStart, range.blockEnd)));
          return {
            pageConfig: setup.pageConfig,
            blockStart: range.blockStart,
            blockEnd: range.blockEnd,
            ...(i > 0 && setup.breakType && { breakType: setup.breakType }),
            ...(setup.pageNumbering && { pageNumbering: setup.pageNumbering }),
            differentFirstPage: !!(headers[i]?.differentFirstPage || footers[i]?.differentFirstPage),
            differentOddEven: !!(headers[i]?.differentOddEven || footers[i]?.differentOddEven),
          };
        });
        measures.push(...measureNotes(sectionWidth));
        const notesPlan = plan(measures);
        newLayout = computeSectionedLayout(blocks, measures, layoutSections, { scale: zoom, pageGap, minLinesAtBreak, ...(notesPlan && { footnotes: createFootnoteHooks(notesPlan, measures, zoom) }) });
        if (notesPlan) newLayout = applyPageNotes(newLayout, notesPlan);
        painterRef.current.updateConfig({ pageConfig: matched[0].pageConfig, sectionHeaders: headers, sectionFooters: footers });
        sectionOwnersRef.current = { ids: matched.map((s) => s.id), headers, footers, headerOwners: ownersOf(matched.map((s) => s.header)), footerOwners: ownersOf(matched.map((s) => s.footer)) };
      } else {
        measures = [...measurerRef.current.measureBlocks(bodyBlocks), ...measureNotes(() => measureContentWidth)];
        const notesPlan = plan(measures);
        newLayout = computeTrueLayout(bodyBlocks, measures.slice(0, bodyBlocks.length), {
          pageConfig,
          scale: zoom,
          pageGap,
          minLinesAtBreak,
          ...(notesPlan && { footnotes: createFootnoteHooks(notesPlan, measures, zoom) }),
        });
        if (notesPlan) newLayout = applyPageNotes(newLayout, notesPlan);
      }
      measuresRef.current = measures;
      layoutRef.current = newLayout;
      
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

      // 9. A table of contents written since the last layout is brought in line with where the headings are now.
      if (tocRefreshRef.current) {
        tocRefreshRef.current = false;
        runTableOfContentsRef.current?.(view, 'update', false);
      }
    }, [pageConfig, zoom, pageGap, minLinesAtBreak, footnotesAtPageBottom, revisionView, measureContentWidth]);
    
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
    }, [pageConfig, zoom, pageGap, minLinesAtBreak, footnotesAtPageBottom, revisionView, editorView, header, footer, sectionsKey]);
    
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
        onActivePageChange?.(activePageInfoRef.current(layoutRef.current, activePageIndex));
      }
    }, [layout, zoom, onActivePageChange]);
    
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
      goTo: (pos) => goToRef.current?.(pos),
      tableOfContents: (action) => (editorView && editable ? runTableOfContentsRef.current?.(editorView, action) ?? false : false),
      insertNote: (kind) => {
        if (!editorView || !editable) return false;
        if (!insertNoteReference(kind)(editorView.state, editorView.dispatch)) return false;
        // The caret is just after the new reference (an empty note takes two positions).
        const placed = noteOrdinalAt(editorView.state.doc, editorView.state.selection.from - 2);
        if (placed) setNoteEdit(placed);
        return true;
      },
       
    }), [editorView, layout, editable]);
    
    // Calculate total height
    const totalHeight = layout?.totalHeight ?? pageHeight;

    // The page whose header or footer is being edited: its setup, and the header/footer content it shows (with the
    // id of the section that owns it, which is the one an edit goes to).
    const editedPage = hfEditState && layout ? layout.pages[hfEditState.pageIndex] : undefined;
    const editedPageConfig = hfEditState && layout ? getPageConfig(layout, hfEditState.pageIndex) : pageConfig;
    const owners = sectionOwnersRef.current;
    const editedSection =
      sections && owners && editedPage?.sectionIndex !== undefined
        ? {
            index: editedPage.sectionIndex,
            id: owners.ids[editedPage.sectionIndex],
            header: owners.headers[editedPage.sectionIndex],
            footer: owners.footers[editedPage.sectionIndex],
            headerOwnerId: owners.ids[owners.headerOwners[editedPage.sectionIndex]],
            footerOwnerId: owners.ids[owners.footerOwners[editedPage.sectionIndex]],
          }
        : undefined;
    
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
            ref={pagesWrapperRef}
            className="true-layout-pages-wrapper"
            style={{
              position: 'relative',
              width: layout ? Math.floor(getMaxPageWidth(layout) * zoom) : pageWidth,
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

            {/* Editor of the note being edited, over its entry in the notes area */}
            {noteEdit && noteEditPos && editorView && findNoteRef(editorView.state.doc, noteEdit.kind, noteEdit.ordinal) && (
              <NoteEditor
                key={noteKeyString(noteEdit)}
                note={findNoteRef(editorView.state.doc, noteEdit.kind, noteEdit.ordinal)!.node}
                position={noteEditPos}
                scale={zoom}
                onClose={(noteDoc) => {
                  // The text goes back into the note's reference in one step (undoable like any edit).
                  const found = findNoteRef(editorView.state.doc, noteEdit.kind, noteEdit.ordinal);
                  if (found) {
                    const text = noteTextFragment(docsSchema, noteDoc);
                    if (!found.node.content.eq(text)) {
                      editorView.dispatch(editorView.state.tr.replaceWith(found.pos + 1, found.pos + found.node.nodeSize - 1, text));
                    }
                  }
                  setNoteEdit(null);
                  inputBridgeRef.current?.focus();
                }}
              />
            )}
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

          .note-editor .ProseMirror {
            outline: none;
            white-space: pre-wrap;
            word-wrap: break-word;
            min-height: 1.5em;
          }

          .note-editor .ProseMirror p {
            margin: 0 !important;
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
            content={(hfEditState.type === 'header' ? editedSection?.header ?? header : editedSection?.footer ?? footer) ?? { blocks: [] }}
            onChange={(newContent) => {
              if (hfEditState.type === 'header') {
                onHeaderChange?.(newContent, editedSection?.headerOwnerId);
              } else {
                onFooterChange?.(newContent, editedSection?.footerOwnerId);
              }
            }}
            isEditing={hfEditState.isEditing}
            onClose={() => setHfEditState(null)}
            linkToPrevious={
              editedSection && editedSection.index > 0 && onHeaderFooterLinkChange
                ? {
                    linked: (hfEditState.type === 'header' ? owners!.headerOwners : owners!.footerOwners)[editedSection.index] !== editedSection.index,
                    onChange: (linked) =>
                      onHeaderFooterLinkChange(
                        hfEditState.type,
                        editedSection.id,
                        linked,
                        (hfEditState.type === 'header' ? editedSection.header : editedSection.footer) ?? { blocks: [] },
                      ),
                  }
                : undefined
            }
            pageIndex={hfEditState.pageIndex}
            totalPages={layout?.pages.length ?? 1}
            documentTitle={documentTitle}
            contentWidth={Math.floor((editedPageConfig.width - editedPageConfig.margins.left - editedPageConfig.margins.right) * zoom)}
            scale={zoom}
            position={(() => {
              // Calculate the position of the header/footer area
              const containerRect = pagesContainerRef.current!.getBoundingClientRect();
              const viewportRect = viewportRef.current!.getBoundingClientRect();
              const scrollTop = viewportRef.current!.scrollTop;
              
              // Calculate page Y position in the document
              const pageY = layout ? getPageY(layout, hfEditState.pageIndex) : 0;
              // Convert to viewport position
              const pageTop = containerRect.top + pageY - scrollTop + viewportRect.top - containerRect.top;
              
              const margins = {
                top: editedPageConfig.margins.top * zoom,
                bottom: editedPageConfig.margins.bottom * zoom,
                left: editedPageConfig.margins.left * zoom,
                right: editedPageConfig.margins.right * zoom,
                header: (editedPageConfig.margins.header ?? 48) * zoom,
                footer: (editedPageConfig.margins.footer ?? 48) * zoom,
              };
              
              return {
                top: hfEditState.type === 'header' ? margins.header : undefined,
                bottom: hfEditState.type === 'footer' ? margins.footer : undefined,
                left: containerRect.left + (layout ? getPageX(layout, hfEditState.pageIndex) : 0) + margins.left,
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

