/**
 * Input Bridge - Forwards events from visible pages to hidden ProseMirror editor
 * 
 * The true layout engine renders content independently on each page, but editing
 * happens in a hidden ProseMirror instance. This bridge:
 * 
 * 1. Captures keyboard events on the visible content
 * 2. Forwards them to the hidden editor
 * 3. Maps click positions to ProseMirror positions
 * 4. Handles selection via mouse drag
 */

import { EditorView } from 'prosemirror-view';
import { TextSelection } from 'prosemirror-state';
import { FlowBlock } from './flow-blocks';
import { Measure, hasLineData } from './measurer';
import { DocumentLayout, PageConfig, findPageAtY, getPageContentOrigin } from './true-layout-engine';
import { createBlockPositionMap } from './pm-to-blocks';
import { Node as PmNode } from 'prosemirror-model';
import { typeOutsideDeletion } from '@b32nio/spindle-docs-core';
import { constrainSelectionToIsolatingBlock } from './selection-constraint';

// ============================================================================
// Debug Logging
// ============================================================================

const DEBUG_SELECTION = false;  // Set to true to enable performance logging
const debugLog = (...args: unknown[]) => {
  if (DEBUG_SELECTION) {
    console.log('[INPUT-BRIDGE]', ...args);
  }
};

// ============================================================================
// Types
// ============================================================================

/**
 * Position within the visible content
 */
export interface VisiblePosition {
  /** Page index */
  pageIndex: number;
  /** X coordinate within the page content area */
  x: number;
  /** Y coordinate within the page content area */
  y: number;
}

/**
 * Callback for when selection changes
 */
export type SelectionChangeCallback = (from: number, to: number) => void;

/**
 * Callback for when focus changes
 */
export type FocusChangeCallback = (focused: boolean) => void;

/**
 * Cell selection info
 */
export interface CellSelection {
  /** Cell element */
  cellElement: HTMLElement;
  /** PM position start */
  pmStart: number;
  /** PM position end */
  pmEnd: number;
  /** Row index */
  rowIndex: number;
  /** Column index */
  colIndex: number;
}

/**
 * Callback for when cell selection changes
 */
export type CellSelectionChangeCallback = (selection: CellSelection | null) => void;

// ============================================================================
// Input Bridge Class
// ============================================================================

/**
 * InputBridge - Bridges input between visible pages and hidden editor
 */
export class InputBridge {
  private hiddenEditor: EditorView | null = null;
  private visibleMount: HTMLElement | null = null;
  private scrollContainer: HTMLElement | null = null;
  private layout: DocumentLayout | null = null;
  private blocks: FlowBlock[] = [];
  private measures: Measure[] = [];
  private blockPositionMap: Map<string, { start: number; end: number }> = new Map();
  private pageConfig: PageConfig | null = null;
  private scale: number = 1;
  
  // Event listeners (stored for cleanup)
  private boundHandlers: {
    keydown: (e: KeyboardEvent) => void;
    beforeinput: (e: InputEvent) => void;
    compositionstart: (e: CompositionEvent) => void;
    compositionupdate: (e: CompositionEvent) => void;
    compositionend: (e: CompositionEvent) => void;
    mousedown: (e: MouseEvent) => void;
    mousemove: (e: MouseEvent) => void;
    mouseup: (e: MouseEvent) => void;
    focus: (e: FocusEvent) => void;
    blur: (e: FocusEvent) => void;
    paste: (e: ClipboardEvent) => void;
    copy: (e: ClipboardEvent) => void;
    cut: (e: ClipboardEvent) => void;
  };
  
  // Mouse drag state
  private isDragging: boolean = false;
  private dragStartPos: number | null = null;
  
  // Auto-scroll state for drag selection
  private autoScrollRAF: number | null = null;
  private lastMouseY: number = 0;
  private lastMouseX: number = 0;
  
  // Deduplication state for selection updates during drag
  private lastProcessedPos: number = -1;  // Cache last position to skip redundant updates
  
  // Multi-click tracking for word/paragraph/document selection
  private clickCount: number = 0;
  private lastClickTime: number = 0;
  private lastClickX: number = 0;
  private lastClickY: number = 0;
  private readonly MULTI_CLICK_THRESHOLD = 500; // ms between clicks
  private readonly MULTI_CLICK_DISTANCE = 5; // px tolerance for multi-click
  
  // Cell selection state (for two-click table editing)
  private selectedCell: CellSelection | null = null;
  private isEditingCell: boolean = false;
  
  // Callbacks
  private onSelectionChange: SelectionChangeCallback | null = null;
  private onFocusChange: FocusChangeCallback | null = null;
  private onCellSelectionChange: CellSelectionChangeCallback | null = null;
  
  constructor() {
    // Bind handlers for visible mount events
    // Forward events from visible surface to hidden editor
    this.boundHandlers = {
      keydown: this.handleKeyDown.bind(this),
      beforeinput: this.handleBeforeInput.bind(this),
      compositionstart: this.handleComposition.bind(this),
      compositionupdate: this.handleComposition.bind(this),
      compositionend: this.handleComposition.bind(this),
      mousedown: this.handleMouseDown.bind(this),
      mousemove: this.handleMouseMove.bind(this),
      mouseup: this.handleMouseUp.bind(this),
      focus: this.handleFocus.bind(this),
      blur: this.handleBlur.bind(this),
      paste: this.handlePaste.bind(this),
      copy: this.handleCopy.bind(this),
      cut: this.handleCut.bind(this),
    };
  }
  
  /**
   * Initialize the bridge with the hidden editor and visible mount
   */
  initialize(
    hiddenEditor: EditorView,
    visibleMount: HTMLElement,
    options?: {
      onSelectionChange?: SelectionChangeCallback;
      onFocusChange?: FocusChangeCallback;
      onCellSelectionChange?: CellSelectionChangeCallback;
      scrollContainer?: HTMLElement;
    }
  ): void {
    this.hiddenEditor = hiddenEditor;
    this.visibleMount = visibleMount;
    this.scrollContainer = options?.scrollContainer ?? null;
    this.onSelectionChange = options?.onSelectionChange ?? null;
    this.onFocusChange = options?.onFocusChange ?? null;
    this.onCellSelectionChange = options?.onCellSelectionChange ?? null;
    
    // Make visible mount focusable (NOT contentEditable - that interferes with selection)
    // The hidden ProseMirror editor handles all text input directly via focus
    visibleMount.tabIndex = 0;
    visibleMount.style.outline = 'none';
    visibleMount.style.userSelect = 'none';  // Prevent native text selection
    visibleMount.style.cursor = 'text';  // Show text cursor
    visibleMount.setAttribute('data-gramm', 'false');  // Disable Grammarly
    visibleMount.setAttribute('spellcheck', 'false');  // Disable spellcheck
    
    // Add event listeners for keyboard input - forward to hidden editor
    // Note: We listen on visible mount to capture focus, then forward to hidden editor
    visibleMount.addEventListener('keydown', this.boundHandlers.keydown);
    
    // Listen for beforeinput on the WINDOW to capture text input from the hidden editor
    // when focus is there (this handles IME and other input methods)
    window.addEventListener('beforeinput', this.boundHandlers.beforeinput);
    
    // Add event listeners for IME composition
    visibleMount.addEventListener('compositionstart', this.boundHandlers.compositionstart);
    visibleMount.addEventListener('compositionupdate', this.boundHandlers.compositionupdate);
    visibleMount.addEventListener('compositionend', this.boundHandlers.compositionend);
    
    // Add event listeners for mouse interactions
    visibleMount.addEventListener('mousedown', this.boundHandlers.mousedown);
    visibleMount.addEventListener('focus', this.boundHandlers.focus);
    visibleMount.addEventListener('blur', this.boundHandlers.blur);
    
    // Add event listeners for clipboard
    visibleMount.addEventListener('paste', this.boundHandlers.paste);
    visibleMount.addEventListener('copy', this.boundHandlers.copy);
    visibleMount.addEventListener('cut', this.boundHandlers.cut);
    
    // Add document-level listeners for drag
    document.addEventListener('mousemove', this.boundHandlers.mousemove);
    document.addEventListener('mouseup', this.boundHandlers.mouseup);
  }
  
  /**
   * Update the layout data
   */
  updateLayout(
    layout: DocumentLayout,
    blocks: FlowBlock[],
    measures: Measure[],
    doc: PmNode
  ): void {
    this.layout = layout;
    this.blocks = blocks;
    this.measures = measures;
    this.pageConfig = layout.pageConfig;
    this.scale = layout.scale;
    
    // Build position map
    this.blockPositionMap = createBlockPositionMap(doc, blocks);
  }
  
  /**
   * Focus the visible mount (and hidden editor)
   */
  focus(): void {
    if (this.visibleMount) {
      this.visibleMount.focus();
    }
    if (this.hiddenEditor) {
      this.hiddenEditor.focus();
    }
  }
  
  /**
   * Check if the bridge has focus
   */
  hasFocus(): boolean {
    return document.activeElement === this.visibleMount || 
           (this.hiddenEditor?.hasFocus() ?? false);
  }
  
  // ============================================================================
  // Event Handlers
  // ============================================================================
  
  /**
   * Handle keydown events - forward to hidden editor
   * Capture on visible surface, forward to hidden editor
   */
  private handleKeyDown(e: KeyboardEvent): void {
    if (!this.hiddenEditor) return;
    
    // Skip if already handled
    if (e.defaultPrevented) return;
    
    // Skip IME composition events (handled by composition handlers)
    if (e.isComposing || e.keyCode === 229) return;
    
    // If cell is selected but not in edit mode, handle special keys
    if (this.selectedCell && !this.isEditingCell) {
      // Enter key enters edit mode
      if (e.key === 'Enter') {
        const cell = this.selectedCell;
        this.isEditingCell = true;
        this.clearCellSelectionVisual();
        // Notify that we're exiting cell selection mode (entering edit mode)
        this.onCellSelectionChange?.(null);
        // Position cursor at start of cell content
        this.setSelection(cell.pmStart + 2, cell.pmStart + 2);
        e.preventDefault();
        return;
      }
      // Escape clears selection
      if (e.key === 'Escape') {
        this.clearCellSelection();
        e.preventDefault();
        return;
      }
      // Block other text input keys (but allow modifier keys for toolbar shortcuts)
      if (!e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1) {
        e.preventDefault();
        return;
      }
    }
    
    // Create synthetic event for hidden editor
    const synthetic = new KeyboardEvent('keydown', {
      key: e.key,
      code: e.code,
      keyCode: e.keyCode,
      which: e.which,
      ctrlKey: e.ctrlKey,
      shiftKey: e.shiftKey,
      altKey: e.altKey,
      metaKey: e.metaKey,
      repeat: e.repeat,
      bubbles: true,
      cancelable: true,
    });
    
    // Dispatch to hidden editor synchronously
    this.hiddenEditor.dom.dispatchEvent(synthetic);
    
    // If hidden editor handled it, prevent default on original
    if (synthetic.defaultPrevented) {
      e.preventDefault();
    }
  }
  
  /**
   * Handle beforeinput events - directly insert text into ProseMirror
   * Instead of forwarding synthetic events, we directly create transactions
   * for more reliable cross-browser behavior
   * 
   * Note: This handler is on the window, so we need to check if the event
   * is coming from our hidden editor or somewhere else we care about.
   */
  private handleBeforeInput(e: InputEvent): void {
    if (!this.hiddenEditor) return;
    // Ignore synthetic events — this handler is on the window, so a re-dispatched
    // bubbling event would re-enter here and recurse until the stack overflows.
    if (!e.isTrusted) return;

    // Only handle events from our hidden editor
    const target = e.target as HTMLElement | null;
    if (!target || !this.hiddenEditor.dom.contains(target)) {
      return;
    }
    
    // Block text input when cell is selected but not in edit mode
    if (this.selectedCell && !this.isEditingCell) {
      e.preventDefault();
      return;
    }
    
    // Skip if already handled
    if (e.defaultPrevented) return;
    
    // Skip composition events (handled separately)
    if (e.isComposing) return;
    
    // ALWAYS prevent default to stop text from being inserted into visible mount
    e.preventDefault();
    
    const { state } = this.hiddenEditor;
    const view = this.hiddenEditor;
    
    // Handle different input types
    switch (e.inputType) {
      case 'insertText':
      case 'insertReplacementText':
        if (e.data) {
          // Insert the text at current selection (typing inside a tracked deletion is live text, not more deletion)
          const { from, to } = state.selection;
          view.dispatch(typeOutsideDeletion(state, from, to, e.data) ?? state.tr.insertText(e.data));
        }
        break;
        
      case 'insertLineBreak':
      case 'insertParagraph':
        // These are handled by ProseMirror's keymap (Enter key -> splitListItem or baseKeymap)
        // Don't handle here - let the keydown event dispatch take care of it
        // The keydown handler already dispatched Enter to ProseMirror which handles 
        // paragraph splitting, list item creation, etc. properly
        break;
        
      case 'deleteContentBackward':
        // Delete character before cursor
        if (!state.selection.empty) {
          view.dispatch(state.tr.deleteSelection());
        } else if (state.selection.from > 0) {
          view.dispatch(state.tr.delete(state.selection.from - 1, state.selection.from));
        }
        break;
        
      case 'deleteContentForward':
        // Delete character after cursor
        if (!state.selection.empty) {
          view.dispatch(state.tr.deleteSelection());
        } else if (state.selection.to < state.doc.content.size) {
          view.dispatch(state.tr.delete(state.selection.from, state.selection.from + 1));
        }
        break;
        
      case 'deleteWordBackward':
      case 'deleteWordForward':
      case 'deleteSoftLineBackward':
      case 'deleteSoftLineForward':
      case 'deleteHardLineBackward':
      case 'deleteHardLineForward':
        // For these, just delete selection or single char for now
        if (!state.selection.empty) {
          view.dispatch(state.tr.deleteSelection());
        } else if (e.inputType.includes('Backward') && state.selection.from > 0) {
          view.dispatch(state.tr.delete(state.selection.from - 1, state.selection.from));
        } else if (e.inputType.includes('Forward') && state.selection.to < state.doc.content.size) {
          view.dispatch(state.tr.delete(state.selection.from, state.selection.from + 1));
        }
        break;
        
      default:
        // Unknown/exotic input types (e.g. insertTranspose from Ctrl+T,
        // insertFromYank from Ctrl+Y, formatBold from an OS menu): we've already
        // preventDefault'd, and ProseMirror's own beforeinput handler is a no-op,
        // so there is nothing to forward. Re-dispatching a synthetic bubbling
        // event here re-entered this window listener and recursed to a
        // stack-overflow crash — just ignore the input.
        break;
    }
  }
  
  /**
   * Handle composition events (IME input) - forward to hidden editor
   */
  private handleComposition(e: CompositionEvent): void {
    if (!this.hiddenEditor) return;
    
    // Create synthetic composition event
    const synthetic = new CompositionEvent(e.type, {
      data: e.data || '',
      bubbles: true,
      cancelable: true,
    });
    
    // Dispatch to hidden editor
    this.hiddenEditor.dom.dispatchEvent(synthetic);
  }
  
  /**
   * Handle mousedown events - start selection and focus hidden editor
   * Implements:
   * - Single click: position cursor
   * - Double click: select word
   * - Triple click: select paragraph/block
   * - Quadruple click: select entire document
   * 
   * Also implements two-click behavior for table cells:
   * - First click: select the cell (visually, no text selection)
   * - Second click on same cell: enter edit mode
   */
  private handleMouseDown(e: MouseEvent): void {
    const startTime = performance.now();
    debugLog('mousedown START', { x: e.clientX, y: e.clientY });
    
    if (!this.hiddenEditor || !this.layout || !this.visibleMount) {
      debugLog('mousedown ABORT - missing deps', { 
        hasEditor: !!this.hiddenEditor, 
        hasLayout: !!this.layout, 
        hasMount: !!this.visibleMount 
      });
      return;
    }
    
    // Track multi-click (double, triple, quadruple)
    const now = performance.now();
    const timeDelta = now - this.lastClickTime;
    const distX = Math.abs(e.clientX - this.lastClickX);
    const distY = Math.abs(e.clientY - this.lastClickY);
    const isMultiClick = timeDelta < this.MULTI_CLICK_THRESHOLD && 
                         distX < this.MULTI_CLICK_DISTANCE && 
                         distY < this.MULTI_CLICK_DISTANCE;
    
    if (isMultiClick) {
      this.clickCount = (this.clickCount % 4) + 1; // Cycle 1-4
    } else {
      this.clickCount = 1;
    }
    
    this.lastClickTime = now;
    this.lastClickX = e.clientX;
    this.lastClickY = e.clientY;
    
    debugLog('Multi-click detection:', { clickCount: this.clickCount, isMultiClick, timeDelta });
    
    // Check if click is on a table cell
    const clickedCell = this.getCellAtPoint(e.clientX, e.clientY);
    
    if (clickedCell) {
      // Click is on a table cell
      if (this.selectedCell && this.isSameCell(this.selectedCell, clickedCell)) {
        // Second click on same cell - enter edit mode
        this.isEditingCell = true;
        this.clearCellSelectionVisual(); // Remove visual selection
        this.hiddenEditor.focus();
        
        // Notify that we're exiting cell selection mode (entering edit mode)
        // This tells the selection overlay to show the caret again
        this.onCellSelectionChange?.(null);
        
        // Position cursor at click location within the cell
        const pos = this.coordsToPosition(e.clientX, e.clientY);
        if (pos !== null) {
          // Apply multi-click selection within the cell
          if (this.clickCount >= 2) {
            this.handleMultiClickSelection(pos, clickedCell.pmStart, clickedCell.pmEnd);
          } else {
            this.setSelection(pos, pos);
          }
          this.dragStartPos = pos;
        }
        
        this.isDragging = true;
        e.preventDefault();
        return;
      } else {
        // First click on cell - select the cell visually (NO cursor change)
        this.selectCell(clickedCell);
        this.isEditingCell = false;
        
        // Don't change PM selection - cell commands use selectedCell.pmStart directly
        // Just focus the editor to maintain keyboard event handling
        this.hiddenEditor.focus();
        e.preventDefault();
        return;
      }
    }
    
    // Click outside table cells - clear cell selection and normal behavior
    if (this.selectedCell) {
      this.clearCellSelection();
    }
    
    // Focus the hidden editor for input handling
    this.hiddenEditor.focus();
    
    // Get position in document and set initial PM selection (for immediate cursor)
    const coordsStartTime = performance.now();
    const pos = this.coordsToPosition(e.clientX, e.clientY);
    const coordsTime = performance.now() - coordsStartTime;
    debugLog('coordsToPosition took', coordsTime.toFixed(2), 'ms, result:', pos);
    
    if (pos !== null) {
      const selStartTime = performance.now();
      
      // Handle multi-click selection
      if (this.clickCount >= 2) {
        this.handleMultiClickSelection(pos);
      } else {
        this.setSelection(pos, pos);
        this.dragStartPos = pos;
      }
      
      const selTime = performance.now() - selStartTime;
      debugLog('setSelection took', selTime.toFixed(2), 'ms');
    } else {
      debugLog('WARNING: coordsToPosition returned null!');
    }
    
    // Only enable dragging for single click (not multi-click selections)
    if (this.clickCount === 1) {
      this.isDragging = true;
    }
    
    // Store mouse position for auto-scroll
    this.lastMouseX = e.clientX;
    this.lastMouseY = e.clientY;
    
    // Prevent default to stop any native behaviors (text selection, etc.)
    // We handle selection rendering via our custom SelectionOverlay
    e.preventDefault();
    
    const totalTime = performance.now() - startTime;
    debugLog('mousedown COMPLETE in', totalTime.toFixed(2), 'ms');
  }
  
  /**
   * Handle multi-click selection (double, triple, quadruple click)
   * - Double click (2): select word
   * - Triple click (3): select paragraph/block
   * - Quadruple click (4): select entire document
   */
  private handleMultiClickSelection(pos: number, constrainStart?: number, constrainEnd?: number): void {
    if (!this.hiddenEditor) return;
    
    const { state } = this.hiddenEditor;
    const doc = state.doc;
    
    if (this.clickCount === 2) {
      // Double-click: select word
      const wordRange = this.getWordRangeAtPosition(pos);
      if (wordRange) {
        let { from, to } = wordRange;
        // Constrain to cell bounds if provided
        if (constrainStart !== undefined) from = Math.max(from, constrainStart);
        if (constrainEnd !== undefined) to = Math.min(to, constrainEnd);
        this.setSelection(from, to);
        // Set drag anchors to word boundaries for word-wise drag selection
        this.dragStartPos = from;
      } else {
        this.setSelection(pos, pos);
        this.dragStartPos = pos;
      }
    } else if (this.clickCount === 3) {
      // Triple-click: select paragraph/block
      const blockRange = this.getBlockRangeAtPosition(pos);
      if (blockRange) {
        let { from, to } = blockRange;
        // Constrain to cell bounds if provided
        if (constrainStart !== undefined) from = Math.max(from, constrainStart);
        if (constrainEnd !== undefined) to = Math.min(to, constrainEnd);
        this.setSelection(from, to);
        this.dragStartPos = from;
      } else {
        this.setSelection(pos, pos);
        this.dragStartPos = pos;
      }
    } else if (this.clickCount === 4) {
      // Quadruple-click: select entire document
      // If constrained to cell, select entire cell content instead
      if (constrainStart !== undefined && constrainEnd !== undefined) {
        this.setSelection(constrainStart + 1, constrainEnd - 1);
      } else {
        // Select entire document (from first text position to last)
        this.setSelection(1, doc.content.size - 1);
      }
      this.dragStartPos = 1;
    }
  }
  
  /**
   * Get the word range at the given ProseMirror position
   * Returns { from, to } positions that encompass the word
   */
  private getWordRangeAtPosition(pos: number): { from: number; to: number } | null {
    if (!this.hiddenEditor) return null;
    
    const { state } = this.hiddenEditor;
    const doc = state.doc;
    
    // Resolve position to get the node and offset
    const $pos = doc.resolve(pos);
    
    // Get the parent node (usually paragraph or similar text container)
    const parent = $pos.parent;
    if (!parent.isTextblock) return null;
    
    // Get the text content of the parent
    const parentText = parent.textContent;
    const parentStart = $pos.start();
    const offsetInParent = pos - parentStart;
    
    if (parentText.length === 0) return null;
    
    // Find word boundaries using regex
    // Word characters: letters, numbers, and common word chars
    const isWordChar = (char: string): boolean => {
      return /[\w\u00C0-\u024F\u1E00-\u1EFF]/.test(char);
    };
    
    // If we're not on a word character, try to find the nearest word
    const currentChar = parentText[offsetInParent] || '';
    const prevChar = offsetInParent > 0 ? parentText[offsetInParent - 1] : '';
    
    let wordStart = offsetInParent;
    let wordEnd = offsetInParent;
    
    // If current position is not on a word char, check if we're right after one
    if (!isWordChar(currentChar)) {
      if (isWordChar(prevChar)) {
        // Position is right after a word, select that word
        wordStart = offsetInParent - 1;
        wordEnd = offsetInParent;
      } else {
        // Not on a word - select whitespace or punctuation as a unit
        // Find the extent of the current non-word character sequence
        while (wordStart > 0 && !isWordChar(parentText[wordStart - 1])) {
          wordStart--;
        }
        while (wordEnd < parentText.length && !isWordChar(parentText[wordEnd])) {
          wordEnd++;
        }
        return {
          from: parentStart + wordStart,
          to: parentStart + wordEnd
        };
      }
    }
    
    // Find word start
    while (wordStart > 0 && isWordChar(parentText[wordStart - 1])) {
      wordStart--;
    }
    
    // Find word end
    while (wordEnd < parentText.length && isWordChar(parentText[wordEnd])) {
      wordEnd++;
    }
    
    // Ensure we have a valid range
    if (wordStart === wordEnd) return null;
    
    return {
      from: parentStart + wordStart,
      to: parentStart + wordEnd
    };
  }
  
  /**
   * Get the block/paragraph range at the given ProseMirror position
   * Returns { from, to } positions that encompass the block
   */
  private getBlockRangeAtPosition(pos: number): { from: number; to: number } | null {
    if (!this.hiddenEditor) return null;
    
    const { state } = this.hiddenEditor;
    const doc = state.doc;
    
    // Resolve position
    const $pos = doc.resolve(pos);
    
    // Find the closest block-level node
    // Walk up until we find a node that's a direct child of the doc or a block container
    let depth = $pos.depth;
    while (depth > 0) {
      const node = $pos.node(depth);
      const parent = $pos.node(depth - 1);
      
      // Check if this is a block node (paragraph, heading, list item, etc.)
      if (node.isBlock && parent.type.name === 'doc') {
        const start = $pos.start(depth);
        const end = $pos.end(depth);
        return { from: start, to: end };
      }
      
      // Also handle list items - select the content of the list item
      if (node.isTextblock) {
        const start = $pos.start(depth);
        const end = $pos.end(depth);
        return { from: start, to: end };
      }
      
      depth--;
    }
    
    return null;
  }
  
  /**
   * Get the table cell at the given point
   */
  private getCellAtPoint(clientX: number, clientY: number): CellSelection | null {
    const elements = document.elementsFromPoint(clientX, clientY);
    
    // First, try to find a table-cell directly in the elements
    for (const el of elements) {
      const htmlEl = el as HTMLElement;
      
      // Check if this element is a table cell
      if (htmlEl.classList.contains('table-cell') && 
          htmlEl.dataset.pmStart !== undefined && 
          htmlEl.dataset.pmEnd !== undefined) {
        return this.createCellSelection(htmlEl);
      }
      
      // Also check if this element is inside a table cell (e.g., clicking on text)
      const cellEl = htmlEl.closest('.table-cell') as HTMLElement | null;
      if (cellEl && 
          cellEl.dataset.pmStart !== undefined && 
          cellEl.dataset.pmEnd !== undefined) {
        return this.createCellSelection(cellEl);
      }
    }
    
    return null;
  }
  
  /**
   * Create a CellSelection from a cell element
   */
  private createCellSelection(cellEl: HTMLElement): CellSelection | null {
    const pmStart = parseInt(cellEl.dataset.pmStart!, 10);
    const pmEnd = parseInt(cellEl.dataset.pmEnd!, 10);
    
    if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) return null;
    
    // Find row and column indices
    const tr = cellEl.closest('tr');
    const table = cellEl.closest('table');
    
    let rowIndex = 0;
    let colIndex = 0;
    
    if (tr && table) {
      const rows = table.querySelectorAll('tr');
      rowIndex = Array.from(rows).indexOf(tr);
      
      const cells = tr.querySelectorAll('td, th');
      colIndex = Array.from(cells).indexOf(cellEl);
    }
    
    return {
      cellElement: cellEl,
      pmStart,
      pmEnd,
      rowIndex,
      colIndex,
    };
  }
  
  /**
   * Check if two cell selections refer to the same cell
   */
  private isSameCell(a: CellSelection, b: CellSelection): boolean {
    return a.pmStart === b.pmStart && a.pmEnd === b.pmEnd;
  }
  
  /**
   * Select a cell
   */
  private selectCell(cell: CellSelection): void {
    // Clear previous visual selection
    this.clearCellSelectionVisual();
    
    // Set new selection
    this.selectedCell = cell;
    
    // Apply visual selection
    this.applyCellSelectionVisual();
    
    // Notify listeners
    this.onCellSelectionChange?.(cell);
  }
  
  /**
   * Apply visual selection to the currently selected cell
   * Call this after DOM repaint to restore the visual highlight
   */
  applyCellSelectionVisual(): void {
    if (!this.selectedCell || !this.visibleMount) return;
    
    // Find the cell element by PM position data attributes
    const cellEl = this.visibleMount.querySelector(
      `.table-cell[data-pm-start="${this.selectedCell.pmStart}"][data-pm-end="${this.selectedCell.pmEnd}"]`
    ) as HTMLElement | null;
    
    if (cellEl) {
      cellEl.classList.add('cell-selected');
      // Update the cellElement reference in case DOM was rebuilt
      this.selectedCell.cellElement = cellEl;
    }
  }
  
  /**
   * Remove visual selection from cells (without clearing selection state)
   */
  private clearCellSelectionVisual(): void {
    if (!this.visibleMount) return;
    
    // Remove cell-selected class from any cell that has it
    const selectedCells = this.visibleMount.querySelectorAll('.table-cell.cell-selected');
    selectedCells.forEach(el => el.classList.remove('cell-selected'));
  }
  
  /**
   * Clear cell selection completely
   */
  clearCellSelection(): void {
    if (this.selectedCell) {
      this.clearCellSelectionVisual();
      this.selectedCell = null;
      this.isEditingCell = false;
      this.onCellSelectionChange?.(null);
    }
  }
  
  /**
   * Get current cell selection
   */
  getSelectedCell(): CellSelection | null {
    return this.selectedCell;
  }
  
  /**
   * Check if currently editing a cell
   */
  isInCellEditMode(): boolean {
    return this.isEditingCell;
  }
  
  /**
   * Refresh cell selection visual after DOM repaint
   * Call this from the layout engine after repainting pages
   */
  refreshCellSelectionVisual(): void {
    if (this.selectedCell && !this.isEditingCell) {
      this.applyCellSelectionVisual();
    }
  }
  
  /**
   * Handle mousemove events - update selection during drag
   * Uses position-based deduplication to skip redundant updates.
   */
  private handleMouseMove(e: MouseEvent): void {
    if (!this.isDragging || this.dragStartPos === null || !this.hiddenEditor) return;
    
    const moveStartTime = performance.now();
    
    // Store mouse position for auto-scroll
    this.lastMouseX = e.clientX;
    this.lastMouseY = e.clientY;
    
    // Get position from coordinates
    const coordsStartTime = performance.now();
    const pos = this.coordsToPositionWithClamp(e.clientX, e.clientY);
    const coordsTime = performance.now() - coordsStartTime;
    
    if (pos === null) {
      debugLog('mousemove: coordsToPosition returned null in', coordsTime.toFixed(2), 'ms');
      return;
    }
    
    // Skip if position hasn't changed (common during slow mouse movement)
    if (pos === this.lastProcessedPos) {
      return;  // Silent skip - this is expected
    }
    this.lastProcessedPos = pos;
    
    // Only update PM state if selection actually changed
    const { state } = this.hiddenEditor;
    const from = Math.min(this.dragStartPos, pos);
    const to = Math.max(this.dragStartPos, pos);
    
    if (state.selection.from !== from || state.selection.to !== to) {
      const selStartTime = performance.now();
      this.setSelectionImmediate(this.dragStartPos, pos);
      const selTime = performance.now() - selStartTime;
      
      const totalTime = performance.now() - moveStartTime;
      debugLog('mousemove: pos=' + pos + ', coordsTime=' + coordsTime.toFixed(2) + 'ms, selTime=' + selTime.toFixed(2) + 'ms, total=' + totalTime.toFixed(2) + 'ms');
    }
    
    // Start auto-scroll if mouse is near edges
    this.startAutoScroll();
  }
  
  /**
   * Get position with clamping for coordinates outside the visible area
   */
  private coordsToPositionWithClamp(clientX: number, clientY: number): number | null {
    if (!this.visibleMount || !this.layout || !this.hiddenEditor) return null;
    
    const mountRect = this.visibleMount.getBoundingClientRect();
    
    // Clamp X coordinate to mount bounds for position calculation
    const clampedX = Math.max(mountRect.left, Math.min(mountRect.right, clientX));
    
    // If mouse is above the visible area, get position at top
    if (clientY < mountRect.top) {
      // Find the topmost visible position
      const pos = this.coordsToPosition(clampedX, mountRect.top + 5);
      if (pos !== null) return pos;
    }
    
    // If mouse is below the visible area, get position at bottom
    if (clientY > mountRect.bottom) {
      // Find the bottommost visible position
      const pos = this.coordsToPosition(clampedX, mountRect.bottom - 5);
      if (pos !== null) return pos;
    }
    
    // Normal case - mouse is within bounds
    return this.coordsToPosition(clientX, clientY);
  }
  
  /**
   * Start auto-scrolling when dragging near edges
   */
  private startAutoScroll(): void {
    if (this.autoScrollRAF !== null) return; // Already running
    
    this.autoScrollRAF = requestAnimationFrame(() => this.performAutoScroll());
  }
  
  /**
   * Stop auto-scrolling
   */
  private stopAutoScroll(): void {
    if (this.autoScrollRAF !== null) {
      cancelAnimationFrame(this.autoScrollRAF);
      this.autoScrollRAF = null;
    }
  }
  
  /**
   * Perform auto-scroll based on mouse position
   */
  private performAutoScroll(): void {
    this.autoScrollRAF = null;
    
    if (!this.isDragging || !this.scrollContainer || this.dragStartPos === null) return;
    
    const containerRect = this.scrollContainer.getBoundingClientRect();
    const edgeThreshold = 50; // Pixels from edge to trigger scroll
    const maxScrollSpeed = 20; // Max pixels per frame
    
    let scrollDelta = 0;
    
    // Check if mouse is near top edge
    if (this.lastMouseY < containerRect.top + edgeThreshold) {
      const distance = containerRect.top + edgeThreshold - this.lastMouseY;
      scrollDelta = -Math.min(maxScrollSpeed, distance * 0.5);
    }
    // Check if mouse is near bottom edge
    else if (this.lastMouseY > containerRect.bottom - edgeThreshold) {
      const distance = this.lastMouseY - (containerRect.bottom - edgeThreshold);
      scrollDelta = Math.min(maxScrollSpeed, distance * 0.5);
    }
    
    if (scrollDelta !== 0) {
      this.scrollContainer.scrollTop += scrollDelta;
      
      // Update selection after scroll
      const pos = this.coordsToPositionWithClamp(this.lastMouseX, this.lastMouseY);
      if (pos !== null) {
        this.setSelectionImmediate(this.dragStartPos, pos);
      }
      
      // Continue auto-scrolling
      this.autoScrollRAF = requestAnimationFrame(() => this.performAutoScroll());
    }
  }
  
  /**
   * Handle mouseup events - end selection
   */
  private handleMouseUp(_e: MouseEvent): void {
    this.isDragging = false;
    this.dragStartPos = null;
    this.lastProcessedPos = -1;
    this.stopAutoScroll();
  }
  
  /**
   * Handle focus events
   */
  private handleFocus(_e: FocusEvent): void {
    if (this.hiddenEditor) {
      this.hiddenEditor.focus();
    }
    this.onFocusChange?.(true);
  }
  
  /**
   * Handle blur events
   */
  private handleBlur(e: FocusEvent): void {
    // Check if focus is moving to the hidden editor (which is fine)
    const relatedTarget = e.relatedTarget as HTMLElement | null;
    if (relatedTarget && this.hiddenEditor?.dom.contains(relatedTarget)) {
      return;
    }
    this.onFocusChange?.(false);
  }
  
  /**
   * Handle paste events
   */
  private handlePaste(e: ClipboardEvent): void {
    if (!this.hiddenEditor) return;
    
    // Forward paste to hidden editor
    const syntheticEvent = new ClipboardEvent('paste', {
      clipboardData: e.clipboardData,
      bubbles: true,
      cancelable: true,
    });
    
    this.hiddenEditor.dom.dispatchEvent(syntheticEvent);
    
    if (syntheticEvent.defaultPrevented) {
      e.preventDefault();
    }
  }
  
  /**
   * Handle copy events
   */
  private handleCopy(e: ClipboardEvent): void {
    if (!this.hiddenEditor) return;
    
    const syntheticEvent = new ClipboardEvent('copy', {
      clipboardData: e.clipboardData,
      bubbles: true,
      cancelable: true,
    });
    
    this.hiddenEditor.dom.dispatchEvent(syntheticEvent);
    
    if (syntheticEvent.defaultPrevented) {
      e.preventDefault();
    }
  }
  
  /**
   * Handle cut events
   */
  private handleCut(e: ClipboardEvent): void {
    if (!this.hiddenEditor) return;
    
    const syntheticEvent = new ClipboardEvent('cut', {
      clipboardData: e.clipboardData,
      bubbles: true,
      cancelable: true,
    });
    
    this.hiddenEditor.dom.dispatchEvent(syntheticEvent);
    
    if (syntheticEvent.defaultPrevented) {
      e.preventDefault();
    }
  }
  
  // ============================================================================
  // Position Mapping
  // ============================================================================
  
  /**
   * Convert screen coordinates to ProseMirror position using DOM data attributes.
   * Uses pixel-perfect click-to-position mapping via data-pm-start/data-pm-end attributes.
   */
  private coordsToPosition(clientX: number, clientY: number): number | null {
    if (!this.visibleMount) return null;
    
    // Try DOM-based mapping first (more accurate)
    const domStartTime = performance.now();
    const domPos = this.clickToPositionDom(clientX, clientY);
    const domTime = performance.now() - domStartTime;
    
    if (domPos !== null) {
      debugLog('coordsToPosition: DOM mapping succeeded in', domTime.toFixed(2), 'ms, pos:', domPos);
      return domPos;
    }
    
    // Fallback to geometry-based mapping
    const geoStartTime = performance.now();
    const geoPos = this.clickToPositionGeometry(clientX, clientY);
    const geoTime = performance.now() - geoStartTime;
    debugLog('coordsToPosition: DOM failed in', domTime.toFixed(2), 'ms, geo fallback in', geoTime.toFixed(2), 'ms, pos:', geoPos);
    return geoPos;
  }
  
  /**
   * DOM-based click-to-position mapping using data-pm-start/data-pm-end attributes.
   * Uses elementsFromPoint and binary search with Range API for exact character positioning.
   */
  private clickToPositionDom(clientX: number, clientY: number): number | null {
    // Use elementsFromPoint to find elements under the click
    const elementsStartTime = performance.now();
    const hitChain = document.elementsFromPoint(clientX, clientY);
    const elementsTime = performance.now() - elementsStartTime;
    
    if (!hitChain || hitChain.length === 0) {
      debugLog('clickToPositionDom: elementsFromPoint returned empty in', elementsTime.toFixed(2), 'ms');
      return null;
    }
    
    debugLog('clickToPositionDom: elementsFromPoint found', hitChain.length, 'elements in', elementsTime.toFixed(2), 'ms');
    
    // Find a span, line, cell, or fragment element with PM position data
    let targetSpan: HTMLElement | null = null;
    let targetLine: HTMLElement | null = null;
    let targetFragment: HTMLElement | null = null;
    let targetCellParagraph: HTMLElement | null = null;
    let targetTableCell: HTMLElement | null = null;
    
    for (const el of hitChain) {
      const htmlEl = el as HTMLElement;
      
      // Check for span with PM positions (most specific)
      if ((htmlEl.tagName === 'SPAN' || htmlEl.tagName === 'A') && 
          htmlEl.dataset.pmStart !== undefined && 
          htmlEl.dataset.pmEnd !== undefined) {
        targetSpan = htmlEl;
        break;
      }
      
      // Check for line element
      if (htmlEl.classList.contains('line') && 
          htmlEl.dataset.pmStart !== undefined && 
          htmlEl.dataset.pmEnd !== undefined) {
        targetLine = htmlEl;
      }
      
      // Check for cell paragraph (table cell content)
      if (htmlEl.classList.contains('cell-paragraph') && 
          htmlEl.dataset.pmStart !== undefined && 
          htmlEl.dataset.pmEnd !== undefined) {
        targetCellParagraph = htmlEl;
      }
      
      // Check for table cell
      if (htmlEl.classList.contains('table-cell') && 
          htmlEl.dataset.pmStart !== undefined && 
          htmlEl.dataset.pmEnd !== undefined) {
        targetTableCell = htmlEl;
      }
      
      // Check for fragment element
      if (htmlEl.classList.contains('fragment') && 
          htmlEl.dataset.pmStart !== undefined && 
          htmlEl.dataset.pmEnd !== undefined) {
        targetFragment = htmlEl;
      }
    }
    
    // Process the most specific element found
    if (targetSpan) {
      return this.processSpanClick(targetSpan, clientX);
    }
    
    if (targetLine) {
      return this.processLineClick(targetLine, clientX);
    }
    
    // Handle table cell paragraph clicks
    if (targetCellParagraph) {
      return this.processCellParagraphClick(targetCellParagraph, clientX);
    }
    
    // Handle table cell clicks (fallback to cell start if no paragraph found)
    if (targetTableCell) {
      return this.processTableCellClick(targetTableCell, clientX, clientY);
    }
    
    if (targetFragment) {
      // Try to find a line or span within the fragment
      return this.processFragmentClick(targetFragment, clientX, clientY);
    }
    
    return null;
  }
  
  /**
   * Process a click on a table cell paragraph
   */
  private processCellParagraphClick(para: HTMLElement, clientX: number): number | null {
    const pmStart = parseInt(para.dataset.pmStart || '', 10);
    const pmEnd = parseInt(para.dataset.pmEnd || '', 10);
    
    if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) {
      return null;
    }
    
    // Find spans within the cell paragraph
    const spans = Array.from(para.querySelectorAll('span[data-pm-start]')) as HTMLElement[];
    
    if (spans.length === 0) {
      // Empty paragraph - return start position (inside the paragraph)
      return pmStart + 1;
    }
    
    // Find the span at the X coordinate
    for (const span of spans) {
      const rect = span.getBoundingClientRect();
      if (clientX >= rect.left && clientX <= rect.right) {
        return this.processSpanClick(span, clientX);
      }
    }
    
    // Check if click is before first span
    const firstRect = spans[0].getBoundingClientRect();
    if (clientX <= firstRect.left) {
      return pmStart + 1;
    }
    
    // Click is after last span
    return pmEnd - 1;
  }
  
  /**
   * Process a click on a table cell
   */
  private processTableCellClick(cell: HTMLElement, clientX: number, clientY: number): number | null {
    const pmStart = parseInt(cell.dataset.pmStart || '', 10);
    const pmEnd = parseInt(cell.dataset.pmEnd || '', 10);
    
    if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) {
      return null;
    }
    
    // Try to find a cell paragraph at the click position
    const paragraphs = Array.from(cell.querySelectorAll('.cell-paragraph[data-pm-start]')) as HTMLElement[];
    
    for (const para of paragraphs) {
      const rect = para.getBoundingClientRect();
      if (clientY >= rect.top && clientY <= rect.bottom) {
        return this.processCellParagraphClick(para, clientX);
      }
    }
    
    // Fallback to cell start position
    return pmStart + 1;
  }
  
  /**
   * Process a click on a span element with PM position data.
   * Uses binary search with Range API for exact character positioning.
   */
  private processSpanClick(span: HTMLElement, clientX: number): number | null {
    const pmStart = parseInt(span.dataset.pmStart || '', 10);
    const pmEnd = parseInt(span.dataset.pmEnd || '', 10);
    
    if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) {
      return null;
    }
    
    // Get the text node
    const textNode = span.firstChild;
    if (!textNode || textNode.nodeType !== Node.TEXT_NODE || !textNode.textContent) {
      // Empty span - choose closer edge
      const rect = span.getBoundingClientRect();
      return Math.abs(clientX - rect.left) <= Math.abs(clientX - rect.right) ? pmStart : pmEnd;
    }
    
    // Binary search for exact character position
    const charIndex = this.findCharIndexAtX(textNode as Text, span, clientX);
    const textLength = textNode.textContent.length;
    
    // Map character index to PM position
    if (textLength === 0) return pmStart;
    
    const pmRange = pmEnd - pmStart;
    if (pmRange === textLength) {
      return pmStart + charIndex;
    }
    
    // If PM range doesn't match text length, use ratio
    const ratio = charIndex / textLength;
    return Math.round(pmStart + ratio * pmRange);
  }
  
  /**
   * Process a click on a line element.
   */
  private processLineClick(line: HTMLElement, clientX: number): number | null {
    const pmStart = parseInt(line.dataset.pmStart || '', 10);
    const pmEnd = parseInt(line.dataset.pmEnd || '', 10);
    
    if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) {
      return null;
    }
    
    // Find spans within the line
    const spans = Array.from(line.querySelectorAll('span[data-pm-start], a[data-pm-start]')) as HTMLElement[];
    
    if (spans.length === 0) {
      return pmStart;
    }
    
    // Check if click is before first span
    const firstRect = spans[0].getBoundingClientRect();
    if (clientX <= firstRect.left) {
      return pmStart;
    }
    
    // Check if click is after last span
    const lastRect = spans[spans.length - 1].getBoundingClientRect();
    if (clientX >= lastRect.right) {
      return pmEnd;
    }
    
    // Find the span containing the X coordinate
    for (const span of spans) {
      const rect = span.getBoundingClientRect();
      if (clientX >= rect.left && clientX <= rect.right) {
        return this.processSpanClick(span, clientX);
      }
    }
    
    return pmStart;
  }
  
  /**
   * Process a click on a fragment element.
   */
  private processFragmentClick(fragment: HTMLElement, clientX: number, clientY: number): number | null {
    const pmStart = parseInt(fragment.dataset.pmStart || '', 10);
    const pmEnd = parseInt(fragment.dataset.pmEnd || '', 10);
    
    if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) {
      return null;
    }
    
    // Find lines within the fragment
    const lines = Array.from(fragment.querySelectorAll('.line[data-pm-start]')) as HTMLElement[];
    
    if (lines.length === 0) {
      // No line structure, try to find spans directly
      const spans = Array.from(fragment.querySelectorAll('span[data-pm-start], a[data-pm-start]')) as HTMLElement[];
      if (spans.length > 0) {
        // Find span at Y position
        for (const span of spans) {
          const rect = span.getBoundingClientRect();
          if (clientY >= rect.top && clientY <= rect.bottom) {
            return this.processSpanClick(span, clientX);
          }
        }
      }
      return pmStart;
    }
    
    // Find the line at Y position
    for (const line of lines) {
      const rect = line.getBoundingClientRect();
      if (clientY >= rect.top && clientY <= rect.bottom) {
        return this.processLineClick(line, clientX);
      }
    }
    
    // If click is below all lines, return end of last line
    const lastLine = lines[lines.length - 1];
    const lastLineEnd = parseInt(lastLine.dataset.pmEnd || '', 10);
    return Number.isFinite(lastLineEnd) ? lastLineEnd : pmEnd;
  }
  
  /**
   * Binary search to find character index at X coordinate within a text node.
   * Uses Range API for accurate measurement.
   */
  private findCharIndexAtX(textNode: Text, container: HTMLElement, targetX: number): number {
    const text = textNode.textContent || '';
    if (text.length === 0) return 0;
    
    const baseLeft = container.getBoundingClientRect().left;
    const range = document.createRange();
    
    // Binary search
    let lo = 0;
    let hi = text.length;
    
    while (lo < hi) {
      const mid = Math.floor((lo + hi) / 2);
      range.setStart(textNode, 0);
      range.setEnd(textNode, mid);
      const w = range.getBoundingClientRect().width;
      const x = baseLeft + w;
      
      if (x < targetX) {
        lo = mid + 1;
      } else {
        hi = mid;
      }
    }
    
    // Find nearest boundary
    const index = Math.max(0, Math.min(text.length, lo));
    
    const measureAt = (i: number): number => {
      range.setStart(textNode, 0);
      range.setEnd(textNode, i);
      return baseLeft + range.getBoundingClientRect().width;
    };
    
    const xAt = measureAt(index);
    const distAt = Math.abs(xAt - targetX);
    
    // Check if previous boundary is closer
    if (index > 0) {
      const xPrev = measureAt(index - 1);
      const distPrev = Math.abs(xPrev - targetX);
      if (distPrev < distAt) {
        return index - 1;
      }
    }
    
    return index;
  }
  
  /**
   * Fallback geometry-based click-to-position mapping.
   * Used when DOM data attributes are not available.
   */
  private clickToPositionGeometry(clientX: number, clientY: number): number | null {
    if (!this.layout || !this.visibleMount || !this.pageConfig) return null;
    
    // Get mount rect
    const mountRect = this.visibleMount.getBoundingClientRect();
    
    // Convert to coordinates relative to mount
    const relX = clientX - mountRect.left;
    const relY = clientY - mountRect.top;
    
    // Find which page
    const pageIndex = findPageAtY(this.layout, relY);
    const page = this.layout.pages[pageIndex];
    if (!page) return null;
    
    // Convert to coordinates within page content area
    const origin = getPageContentOrigin(this.layout, pageIndex);
    const contentX = relX - origin.x;
    const contentY = relY - origin.y;
    
    // Find which fragment contains this position
    for (const fragment of page.fragments) {
      if (contentY >= fragment.y && contentY < fragment.y + fragment.height) {
        // Found the fragment - map to ProseMirror position
        return this.fragmentToPosition(fragment, contentX, contentY - fragment.y);
      }
    }
    
    // Click is outside all fragments - find the best position
    if (page.fragments.length > 0) {
      // Find the fragment closest to the click position
      let closestFragmentAbove: { fragment: typeof page.fragments[0]; distance: number } | null = null;
      let closestFragmentBelow: { fragment: typeof page.fragments[0]; distance: number } | null = null;
      
      for (const fragment of page.fragments) {
        const fragmentBottom = fragment.y + fragment.height;
        
        if (contentY >= fragmentBottom) {
          // Click is below this fragment
          const distance = contentY - fragmentBottom;
          if (!closestFragmentAbove || distance < closestFragmentAbove.distance) {
            closestFragmentAbove = { fragment, distance };
          }
        } else if (contentY < fragment.y) {
          // Click is above this fragment  
          const distance = fragment.y - contentY;
          if (!closestFragmentBelow || distance < closestFragmentBelow.distance) {
            closestFragmentBelow = { fragment, distance };
          }
        }
      }
      
      // If click is below a fragment and there's a fragment below, position at start of next block
      if (closestFragmentAbove && closestFragmentBelow) {
        // Click is between two fragments - position at start of the lower one
        const posInfo = this.blockPositionMap.get(closestFragmentBelow.fragment.blockId);
        if (posInfo) {
          // +1 to enter the block node
          return posInfo.start + 1;
        }
      }
      
      // If click is below all fragments
      if (closestFragmentAbove && !closestFragmentBelow) {
        const closestBlockIndex = this.blocks.findIndex(b => b.id === closestFragmentAbove!.fragment.blockId);
        
        // Check if there's a next block after this one
        if (closestBlockIndex >= 0 && closestBlockIndex < this.blocks.length - 1) {
          const nextBlock = this.blocks[closestBlockIndex + 1];
          const nextPosInfo = this.blockPositionMap.get(nextBlock.id);
          if (nextPosInfo) {
            // Position at start of next block's content
            return nextPosInfo.start + 1;
          }
        }
        
        // No next block - position at end of closest block
        const posInfo = this.blockPositionMap.get(closestFragmentAbove.fragment.blockId);
        if (posInfo) {
          return posInfo.end;
        }
      }
      
      // If click is above all fragments, position at start of first fragment
      if (closestFragmentBelow && !closestFragmentAbove) {
        const posInfo = this.blockPositionMap.get(closestFragmentBelow.fragment.blockId);
        if (posInfo) {
          return posInfo.start + 1;
        }
      }
    }
    
    return null;
  }
  
  /**
   * Map a position within a fragment to a ProseMirror position
   */
  private fragmentToPosition(
    fragment: { blockId: string; fromLine: number; toLine: number; height: number },
    xWithinContent: number,
    yWithinFragment: number
  ): number | null {
    const posInfo = this.blockPositionMap.get(fragment.blockId);
    if (!posInfo) return null;
    
    // Find the block and measure
    const blockIndex = this.blocks.findIndex(b => b.id === fragment.blockId);
    if (blockIndex === -1) return null;
    
    const block = this.blocks[blockIndex];
    const measure = this.measures[blockIndex];
    
    // Get the text content of the block for character-level positioning
    let blockText = '';
    if (block.kind === 'paragraph' || block.kind === 'heading') {
      blockText = block.runs.map(r => r.kind === 'text' ? r.text : '').join('');
    } else if (block.kind === 'listItem') {
      blockText = block.runs.map(r => r.kind === 'text' ? r.text : '').join('');
    }
    
    // Text content size
    const textSize = blockText.length;
    if (textSize === 0) {
      return posInfo.start + 1; // Position inside empty block
    }
    
    const hasLines = hasLineData(measure) && measure.lines.length > 0;
    const lines = hasLines ? measure.lines : [];
    const totalLines = Math.max(1, lines.length);
    
    // Calculate total characters from segments
    let totalCharsFromSegments = 0;
    for (const line of lines) {
      if (line.segments) {
        for (const seg of line.segments) {
          totalCharsFromSegments += seg.text.length;
        }
      }
    }
    
    // Use segments if available, otherwise use block text length
    const useSegments = totalCharsFromSegments > 0;
    
    // Find which line within the fragment based on Y
    let accumulatedHeight = 0;
    let targetLineIndex = fragment.fromLine;
    
    for (let i = fragment.fromLine; i < fragment.toLine && i < totalLines; i++) {
      const lineHeight = lines[i] ? lines[i].height * this.scale : (fragment.height / (fragment.toLine - fragment.fromLine));
      if (accumulatedHeight + lineHeight > yWithinFragment) {
        targetLineIndex = i;
        break;
      }
      accumulatedHeight += lineHeight;
      targetLineIndex = i;
    }
    
    // Calculate character offset
    let charOffset = 0;
    
    if (useSegments) {
      // Add characters from all lines before the target line
      for (let i = 0; i < targetLineIndex && i < totalLines; i++) {
        const line = lines[i];
        if (line && line.segments) {
          for (const seg of line.segments) {
            charOffset += seg.text.length;
          }
        }
      }
      
      // Estimate position within the target line based on X coordinate
      const targetLine = lines[targetLineIndex];
      if (targetLine && targetLine.segments && targetLine.segments.length > 0) {
        let lineTextLength = 0;
        for (const seg of targetLine.segments) {
          lineTextLength += seg.text.length;
        }
        
        // Use X position to estimate character within line
        // Note: line.width is already in pixels from measurement
        const lineWidth = targetLine.width;
        if (lineWidth > 0 && lineTextLength > 0) {
          const xRatio = Math.max(0, Math.min(1, xWithinContent / lineWidth));
          charOffset += Math.floor(xRatio * lineTextLength);
        }
      }
    } else {
      // Fallback: distribute characters evenly across lines
      const charsPerLine = Math.ceil(textSize / totalLines);
      
      // Characters from lines before target
      charOffset = targetLineIndex * charsPerLine;
      
      // Estimate position within target line based on X
      const lineWidth = lines[targetLineIndex]?.width || (fragment.height > 0 ? 500 : 0); // fallback width
      if (lineWidth > 0) {
        const xRatio = Math.max(0, Math.min(1, xWithinContent / lineWidth));
        charOffset += Math.floor(xRatio * charsPerLine);
      }
    }
    
    // Convert character offset to ProseMirror position
    // ProseMirror positions: start + 1 (enter node) + charOffset
    const pmOffset = Math.min(charOffset, textSize);
    return posInfo.start + 1 + pmOffset;
  }
  
  /**
   * Set selection in the hidden editor
   */
  /**
   * Keep a selection from crossing an isolating boundary (a table or table cell).
   * A raw TextSelection with one endpoint inside a cell and the other outside —
   * or in a different cell — deletes/merges rows and cells on the next edit
   * (there's no prosemirror-tables CellSelection installed). Clamp the head
   * (`to`) into the anchor's (`anchor`) isolating node, or out of the head's, so
   * both endpoints share the same cell/block context.
   */
  private constrainSelectionEnd(doc: PmNode, anchor: number, to: number): number {
    return constrainSelectionToIsolatingBlock(doc, anchor, to);
  }

  private setSelection(from: number, to: number): void {
    if (!this.hiddenEditor) return;

    try {
      const { state } = this.hiddenEditor;
      const docSize = state.doc.content.size;

      // Clamp positions to valid range, and constrain the head to the anchor's
      // isolating context so a cross-cell/table selection can't corrupt on edit.
      const clampedFrom = Math.max(0, Math.min(from, docSize));
      const constrainedTo = this.constrainSelectionEnd(state.doc, clampedFrom, Math.max(0, Math.min(to, docSize)));
      const clampedTo = Math.max(0, Math.min(constrainedTo, docSize));

      // Create and dispatch selection
      const selection = TextSelection.create(
        state.doc,
        Math.min(clampedFrom, clampedTo),
        Math.max(clampedFrom, clampedTo)
      );
      
      const tr = state.tr.setSelection(selection);
      this.hiddenEditor.dispatch(tr);
      
      // Notify callback
      this.onSelectionChange?.(selection.from, selection.to);
    } catch (e) {
      // Ignore selection errors
      console.warn('Selection error:', e);
    }
  }
  
  /**
   * Set selection with immediate visual update (no RAF delay)
   * Used during drag for maximum responsiveness
   */
  private setSelectionImmediate(from: number, to: number): void {
    if (!this.hiddenEditor) return;
    
    try {
      const { state } = this.hiddenEditor;
      const docSize = state.doc.content.size;

      // Clamp to valid range, and constrain the head to the anchor's isolating
      // context (prevents a cross-cell/table drag from corrupting on edit).
      const clampedFrom = Math.max(0, Math.min(from, docSize));
      const constrainedTo = this.constrainSelectionEnd(state.doc, clampedFrom, Math.max(0, Math.min(to, docSize)));
      const clampedTo = Math.max(0, Math.min(constrainedTo, docSize));

      // Skip if selection hasn't changed
      if (state.selection.from === Math.min(clampedFrom, clampedTo) &&
          state.selection.to === Math.max(clampedFrom, clampedTo)) {
        return;
      }
      
      // Create and dispatch selection with a meta flag for immediate rendering
      const selection = TextSelection.create(
        state.doc,
        Math.min(clampedFrom, clampedTo),
        Math.max(clampedFrom, clampedTo)
      );
      
      // Set immediateSelection meta to bypass RAF batching during drag
      const tr = state.tr.setSelection(selection).setMeta('immediateSelection', true);
      this.hiddenEditor.dispatch(tr);
      
      // DON'T call onSelectionChange here - let dispatchTransaction handle it
      // to avoid duplicate updates
    } catch (e) {
      // Ignore selection errors
      console.warn('Selection error:', e);
    }
  }
  
  /**
   * Destroy the bridge and remove all event listeners
   */
  destroy(): void {
    if (this.visibleMount) {
      // Remove keyboard listeners
      this.visibleMount.removeEventListener('keydown', this.boundHandlers.keydown);
      
      // Remove composition listeners
      this.visibleMount.removeEventListener('compositionstart', this.boundHandlers.compositionstart);
      this.visibleMount.removeEventListener('compositionupdate', this.boundHandlers.compositionupdate);
      this.visibleMount.removeEventListener('compositionend', this.boundHandlers.compositionend);
      
      // Remove mouse and focus listeners
      this.visibleMount.removeEventListener('mousedown', this.boundHandlers.mousedown);
      this.visibleMount.removeEventListener('focus', this.boundHandlers.focus);
      this.visibleMount.removeEventListener('blur', this.boundHandlers.blur);
      
      // Remove clipboard listeners
      this.visibleMount.removeEventListener('paste', this.boundHandlers.paste);
      this.visibleMount.removeEventListener('copy', this.boundHandlers.copy);
      this.visibleMount.removeEventListener('cut', this.boundHandlers.cut);
    }
    
    // Remove window-level listeners
    window.removeEventListener('beforeinput', this.boundHandlers.beforeinput);
    
    // Remove document-level listeners
    document.removeEventListener('mousemove', this.boundHandlers.mousemove);
    document.removeEventListener('mouseup', this.boundHandlers.mouseup);
    
    // Cancel any pending animations
    this.stopAutoScroll();
    
    // Clear references
    this.hiddenEditor = null;
    this.visibleMount = null;
    this.scrollContainer = null;
    this.layout = null;
    this.blocks = [];
    this.measures = [];
    this.blockPositionMap.clear();
    this.pageConfig = null;
  }
}

// ============================================================================
// Singleton Instance
// ============================================================================

let inputBridgeInstance: InputBridge | null = null;

/**
 * Get the singleton input bridge instance
 */
export function getInputBridge(): InputBridge {
  if (!inputBridgeInstance) {
    inputBridgeInstance = new InputBridge();
  }
  return inputBridgeInstance;
}

/**
 * Create a new input bridge instance (for testing or multiple editors)
 */
export function createInputBridge(): InputBridge {
  return new InputBridge();
}

