/**
 * Table Interactions - Handles table resizing, dragging, and context menus
 * 
 * This module provides interactive features for tables in the True Layout Engine:
 * 1. Column width resizing via drag handles
 * 2. Row height resizing via drag handles
 * 3. Drag handle for moving the entire table
 * 4. Right-click context menu for adding/removing rows and columns
 */

import { EditorView } from 'prosemirror-view';
import { Node as PmNode } from 'prosemirror-model';
import {
  applyInsertColumn,
  applyDeleteColumn,
  applyInsertRow,
  applyDeleteRow,
  applyColumnResize,
  logicalColumnForCell,
} from './pm-table-columns';

// ============================================================================
// Types
// ============================================================================

export interface TablePosition {
  tableEl: HTMLTableElement;
  pmStart: number;
  pmEnd: number;
}

export interface ResizeState {
  type: 'column' | 'row';
  tableEl: HTMLTableElement;
  index: number;
  startPos: number;
  startSize: number;
  pmStart: number;
}

export interface ContextMenuAction {
  label: string;
  icon?: string;
  action: () => void;
}

// ============================================================================
// Table Interaction Manager
// ============================================================================

export class TableInteractionManager {
  private container: HTMLElement | null = null;
  private editorView: EditorView | null = null;
  private scale: number = 1;
  
  // Resize state
  private resizeState: ResizeState | null = null;
  private resizeHandles: HTMLElement[] = [];
  private dragHandle: HTMLElement | null = null;
  private currentTableEl: HTMLTableElement | null = null;
  private isHoveringHandle: boolean = false;
  private hideTimeout: ReturnType<typeof setTimeout> | null = null;
  
  // Drag state
  private isDragging: boolean = false;
  private dragStartY: number = 0;
  private dragTableEl: HTMLTableElement | null = null;
  private dragGhost: HTMLElement | null = null;
  private dropIndicator: HTMLElement | null = null;
  private dropTargetPos: number | null = null;
  
  // Context menu
  private contextMenu: HTMLElement | null = null;
  private contextMenuTarget: TablePosition | null = null;
  
  // Bound event handlers
  private boundMouseMove: (e: MouseEvent) => void;
  private boundMouseUp: (e: MouseEvent) => void;
  private boundContextMenu: (e: MouseEvent) => void;
  private boundClick: (e: MouseEvent) => void;
  
  // Callbacks
  private onTableUpdate: (() => void) | null = null;
  
  constructor() {
    this.boundMouseMove = this.handleMouseMove.bind(this);
    this.boundMouseUp = this.handleMouseUp.bind(this);
    this.boundContextMenu = this.handleContextMenu.bind(this);
    this.boundClick = this.handleClick.bind(this);
  }
  
  /**
   * Initialize the table interaction manager
   */
  initialize(
    container: HTMLElement,
    editorView: EditorView,
    options?: {
      onTableUpdate?: () => void;
      scale?: number;
    }
  ): void {
    this.container = container;
    this.editorView = editorView;
    this.scale = options?.scale ?? 1;
    this.onTableUpdate = options?.onTableUpdate ?? null;
    
    // Add event listeners
    container.addEventListener('contextmenu', this.boundContextMenu);
    document.addEventListener('click', this.boundClick);
    document.addEventListener('mousemove', this.boundMouseMove);
    document.addEventListener('mouseup', this.boundMouseUp);
    
    // Create resize handles container
    this.createResizeHandles();
    
    // Create context menu element
    this.createContextMenu();
  }
  
  /**
   * Update after layout changes - refresh resize handles for visible tables
   */
  updateLayout(): void {
    this.hideResizeHandles();
    this.hideDragHandle();
    // Handles will be shown on hover
  }
  
  /**
   * Update the scale factor (call when zoom changes)
   */
  setScale(scale: number): void {
    this.scale = scale;
  }
  
  /**
   * Show resize handles for a table on hover
   */
  showHandlesForTable(tableEl: HTMLTableElement): void {
    this.clearResizeHandles();
    this.createColumnResizeHandles(tableEl);
    this.createRowResizeHandles(tableEl);
    this.showDragHandle(tableEl);
  }
  
  /**
   * Hide all resize handles
   */
  hideResizeHandles(): void {
    this.resizeHandles.forEach(h => h.remove());
    this.resizeHandles = [];
  }
  
  /**
   * Hide drag handle
   */
  hideDragHandle(): void {
    if (this.dragHandle) {
      this.dragHandle.style.display = 'none';
    }
    this.currentTableEl = null;
  }
  
  /**
   * Check if currently hovering over any handle
   */
  isHoveringOverHandles(): boolean {
    return this.isHoveringHandle;
  }
  
  /**
   * Schedule hiding of handles with a delay (allows mouse to reach handles)
   */
  scheduleHideHandles(): void {
    // Clear any existing timeout
    if (this.hideTimeout) {
      clearTimeout(this.hideTimeout);
    }
    
    // Schedule hide with a small delay
    this.hideTimeout = setTimeout(() => {
      // Only hide if we're not hovering over handles
      if (!this.isHoveringHandle) {
        this.hideResizeHandles();
        this.hideDragHandle();
      }
      this.hideTimeout = null;
    }, 100);
  }
  
  /**
   * Cancel scheduled hide (call when mouse enters table or handles)
   */
  cancelScheduledHide(): void {
    if (this.hideTimeout) {
      clearTimeout(this.hideTimeout);
      this.hideTimeout = null;
    }
  }
  
  /**
   * Get the current table element that has handles shown
   */
  getCurrentTableEl(): HTMLTableElement | null {
    return this.currentTableEl;
  }
  
  /**
   * Destroy the manager and clean up
   */
  destroy(): void {
    if (this.container) {
      this.container.removeEventListener('contextmenu', this.boundContextMenu);
    }
    document.removeEventListener('click', this.boundClick);
    document.removeEventListener('mousemove', this.boundMouseMove);
    document.removeEventListener('mouseup', this.boundMouseUp);
    
    this.hideResizeHandles();
    this.hideContextMenu();
    this.dragHandle?.remove();
    
    this.container = null;
    this.editorView = null;
  }
  
  // ============================================================================
  // Resize Handles
  // ============================================================================
  
  private createResizeHandles(): void {
    // Handles are created dynamically on table hover
  }
  
  private clearResizeHandles(): void {
    this.resizeHandles.forEach(h => h.remove());
    this.resizeHandles = [];
  }
  
  private createColumnResizeHandles(tableEl: HTMLTableElement): void {
    if (!this.container) return;
    
    const tableRect = tableEl.getBoundingClientRect();
    const containerRect = this.container.getBoundingClientRect();
    
    // Get column positions from first row
    // `rows` and `cells` hold only this table's own, not those of tables nested in its cells
    const firstRow = tableEl.rows[0];
    if (!firstRow) return;
    
    const cells = Array.from(firstRow.cells);
    let accumulatedWidth = 0;
    
    cells.forEach((cell, index) => {
      if (index === cells.length - 1) return; // Don't add handle after last column
      
      const cellRect = (cell as HTMLElement).getBoundingClientRect();
      accumulatedWidth += cellRect.width;
      
      const handle = document.createElement('div');
      handle.className = 'table-col-resize-handle';
      handle.style.cssText = `
        position: absolute;
        top: ${tableRect.top - containerRect.top}px;
        left: ${tableRect.left - containerRect.left + accumulatedWidth - 3}px;
        width: 6px;
        height: ${tableRect.height}px;
        cursor: col-resize;
        background: transparent;
        z-index: 100;
      `;
      handle.dataset.colIndex = String(index);
      
      handle.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.startColumnResize(tableEl, index, e.clientX, cellRect.width);
      });
      
      handle.addEventListener('mouseenter', () => {
        this.isHoveringHandle = true;
        handle.style.background = 'rgba(26, 115, 232, 0.3)';
      });
      
      handle.addEventListener('mouseleave', () => {
        this.isHoveringHandle = false;
        if (!this.resizeState) {
          handle.style.background = 'transparent';
        }
      });
      
      this.container!.appendChild(handle);
      this.resizeHandles.push(handle);
    });
  }
  
  private createRowResizeHandles(tableEl: HTMLTableElement): void {
    if (!this.container) return;
    
    const tableRect = tableEl.getBoundingClientRect();
    const containerRect = this.container.getBoundingClientRect();
    
    const rows = Array.from(tableEl.rows);
    let accumulatedHeight = 0;
    
    rows.forEach((row, index) => {
      if (index === rows.length - 1) return; // Don't add handle after last row
      
      const rowRect = (row as HTMLElement).getBoundingClientRect();
      accumulatedHeight += rowRect.height;
      
      const handle = document.createElement('div');
      handle.className = 'table-row-resize-handle';
      handle.style.cssText = `
        position: absolute;
        top: ${tableRect.top - containerRect.top + accumulatedHeight - 3}px;
        left: ${tableRect.left - containerRect.left}px;
        width: ${tableRect.width}px;
        height: 6px;
        cursor: row-resize;
        background: transparent;
        z-index: 100;
      `;
      handle.dataset.rowIndex = String(index);
      
      handle.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.startRowResize(tableEl, index, e.clientY, rowRect.height);
      });
      
      handle.addEventListener('mouseenter', () => {
        this.isHoveringHandle = true;
        handle.style.background = 'rgba(26, 115, 232, 0.3)';
      });
      
      handle.addEventListener('mouseleave', () => {
        this.isHoveringHandle = false;
        if (!this.resizeState) {
          handle.style.background = 'transparent';
        }
      });
      
      this.container!.appendChild(handle);
      this.resizeHandles.push(handle);
    });
  }
  
  private startColumnResize(tableEl: HTMLTableElement, colIndex: number, startX: number, startWidth: number): void {
    const pmStart = parseInt(tableEl.dataset.pmStart || '0', 10);
    
    this.resizeState = {
      type: 'column',
      tableEl,
      index: colIndex,
      startPos: startX,
      startSize: startWidth,
      pmStart,
    };
    
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
  }
  
  private startRowResize(tableEl: HTMLTableElement, rowIndex: number, startY: number, startHeight: number): void {
    const pmStart = parseInt(tableEl.dataset.pmStart || '0', 10);
    
    this.resizeState = {
      type: 'row',
      tableEl,
      index: rowIndex,
      startPos: startY,
      startSize: startHeight,
      pmStart,
    };
    
    document.body.style.cursor = 'row-resize';
    document.body.style.userSelect = 'none';
  }
  
  // ============================================================================
  // Table Drag (Move Up/Down)
  // ============================================================================
  
  private startTableDrag(tableEl: HTMLTableElement, startY: number): void {
    this.isDragging = true;
    this.dragStartY = startY;
    this.dragTableEl = tableEl;
    this.dropTargetPos = null;
    
    // Change cursor
    document.body.style.cursor = 'grabbing';
    document.body.style.userSelect = 'none';
    if (this.dragHandle) {
      this.dragHandle.style.cursor = 'grabbing';
    }
    
    // Create a ghost element (visual feedback showing the table being dragged)
    const tableRect = tableEl.getBoundingClientRect();
    this.dragGhost = document.createElement('div');
    this.dragGhost.style.cssText = `
      position: fixed;
      top: ${tableRect.top}px;
      left: ${tableRect.left}px;
      width: ${tableRect.width}px;
      height: ${tableRect.height}px;
      background: rgba(26, 115, 232, 0.08);
      border: 2px dashed #1a73e8;
      border-radius: 4px;
      pointer-events: none;
      z-index: 10000;
    `;
    document.body.appendChild(this.dragGhost);
    
    // Create drop indicator line
    this.dropIndicator = document.createElement('div');
    this.dropIndicator.style.cssText = `
      position: fixed;
      left: 0;
      right: 0;
      height: 3px;
      background: #1a73e8;
      pointer-events: none;
      z-index: 10001;
      display: none;
      box-shadow: 0 0 4px rgba(26, 115, 232, 0.5);
    `;
    document.body.appendChild(this.dropIndicator);
    
    // Dim the original table
    tableEl.style.opacity = '0.4';
  }
  
  private handleTableDrag(clientY: number): void {
    if (!this.isDragging || !this.dragTableEl || !this.container) return;
    
    // Move the ghost to follow the mouse
    if (this.dragGhost) {
      const deltaY = clientY - this.dragStartY;
      const tableRect = this.dragTableEl.getBoundingClientRect();
      this.dragGhost.style.top = `${tableRect.top + deltaY}px`;
    }
    
    const tablePmStart = parseInt(this.dragTableEl.dataset.pmStart || '0', 10);
    const tablePmEnd = parseInt(this.dragTableEl.dataset.pmEnd || '0', 10);
    
    // Find all TOP-LEVEL block elements (fragments represent top-level blocks)
    const fragments = this.container.querySelectorAll('.fragment');
    
    interface FragmentInfo {
      element: HTMLElement;
      pmStart: number;
      pmEnd: number;
      top: number;
      bottom: number;
      isListItem: boolean;
    }
    
    const allFragments: FragmentInfo[] = [];
    
    for (const fragment of fragments) {
      const fragmentEl = fragment as HTMLElement;
      
      // Get the pm positions from the fragment or its first child with data-pm-start
      const blockWithPos = fragmentEl.querySelector('[data-pm-start]') as HTMLElement | null;
      if (!blockWithPos) continue;
      
      const pmStart = parseInt(blockWithPos.dataset.pmStart || '0', 10);
      const pmEnd = parseInt(blockWithPos.dataset.pmEnd || '0', 10);
      
      // Skip the table itself
      if (pmStart >= tablePmStart && pmStart <= tablePmEnd) continue;
      
      // Skip if this is inside a table cell
      if (blockWithPos.closest('.table-cell')) continue;
      
      const rect = fragmentEl.getBoundingClientRect();
      
      // Detect if this is a list item (ul or ol inside the fragment)
      const isListItem = fragmentEl.querySelector('ul, ol') !== null;
      
      allFragments.push({
        element: fragmentEl,
        pmStart,
        pmEnd,
        top: rect.top,
        bottom: rect.bottom,
        isListItem,
      });
    }
    
    // Sort by vertical position
    allFragments.sort((a, b) => a.top - b.top);
    
    // Group consecutive list items together
    interface DropTarget {
      element: HTMLElement;
      pmStart: number;
      pmEnd: number;
      top: number;
      bottom: number;
    }
    
    const validTargets: DropTarget[] = [];
    let i = 0;
    
    while (i < allFragments.length) {
      const frag = allFragments[i];
      
      if (frag.isListItem) {
        // Start of a list group - find all consecutive list items
        const groupStart = frag.pmStart;
        let groupEnd = frag.pmEnd;
        const groupTop = frag.top;
        let groupBottom = frag.bottom;
        const firstElement = frag.element;
        
        // Look ahead for consecutive list items
        let j = i + 1;
        while (j < allFragments.length && allFragments[j].isListItem) {
          groupEnd = allFragments[j].pmEnd;
          groupBottom = allFragments[j].bottom;
          j++;
        }
        
        // Add the entire list group as one target
        validTargets.push({
          element: firstElement,
          pmStart: groupStart,
          pmEnd: groupEnd,
          top: groupTop,
          bottom: groupBottom,
        });
        
        i = j; // Skip past all list items in this group
      } else {
        // Regular block
        validTargets.push({
          element: frag.element,
          pmStart: frag.pmStart,
          pmEnd: frag.pmEnd,
          top: frag.top,
          bottom: frag.bottom,
        });
        i++;
      }
    }
    
    // Find the best drop position
    let bestTarget: DropTarget | null = null;
    let insertBefore = true;
    
    if (validTargets.length > 0) {
      // Check if we're above the first target
      if (clientY < validTargets[0].top) {
        bestTarget = validTargets[0];
        insertBefore = true;
      } else {
        // Find which block/group we're in or closest to
        for (let idx = 0; idx < validTargets.length; idx++) {
          const target = validTargets[idx];
          
          if (clientY >= target.top && clientY <= target.bottom) {
            // We're inside this block/group - snap to top or bottom based on position
            const midPoint = (target.top + target.bottom) / 2;
            if (clientY < midPoint) {
              bestTarget = target;
              insertBefore = true;
            } else {
              bestTarget = target;
              insertBefore = false;
            }
            break;
          } else if (clientY > target.bottom) {
            // We're below this block
            bestTarget = target;
            insertBefore = false;
            // Continue to check if there's a block below
          } else if (clientY < target.top && bestTarget) {
            // We're between bestTarget and this target
            break;
          }
        }
      }
    }
    
    // Position the drop indicator
    if (bestTarget && this.dropIndicator) {
      const indicatorY = insertBefore ? bestTarget.top : bestTarget.bottom;
      
      // Get the page boundaries for proper left/right positioning
      const pageEl = bestTarget.element.closest('.page-content') || bestTarget.element.closest('.editor-page');
      let left = bestTarget.element.getBoundingClientRect().left;
      let width = bestTarget.element.getBoundingClientRect().width;
      
      if (pageEl) {
        const pageRect = pageEl.getBoundingClientRect();
        left = pageRect.left + 40;
        width = pageRect.width - 80;
      }
      
      this.dropIndicator.style.display = 'block';
      this.dropIndicator.style.top = `${indicatorY - 1}px`;
      this.dropIndicator.style.left = `${left}px`;
      this.dropIndicator.style.width = `${width}px`;
      
      // Store the target position
      this.dropTargetPos = insertBefore ? bestTarget.pmStart : bestTarget.pmEnd;
    } else if (this.dropIndicator) {
      this.dropIndicator.style.display = 'none';
      this.dropTargetPos = null;
    }
  }
  
  private endTableDrag(): void {
    if (!this.isDragging || !this.dragTableEl || !this.editorView) {
      this.cleanupDrag();
      return;
    }
    
    const { state, dispatch } = this.editorView;
    const tableEl = this.dragTableEl;
    const pmStart = parseInt(tableEl.dataset.pmStart || '0', 10);
    const pmEnd = parseInt(tableEl.dataset.pmEnd || '0', 10);
    
    // Use the stored drop target position
    const targetPos = this.dropTargetPos;
    
    // Only move if we have a valid target and it's different from current position
    if (targetPos !== null && targetPos !== pmStart && targetPos !== pmEnd) {
      try {
        // Get the table node
        const $start = state.doc.resolve(pmStart);
        let tableNode = null;
        let tablePos = pmStart;
        
        // Find the table node
        for (let d = $start.depth; d >= 0; d--) {
          const node = $start.node(d);
          if (node.type.name === 'table') {
            tableNode = node;
            tablePos = $start.before(d);
            break;
          }
        }
        
        if (!tableNode) {
          // pmStart might be at the table itself
          const node = state.doc.nodeAt(pmStart);
          if (node?.type.name === 'table') {
            tableNode = node;
            tablePos = pmStart;
          }
        }
        
        if (tableNode) {
          let tr = state.tr;
          
          // Delete the table from its current position
          tr = tr.delete(tablePos, tablePos + tableNode.nodeSize);
          
          // Adjust target position if it's after the deleted content
          let adjustedTargetPos = targetPos;
          if (targetPos > tablePos) {
            adjustedTargetPos = targetPos - tableNode.nodeSize;
          }
          
          // Insert at the new position
          tr = tr.insert(adjustedTargetPos, tableNode);
          
          dispatch(tr);
          
          // Trigger layout update
          if (this.onTableUpdate) {
            this.onTableUpdate();
          }
        }
      } catch (err) {
        console.error('Error moving table:', err);
      }
    }
    
    this.cleanupDrag();
  }
  
  private cleanupDrag(): void {
    // Remove drop indicator
    if (this.dropIndicator) {
      this.dropIndicator.remove();
      this.dropIndicator = null;
    }
    
    // Remove ghost (if any)
    if (this.dragGhost) {
      this.dragGhost.remove();
      this.dragGhost = null;
    }
    
    // Restore table opacity
    if (this.dragTableEl) {
      this.dragTableEl.style.opacity = '1';
      this.dragTableEl = null;
    }
    
    // Reset state
    this.isDragging = false;
    this.dragStartY = 0;
    this.dropTargetPos = null;
    
    // Reset cursors
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
    if (this.dragHandle) {
      this.dragHandle.style.cursor = 'grab';
    }
  }
  
  // ============================================================================
  // Drag Handle
  // ============================================================================
  
  private showDragHandle(tableEl: HTMLTableElement): void {
    if (!this.container) return;
    
    if (!this.dragHandle) {
      this.dragHandle = document.createElement('div');
      this.dragHandle.className = 'table-drag-handle';
      this.dragHandle.innerHTML = `
        <svg width="12" height="20" viewBox="0 0 12 20" fill="currentColor">
          <circle cx="3" cy="4" r="1.5"/>
          <circle cx="3" cy="10" r="1.5"/>
          <circle cx="3" cy="16" r="1.5"/>
          <circle cx="9" cy="4" r="1.5"/>
          <circle cx="9" cy="10" r="1.5"/>
          <circle cx="9" cy="16" r="1.5"/>
        </svg>
      `;
      this.dragHandle.style.cssText = `
        position: absolute;
        width: 20px;
        height: 32px;
        background: #fff;
        border: 1px solid #dadce0;
        border-radius: 4px;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: grab;
        z-index: 101;
        color: #5f6368;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
      `;
      
      this.dragHandle.addEventListener('mouseenter', () => {
        this.isHoveringHandle = true;
        this.cancelScheduledHide();
        this.dragHandle!.style.background = '#f1f3f4';
      });
      
      this.dragHandle.addEventListener('mouseleave', () => {
        if (!this.isDragging) {
          this.isHoveringHandle = false;
          this.dragHandle!.style.background = '#fff';
          // Hide handles when leaving the drag handle (if not over table)
          if (!this.currentTableEl?.matches(':hover')) {
            this.hideResizeHandles();
            this.hideDragHandle();
          }
        }
      });
      
      this.dragHandle.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (this.currentTableEl) {
          this.startTableDrag(this.currentTableEl, e.clientY);
        }
      });
      
      this.container.appendChild(this.dragHandle);
    }
    
    const tableRect = tableEl.getBoundingClientRect();
    const containerRect = this.container.getBoundingClientRect();
    
    // Position on the left side, vertically centered
    this.dragHandle.style.display = 'flex';
    this.dragHandle.style.top = `${tableRect.top - containerRect.top + (tableRect.height / 2) - 16}px`;
    this.dragHandle.style.left = `${tableRect.left - containerRect.left - 28}px`;
    
    // Store reference to current table for hover detection
    this.currentTableEl = tableEl;
  }
  
  // ============================================================================
  // Context Menu
  // ============================================================================
  
  private createContextMenu(): void {
    if (this.contextMenu) return;
    
    this.contextMenu = document.createElement('div');
    this.contextMenu.className = 'table-context-menu';
    this.contextMenu.style.cssText = `
      position: fixed;
      background: #fff;
      border: 1px solid #dadce0;
      border-radius: 8px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.15);
      padding: 4px 0;
      z-index: 10000;
      display: none;
      min-width: 180px;
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      font-size: 13px;
    `;
    
    document.body.appendChild(this.contextMenu);
  }
  
  private handleContextMenu(e: MouseEvent): void {
    const target = e.target as HTMLElement;
    const tableCell = target.closest('.table-cell') as HTMLElement | null;
    const tableEl = target.closest('.table-block') as HTMLTableElement | null;
    
    if (!tableCell || !tableEl) return;
    
    e.preventDefault();
    
    const pmStart = parseInt(tableEl.dataset.pmStart || '0', 10);
    const pmEnd = parseInt(tableEl.dataset.pmEnd || '0', 10);
    
    // Find row and physical column indices from the DOM.
    const row = tableCell.closest('tr');
    const rowIndex = row ? Array.from(row.parentElement?.children || []).indexOf(row) : 0;
    const physicalCol = row ? Array.from(row.children).indexOf(tableCell) : 0;

    this.contextMenuTarget = { tableEl, pmStart, pmEnd };

    // Convert the physical cell index to a logical column so column insert/delete
    // stay correct on merged tables (a spanned cell's physical index != column).
    const found = this.findTable(pmStart);
    const colIndex = found
      ? logicalColumnForCell(found.tableNode, found.tablePos, rowIndex, physicalCol)
      : physicalCol;

    this.showContextMenu(e.clientX, e.clientY, rowIndex, colIndex);
  }
  
  private showContextMenu(x: number, y: number, rowIndex: number, colIndex: number): void {
    if (!this.contextMenu) return;
    
    const actions: ContextMenuAction[] = [
      {
        label: 'Insert row above',
        icon: '↑',
        action: () => this.insertRow(rowIndex, 'before'),
      },
      {
        label: 'Insert row below',
        icon: '↓',
        action: () => this.insertRow(rowIndex, 'after'),
      },
      {
        label: 'Insert column left',
        icon: '←',
        action: () => this.insertColumn(colIndex, 'before'),
      },
      {
        label: 'Insert column right',
        icon: '→',
        action: () => this.insertColumn(colIndex, 'after'),
      },
      {
        label: '—',
        action: () => {}, // Separator
      },
      {
        label: 'Delete row',
        icon: '✕',
        action: () => this.deleteRow(rowIndex),
      },
      {
        label: 'Delete column',
        icon: '✕',
        action: () => this.deleteColumn(colIndex),
      },
    ];
    
    this.contextMenu.innerHTML = '';
    
    for (const action of actions) {
      if (action.label === '—') {
        const separator = document.createElement('div');
        separator.style.cssText = `
          height: 1px;
          background: #e8eaed;
          margin: 4px 0;
        `;
        this.contextMenu.appendChild(separator);
        continue;
      }
      
      const item = document.createElement('div');
      item.className = 'context-menu-item';
      item.style.cssText = `
        padding: 8px 12px;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 8px;
      `;
      
      if (action.icon) {
        const icon = document.createElement('span');
        icon.textContent = action.icon;
        icon.style.cssText = `
          width: 16px;
          text-align: center;
          opacity: 0.7;
        `;
        item.appendChild(icon);
      }
      
      const label = document.createElement('span');
      label.textContent = action.label;
      item.appendChild(label);
      
      item.addEventListener('mouseenter', () => {
        item.style.background = '#f1f3f4';
      });
      
      item.addEventListener('mouseleave', () => {
        item.style.background = 'transparent';
      });
      
      item.addEventListener('click', () => {
        action.action();
        this.hideContextMenu();
      });
      
      this.contextMenu.appendChild(item);
    }
    
    this.contextMenu.style.display = 'block';
    this.contextMenu.style.left = `${x}px`;
    this.contextMenu.style.top = `${y}px`;
    
    // Adjust position if menu goes off-screen
    const rect = this.contextMenu.getBoundingClientRect();
    if (rect.right > window.innerWidth) {
      this.contextMenu.style.left = `${x - rect.width}px`;
    }
    if (rect.bottom > window.innerHeight) {
      this.contextMenu.style.top = `${y - rect.height}px`;
    }
  }
  
  private hideContextMenu(): void {
    if (this.contextMenu) {
      this.contextMenu.style.display = 'none';
    }
    this.contextMenuTarget = null;
  }
  
  // ============================================================================
  // Table Modification Commands
  // ============================================================================
  
  // Span-aware row insert: widens a rowspan cell that straddles the boundary and
  // fills only the uncovered logical columns of the new row (rows map 1:1 to
  // table_row nodes, so rowIndex is already logical).
  private insertRow(rowIndex: number, position: 'before' | 'after'): void {
    if (!this.editorView || !this.contextMenuTarget) return;
    const { pmStart } = this.contextMenuTarget;
    const { state, dispatch } = this.editorView;
    const found = this.findTable(pmStart);
    if (!found) return;

    const tr = state.tr;
    if (applyInsertRow(tr, found.tableNode, found.tablePos, rowIndex, position, state.schema)) {
      dispatch(tr);
      this.onTableUpdate?.();
    }
  }
  
  /** Resolve the enclosing table node and its start position from a pmStart. */
  private findTable(pmStart: number): { tableNode: PmNode; tablePos: number } | null {
    if (!this.editorView) return null;
    const { state } = this.editorView;
    const $pos = state.doc.resolve(pmStart + 1);
    for (let d = $pos.depth; d >= 0; d--) {
      const node = $pos.node(d);
      if (node.type.name === 'table') {
        return { tableNode: node, tablePos: $pos.before(d) };
      }
    }
    return null;
  }

  // colIndex is a LOGICAL column (handleContextMenu converts the clicked cell's
  // physical index). The span-aware transform widens a straddling cell or splices
  // a real cell per row, staying correct on merged tables.
  private insertColumn(colIndex: number, position: 'before' | 'after'): void {
    if (!this.editorView || !this.contextMenuTarget) return;
    const { pmStart } = this.contextMenuTarget;
    const { state, dispatch } = this.editorView;
    const found = this.findTable(pmStart);
    if (!found) return;

    const targetColIndex = position === 'after' ? colIndex + 1 : colIndex;
    const tr = state.tr;
    if (applyInsertColumn(tr, found.tableNode, found.tablePos, targetColIndex, state.schema)) {
      dispatch(tr);
      this.onTableUpdate?.();
    }
  }
  
  // Span-aware row delete: shrinks a rowspan cell that covers the deleted row and
  // relocates a rowspan cell's content down instead of losing it.
  private deleteRow(rowIndex: number): void {
    if (!this.editorView || !this.contextMenuTarget) return;
    const { pmStart } = this.contextMenuTarget;
    const { state, dispatch } = this.editorView;
    const found = this.findTable(pmStart);
    if (!found) return;

    const tr = state.tr;
    if (applyDeleteRow(tr, found.tableNode, found.tablePos, rowIndex, state.schema)) {
      dispatch(tr);
      this.onTableUpdate?.();
    }
  }
  
  // colIndex is a LOGICAL column. The span-aware transform shrinks a spanning
  // cell or removes the origin cell, once per cell, so merged tables stay a
  // rectangular grid instead of going ragged.
  private deleteColumn(colIndex: number): void {
    if (!this.editorView || !this.contextMenuTarget) return;
    const { pmStart } = this.contextMenuTarget;
    const { state, dispatch } = this.editorView;
    const found = this.findTable(pmStart);
    if (!found) return;

    const tr = state.tr;
    if (applyDeleteColumn(tr, found.tableNode, found.tablePos, colIndex)) {
      dispatch(tr);
      this.onTableUpdate?.();
    }
  }
  
  // ============================================================================
  // Event Handlers
  // ============================================================================
  
  private handleMouseMove(e: MouseEvent): void {
    // Handle table drag
    if (this.isDragging) {
      this.handleTableDrag(e.clientY);
      return;
    }
    
    if (!this.resizeState) return;
    
    const { type, tableEl, index, startPos, startSize } = this.resizeState;
    
    if (type === 'column') {
      const delta = e.clientX - startPos;
      const newWidth = Math.max(40, startSize + delta); // Minimum 40px width
      
      // Apply to all cells in this column
      const rows = Array.from(tableEl.rows);
      rows.forEach((row) => {
        const cells = Array.from(row.cells);
        if (cells[index]) {
          (cells[index] as HTMLElement).style.width = `${newWidth}px`;
        }
      });
    } else if (type === 'row') {
      const delta = e.clientY - startPos;
      const newHeight = Math.max(24, startSize + delta); // Minimum 24px height
      
      const rows = Array.from(tableEl.rows);
      if (rows[index]) {
        (rows[index] as HTMLElement).style.height = `${newHeight}px`;
      }
    }
  }
  
  private handleMouseUp(_e: MouseEvent): void {
    // Handle table drag end
    if (this.isDragging) {
      this.endTableDrag();
      return;
    }
    
    if (this.resizeState) {
      // Update ProseMirror with new dimensions
      this.applyResizeToProseMirror();
      
      this.resizeState = null;
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      
      // Trigger layout update
      this.onTableUpdate?.();
    }
  }
  
  private applyResizeToProseMirror(): void {
    if (!this.resizeState || !this.editorView) return;
    
    const { type, tableEl, index, pmStart } = this.resizeState;
    const { state, dispatch } = this.editorView;
    
    if (type === 'column') {
      // Get the new width from the DOM (this is the rendered/scaled width)
      const rows = Array.from(tableEl.rows);
      if (rows.length === 0) return;
      
      const cells = Array.from(rows[0].cells);
      if (!cells[index]) return;
      
      const renderedWidth = (cells[index] as HTMLElement).offsetWidth;
      // Convert to unscaled width for storage in ProseMirror
      const newWidth = Math.round(renderedWidth / this.scale);
      
      // pmStart is the position of the table node itself
      // We need to find it in the document
      let tableNode = null;
      let tablePos = pmStart;
      
      // Verify this position contains a table
      if (pmStart < state.doc.content.size) {
        const nodeAtPos = state.doc.nodeAt(pmStart);
        if (nodeAtPos && nodeAtPos.type.name === 'table') {
          tableNode = nodeAtPos;
        }
      }
      
      // If not found directly, try resolving
      if (!tableNode) {
        const $pos = state.doc.resolve(Math.min(pmStart + 1, state.doc.content.size));
        for (let d = $pos.depth; d >= 0; d--) {
          const node = $pos.node(d);
          if (node.type.name === 'table') {
            tableNode = node;
            tablePos = $pos.before(d);
            break;
          }
        }
      }
      
      if (!tableNode) return;

      // The handle's `index` is a physical cell index in row 0; resize the
      // logical column it starts at, span-aware, so merged tables don't get
      // colwidth stamped on the wrong column (or a length-1 array on a span).
      const logicalCol = logicalColumnForCell(tableNode, tablePos, 0, index);
      const tr = state.tr;
      if (applyColumnResize(tr, tableNode, tablePos, logicalCol, newWidth)) {
        dispatch(tr);
      }
    }
    // Row height is typically handled by cell content, but we could add similar logic if needed
  }
  
  private handleClick(e: MouseEvent): void {
    // Hide context menu on click outside
    if (this.contextMenu && !this.contextMenu.contains(e.target as Node)) {
      this.hideContextMenu();
    }
  }
}

// ============================================================================
// Factory Function
// ============================================================================

let tableInteractionInstance: TableInteractionManager | null = null;

/**
 * Get the singleton table interaction manager
 */
export function getTableInteractionManager(): TableInteractionManager {
  if (!tableInteractionInstance) {
    tableInteractionInstance = new TableInteractionManager();
  }
  return tableInteractionInstance;
}

/**
 * Create a new table interaction manager instance
 */
export function createTableInteractionManager(): TableInteractionManager {
  return new TableInteractionManager();
}

