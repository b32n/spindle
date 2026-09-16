import React, { memo, useState, useCallback, useRef, useEffect, useMemo } from 'react';
import { useWorkbook } from '../context/WorkbookContext';
import { CanvasGrid, type ContextMenuType } from './CanvasGrid';
import { EditOverlay, type EditOverlayRef } from './EditOverlay';
import { FormulaAutocomplete } from './FormulaAutocomplete';
import { FormulaSignatureHint } from './FormulaSignatureHint';
import { useFormulaAssist } from '../hooks/useFormulaAssist';
import { FormulaBar } from './FormulaBar';
import { Toolbar } from './Toolbar';
import { SheetTabs } from './SheetTabs';
import { CommentsPanel } from './CommentsPanel';
import { ContextMenu } from './ContextMenu';
import { HeaderContextMenu } from './HeaderContextMenu';
import { FilterModal } from './FilterModal';
import { FormatCellsModal } from './FormatCellsModal';
import { FindReplaceModal, type FindReplaceState } from './FindReplaceModal';
import type { CellPosition, Selection, CellFormat, ColumnFilter, CellStyle, SortOrder, FormatType } from '@b32nio/spindle-sheets-core';
import { columnIndexToLabel } from '@b32nio/spindle-sheets-core';
import { parseDateString } from '@b32nio/spindle-sheets-core';
import { findMatches, computeReplacement } from '@b32nio/spindle-sheets-core';

export interface WorkbookCanvasProps {
  className?: string;
  style?: React.CSSProperties;
  width?: number;
  height?: number;
  rowHeight?: number;
  colWidth?: number;
}

export const WorkbookCanvas = memo(function WorkbookCanvas({
  className,
  style,
  width = 800,
  height = 600,
  rowHeight = 20,
  colWidth = 100,
}: WorkbookCanvasProps) {
  const { workbook } = useWorkbook();
  const [activeCell, setActiveCell] = useState<CellPosition | null>(null);
  const [editingCell, setEditingCell] = useState<CellPosition | null>(null);
  const [editValue, setEditValue] = useState('');
  const [caret, setCaret] = useState(0);
  const [editingCellFormat, setEditingCellFormat] = useState<CellFormat | undefined>(undefined);
  const [dimensionVersion, setDimensionVersion] = useState(0);
  const [showComments, setShowComments] = useState(false);
  const [originalEditingSheetId, setOriginalEditingSheetId] = useState<string | null>(null);
  const canvasGridRef = useRef<HTMLDivElement>(null);
  const editOverlayRef = useRef<EditOverlayRef>(null);
  const floatingInputRef = useRef<HTMLInputElement>(null);
  const lastInsertRef = useRef<{ start: number; end: number } | null>(null);
  
  // Floating input position (for dragging)
  const [floatingInputPos, setFloatingInputPos] = useState<{ x: number; y: number } | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const dragStartRef = useRef<{ mouseX: number; mouseY: number; posX: number; posY: number } | null>(null);
  // Store the editing cell's visual position when editing starts (for floating input placement)
  const editingCellPosRef = useRef<{ x: number; y: number; width: number; height: number } | null>(null);
  
  // Context menu state
  const [contextMenu, setContextMenu] = useState<ContextMenuType | null>(null);

  // Filter modal state
  const [filterModal, setFilterModal] = useState<{
    isOpen: boolean;
    column: number;
    existingFilter?: ColumnFilter;
  } | null>(null);

  // Format cells modal state
  const [formatModal, setFormatModal] = useState<{
    isOpen: boolean;
    currentFormat?: CellFormat;
    sampleValue?: number;
  } | null>(null);

  // Find & replace state
  const [findReplace, setFindReplace] = useState<FindReplaceState & { isOpen: boolean }>({
    isOpen: false,
    query: '',
    replacement: '',
    matchCase: false,
    wholeCell: false,
    searchFormulas: false,
  });
  const [activeMatchIndex, setActiveMatchIndex] = useState(-1);
  // "Select + reveal a cell" function exposed by CanvasGrid.
  const navigateRef = useRef<((cell: CellPosition) => void) | null>(null);
  
  // Clipboard handlers from CanvasGrid
  const clipboardHandlersRef = useRef<{
    copy: () => void;
    cut: () => void;
    paste: (targetCell?: CellPosition) => Promise<void>;
  } | null>(null);
  
  // Check if we're editing a formula
  const isEditingFormula = editingCell !== null && editValue.startsWith('=');

  // Formula autocomplete + parameter help for the in-cell editor.
  const onFormulaAccept = useCallback((nextValue: string, nextCaret: number) => {
    setEditValue(nextValue);
    setCaret(nextCaret);
    editOverlayRef.current?.setSelection(nextCaret);
  }, []);
  const formulaAssist = useFormulaAssist({
    value: editValue,
    caret,
    enabled: isEditingFormula,
    onAccept: onFormulaAccept,
  });
  
  // Check if we're on a different sheet than where editing started
  // Use workbook.activeSheetId directly to get the current value immediately (not via state)
  const currentSheetId = workbook.activeSheetId;
  const isOnDifferentSheet = originalEditingSheetId !== null && currentSheetId !== originalEditingSheetId;
  
  // Calculate layout dimensions (needed for floating input positioning)
  const headerHeight = rowHeight;
  const toolbarHeight = 60;
  const formulaBarHeight = 32;
  const canvasAreaHeight = height - toolbarHeight - formulaBarHeight;
  const canvasAreaWidth = width;
  
  // Focus floating input when it appears, or EditOverlay when returning to original sheet
  // This effect runs when dimensionVersion changes (which happens after sheet switch)
  useEffect(() => {
    if (isEditingFormula) {
      if (isOnDifferentSheet) {
        // Focus floating input when on different sheet
        setTimeout(() => {
          if (floatingInputRef.current) {
            floatingInputRef.current.focus();
            const len = editValue.length;
            floatingInputRef.current.setSelectionRange(len, len);
          }
        }, 50);
        
        // Calculate initial position below the editing cell (if not already set)
        if (!floatingInputPos && editingCellPosRef.current) {
          const cellPos = editingCellPosRef.current;
          // Position below the cell, with some padding
          const floatingWidth = 280;
          let x = cellPos.x;
          let y = cellPos.y + cellPos.height + 4; // 4px below the cell
          
          // Make sure it doesn't go off the right edge
          if (x + floatingWidth > canvasAreaWidth - 8) {
            x = Math.max(8, canvasAreaWidth - floatingWidth - 8);
          }
          
          // Make sure it doesn't go off the bottom edge (leave room for ~80px height)
          if (y + 80 > canvasAreaHeight - 8) {
            // Position above the cell instead
            y = Math.max(8, cellPos.y - 80 - 4);
          }
          
          setFloatingInputPos({ x, y });
        } else if (!floatingInputPos) {
          // Fallback to top-left if no cell position stored
          setFloatingInputPos({ x: 8, y: 8 });
        }
      } else {
        // Focus EditOverlay when back on original sheet
        setTimeout(() => {
          if (editOverlayRef.current) {
            editOverlayRef.current.focus();
          }
        }, 50);
      }
    }
  }, [dimensionVersion, isOnDifferentSheet, isEditingFormula, editValue.length, floatingInputPos, canvasAreaWidth, canvasAreaHeight]);
  
  // Reset floating input position when editing ends
  useEffect(() => {
    if (!editingCell) {
      setFloatingInputPos(null);
      editingCellPosRef.current = null;
    }
  }, [editingCell]);
  
  // Handle floating input drag
  const handleFloatingDragStart = useCallback((e: React.MouseEvent) => {
    if (floatingInputPos) {
      e.preventDefault();
      setIsDragging(true);
      dragStartRef.current = {
        mouseX: e.clientX,
        mouseY: e.clientY,
        posX: floatingInputPos.x,
        posY: floatingInputPos.y,
      };
    }
  }, [floatingInputPos]);
  
  const handleFloatingDragMove = useCallback((e: MouseEvent) => {
    if (isDragging && dragStartRef.current) {
      const deltaX = e.clientX - dragStartRef.current.mouseX;
      const deltaY = e.clientY - dragStartRef.current.mouseY;
      setFloatingInputPos({
        x: dragStartRef.current.posX + deltaX,
        y: dragStartRef.current.posY + deltaY,
      });
    }
  }, [isDragging]);
  
  const handleFloatingDragEnd = useCallback(() => {
    setIsDragging(false);
    dragStartRef.current = null;
  }, []);
  
  // Add global mouse listeners for dragging
  useEffect(() => {
    if (isDragging) {
      window.addEventListener('mousemove', handleFloatingDragMove);
      window.addEventListener('mouseup', handleFloatingDragEnd);
      return () => {
        window.removeEventListener('mousemove', handleFloatingDragMove);
        window.removeEventListener('mouseup', handleFloatingDragEnd);
      };
    }
    return undefined;
  }, [isDragging, handleFloatingDragMove, handleFloatingDragEnd]);

  // Store a raw edited string into a cell: formula, blank, number, an
  // auto-detected date, or text. Shared by the in-cell editor and formula bar.
  const applyCellInput = useCallback(
    (sheetId: string | undefined, row: number, col: number, raw: string) => {
      if (raw.startsWith('=')) {
        workbook.setFormula(sheetId, row, col, raw);
        return;
      }
      if (raw === '') {
        workbook.setCellValue(sheetId, row, col, null);
        return;
      }
      const numValue = Number(raw);
      if (!isNaN(numValue) && raw.trim() !== '' && isFinite(numValue)) {
        workbook.setCellValue(sheetId, row, col, numValue);
        return;
      }
      // Recognise a typed date and store it as an Excel serial with a date
      // format — matching how Excel / Google Sheets auto-detect on entry.
      const dateSerial = parseDateString(raw);
      if (dateSerial !== null) {
        // Pick a display pattern matching the user's input style: slashes
        // round-trip as MM/DD/YYYY, dashes/dots as DD-MM-YYYY, ISO as YYYY-MM-DD.
        const trimmed = raw.trim();
        let dateFormat = 'MM/DD/YYYY';
        if (/^\d{4}[/.-]/.test(trimmed)) dateFormat = 'YYYY-MM-DD';
        else if (/^\d{1,2}[.-]/.test(trimmed)) dateFormat = 'DD-MM-YYYY';
        const existing = workbook.getCell(sheetId, row, col);
        // `format` is resolved into the pool by setCell (see applyFormatToSelection).
        const dateCell = {
          ...existing,
          value: dateSerial,
          formula: undefined,
          formulaAst: undefined,
          format: { type: 'date' as const, dateFormat },
        };
        workbook.batch(() => {
          workbook.setCell(sheetId, row, col, dateCell);
        });
        return;
      }
      workbook.setCellValue(sheetId, row, col, raw);
    },
    [workbook]
  );

  // Handle formula bar changes
  const handleFormulaChange = useCallback(
    (formula: string) => {
      if (activeCell) {
        applyCellInput(undefined, activeCell.row, activeCell.col, formula);
        // Trigger re-render to show updated cell value
        setDimensionVersion(v => v + 1);
      }
    },
    [activeCell, applyCellInput]
  );

  // Commit edit and close editor
  // The value parameter comes from EditOverlay and is the most current value
  // moveToNextCell: if true, move to the cell below after committing (Enter key behavior)
  const commitEdit = useCallback((value?: string, moveToNextCell?: boolean) => {
    // Use passed value if provided, otherwise fall back to editValue state
    const valueToCommit = value !== undefined ? value : editValue;
    
    // Always save to the original editing sheet if we have one
    const targetSheetId = originalEditingSheetId || undefined;
    
    // Store the editing cell position before clearing state (needed for move to next cell)
    const currentEditingCell = editingCell;
    
    // A spilled cell is read-only — its content comes from the anchor formula.
    const isSpilled =
      currentEditingCell !== null &&
      workbook.isSpilledCell(targetSheetId, currentEditingCell.row, currentEditingCell.col);

    if (currentEditingCell && valueToCommit !== undefined && !isSpilled) {
      applyCellInput(targetSheetId, currentEditingCell.row, currentEditingCell.col, valueToCommit);
      // Trigger re-render to show updated cell value
      setDimensionVersion(v => v + 1);
    }
    setEditingCell(null);
    setEditValue('');
    setEditingCellFormat(undefined);
    setOriginalEditingSheetId(null);
    lastInsertRef.current = null;
    
    // Move to the next cell (row below) if requested (Enter key behavior)
    if (moveToNextCell && currentEditingCell) {
      const sheet = workbook.getSheet(targetSheetId);
      const newRow = Math.min(currentEditingCell.row + 1, sheet.rowCount - 1);
      const newCell = { row: newRow, col: currentEditingCell.col };
      
      setActiveCell(newCell);
      
      // Update selection to the new cell
      const newSelection: Selection = {
        ranges: [{
          startRow: newRow,
          endRow: newRow,
          startCol: currentEditingCell.col,
          endCol: currentEditingCell.col,
        }],
        activeCell: newCell,
      };
      workbook.setSelection(newSelection);
    }
    
    // Refocus the canvas so keyboard navigation continues to work
    // Use preventScroll to avoid the UI jumping when focus changes
    requestAnimationFrame(() => {
      const canvas = canvasGridRef.current?.querySelector('canvas');
      canvas?.focus({ preventScroll: true });
    });
  }, [editingCell, editValue, workbook, originalEditingSheetId, applyCellInput]);

  // Handle active cell change from canvas
  const handleActiveCellChange = useCallback((cell: CellPosition | null) => {
    setActiveCell(cell);
    // Stop editing if cell changes - BUT NOT if we're editing a formula
    // During formula editing, clicking cells should insert references, not cancel editing
    if (editingCell && cell && (cell.row !== editingCell.row || cell.col !== editingCell.col)) {
      // Don't cancel if we're editing a formula (on any sheet)
      if (!isEditingFormula) {
        commitEdit();
      }
    }
  }, [editingCell, commitEdit, isEditingFormula]);

  // Handle cell edit start from canvas
  const handleCellEdit = useCallback((cell: CellPosition, value: string) => {
    // A spilled cell is read-only — editing happens on the anchor formula.
    if (workbook.isSpilledCell(undefined, cell.row, cell.col)) return;

    // Get cell format for date/time picker detection
    const sheet = workbook.getSheet();
    const formatPool = workbook.getFormatPool();
    const cellData = sheet.getCell(cell.row, cell.col);
    const cellFormat = cellData?.formatId ? formatPool.get(cellData.formatId) : undefined;

    setEditingCell(cell);
    setEditValue(value);
    setCaret(value.length);
    setEditingCellFormat(cellFormat);
    // Track which sheet the edit started on (for cross-sheet formula references)
    setOriginalEditingSheetId(workbook.activeSheetId);
    lastInsertRef.current = null; // Clear any previous insert tracking
    
    // Store the cell's visual position for floating input placement
    let x = colWidth; // Start after row header
    for (let c = 0; c < cell.col; c++) {
      x += sheet.getColWidth(c);
    }
    let y = headerHeight; // Start after column header  
    for (let r = 0; r < cell.row; r++) {
      y += sheet.getRowHeight(r);
    }
    editingCellPosRef.current = {
      x,
      y,
      width: sheet.getColWidth(cell.col),
      height: sheet.getRowHeight(cell.row),
    };
  }, [workbook, colWidth, headerHeight]);
  
  // Handle inserting cell reference into formula
  const handleInsertCellReference = useCallback((reference: string, isNewSelection?: boolean) => {
    if (isEditingFormula) {
      // If on the original sheet, use the EditOverlay's insertAtCursor
      if (!isOnDifferentSheet && editOverlayRef.current) {
        // For EditOverlay, pass !isNewSelection as replaceExisting (replace only if continuing drag)
        editOverlayRef.current.insertAtCursor(reference, !isNewSelection);
      } else if (isOnDifferentSheet && floatingInputRef.current) {
        // Use floating input for insertion when on different sheet
        const input = floatingInputRef.current;
        const currentValue = editValue;
        
        let start: number;
        let end: number;
        
        // If this is a new selection, clear the last insert ref and use cursor position
        if (isNewSelection) {
          lastInsertRef.current = null;
        }
        
        if (lastInsertRef.current) {
          // Replace the previously inserted reference (for drag)
          start = lastInsertRef.current.start;
          end = lastInsertRef.current.end;
        } else {
          // Insert at cursor position
          start = input.selectionStart ?? currentValue.length;
          end = input.selectionEnd ?? start;
        }
        
        // Create new value with inserted text
        const newValue = currentValue.slice(0, start) + reference + currentValue.slice(end);
        setEditValue(newValue);
        
        // Track the inserted reference position
        lastInsertRef.current = { start, end: start + reference.length };
        
        // Update cursor position
        requestAnimationFrame(() => {
          if (floatingInputRef.current) {
            const newCursorPos = start + reference.length;
            floatingInputRef.current.focus();
            floatingInputRef.current.setSelectionRange(newCursorPos, newCursorPos);
          }
        });
      }
    }
  }, [isEditingFormula, isOnDifferentSheet, editValue]);

  // Handle selection change
  const handleSelectionChange = useCallback((selection: Selection) => {
    workbook.setSelection(selection);
  }, [workbook]);

  // Handle direct sort by direction (for context menu)
  const handleSortColumnByDirection = useCallback((column: number, direction: 'asc' | 'desc') => {
    // Sorting would scatter merged regions — refuse, matching Excel/Sheets.
    if (workbook.getSheet().getMergedRegions().length > 0) {
      window.alert('Cannot sort a sheet that contains merged cells. Unmerge them first.');
      return;
    }

    const sortOrder: SortOrder[] = [{ column, direction }];

    workbook.setSortOrder(sortOrder);
    workbook.sortSheet();

    // Trigger re-render to show updated sort indicators
    setDimensionVersion(v => v + 1);
  }, [workbook]);

  // Handle context menu open
  const handleContextMenu = useCallback((menu: ContextMenuType) => {
    setContextMenu(menu);
  }, []);

  // Handle context menu close
  const handleContextMenuClose = useCallback(() => {
    setContextMenu(null);
  }, []);

  // Handle clipboard handlers from CanvasGrid
  const handleClipboardReady = useCallback((handlers: {
    copy: () => void;
    cut: () => void;
    paste: (targetCell?: CellPosition) => Promise<void>;
  }) => {
    clipboardHandlersRef.current = handlers;
  }, []);

  // Find & replace -----------------------------------------------------------
  const handleNavigateReady = useCallback((navigate: (cell: CellPosition) => void) => {
    navigateRef.current = navigate;
  }, []);

  // dimensionVersion is a dep so matches refresh after edits and replacements.
  const findReplaceMatches = useMemo(() => {
    if (!findReplace.isOpen || !findReplace.query) return [];
    return findMatches(workbook.getSheet(), findReplace.query, {
      matchCase: findReplace.matchCase,
      wholeCell: findReplace.wholeCell,
      searchFormulas: findReplace.searchFormulas,
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [findReplace, workbook, dimensionVersion]);

  // A changed match set invalidates the current position.
  useEffect(() => {
    setActiveMatchIndex(-1);
  }, [findReplaceMatches]);

  const handleFindReplaceChange = useCallback((patch: Partial<FindReplaceState>) => {
    setFindReplace((fr) => ({ ...fr, ...patch }));
  }, []);

  const handleCloseFindReplace = useCallback(() => {
    setFindReplace((fr) => ({ ...fr, isOpen: false }));
  }, []);

  const handleFindNext = useCallback(() => {
    if (findReplaceMatches.length === 0) return;
    const next = (activeMatchIndex + 1) % findReplaceMatches.length;
    setActiveMatchIndex(next);
    const m = findReplaceMatches[next];
    navigateRef.current?.({ row: m.row, col: m.col });
  }, [findReplaceMatches, activeMatchIndex]);

  const handleFindPrev = useCallback(() => {
    const n = findReplaceMatches.length;
    if (n === 0) return;
    const prev = (activeMatchIndex - 1 + n) % n;
    setActiveMatchIndex(prev);
    const m = findReplaceMatches[prev];
    navigateRef.current?.({ row: m.row, col: m.col });
  }, [findReplaceMatches, activeMatchIndex]);

  const handleReplace = useCallback(() => {
    if (!findReplace.query || findReplaceMatches.length === 0) return;
    const idx = activeMatchIndex >= 0 ? activeMatchIndex : 0;
    const m = findReplaceMatches[idx];
    if (!m) return;
    const result = computeReplacement(
      workbook.getSheet().getCell(m.row, m.col),
      findReplace.query,
      findReplace.replacement,
      {
        matchCase: findReplace.matchCase,
        wholeCell: findReplace.wholeCell,
        searchFormulas: findReplace.searchFormulas,
      }
    );
    if (result.kind === 'none') return;
    workbook.batch(() => {
      if (result.kind === 'value') {
        workbook.setCellValue(undefined, m.row, m.col, result.value);
      } else if (result.kind === 'formula') {
        workbook.setFormula(undefined, m.row, m.col, result.formula);
      }
    });
    setDimensionVersion((v) => v + 1);
  }, [findReplace, findReplaceMatches, activeMatchIndex, workbook]);

  const handleReplaceAll = useCallback(() => {
    if (!findReplace.query) return;
    const opts = {
      matchCase: findReplace.matchCase,
      wholeCell: findReplace.wholeCell,
      searchFormulas: findReplace.searchFormulas,
    };
    const sheet = workbook.getSheet();
    const matches = findMatches(sheet, findReplace.query, opts);
    if (matches.length === 0) return;
    workbook.batch(() => {
      for (const m of matches) {
        const result = computeReplacement(
          sheet.getCell(m.row, m.col),
          findReplace.query,
          findReplace.replacement,
          opts
        );
        if (result.kind === 'value') {
          workbook.setCellValue(undefined, m.row, m.col, result.value);
        } else if (result.kind === 'formula') {
          workbook.setFormula(undefined, m.row, m.col, result.formula);
        }
      }
    });
    setDimensionVersion((v) => v + 1);
  }, [findReplace, workbook]);

  // Cancel edit
  const cancelEdit = useCallback(() => {
    setEditingCell(null);
    setEditValue('');
    setEditingCellFormat(undefined);
    setOriginalEditingSheetId(null);
    lastInsertRef.current = null;
    
    // Refocus the canvas so keyboard navigation continues to work
    // Use preventScroll to avoid the UI jumping when focus changes
    requestAnimationFrame(() => {
      const canvas = canvasGridRef.current?.querySelector('canvas');
      canvas?.focus({ preventScroll: true });
    });
  }, []);

  // Calculate edit overlay position
  const getEditOverlayPosition = useCallback(() => {
    if (!editingCell) return null;
    
    const sheet = workbook.getSheet();
    
    // Calculate x position
    let x = colWidth; // Start after row header
    for (let c = 0; c < editingCell.col; c++) {
      x += sheet.getColWidth(c);
    }
    
    // Calculate y position
    let y = headerHeight; // Start after column header
    for (let r = 0; r < editingCell.row; r++) {
      y += sheet.getRowHeight(r);
    }
    
    const cellWidth = sheet.getColWidth(editingCell.col);
    const cellHeight = sheet.getRowHeight(editingCell.row);
    
    return { x, y, width: cellWidth, height: cellHeight };
  }, [editingCell, workbook, colWidth, headerHeight]);

  // Helper function to apply style to selection
  const applyStyleToSelection = useCallback(
    (styleUpdater: (currentStyle: CellStyle) => CellStyle) => {
      const selection = workbook.getSelection();
      if (selection.ranges.length > 0) {
        // Record history before style changes for undo/redo
        workbook.recordHistory();
        
        const stylePool = workbook.getStylePool();
        selection.ranges.forEach((range) => {
          const minRow = Math.min(range.startRow, range.endRow);
          const maxRow = Math.max(range.startRow, range.endRow);
          const minCol = Math.min(range.startCol, range.endCol);
          const maxCol = Math.max(range.startCol, range.endCol);
          for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
              const cell = workbook.getCell(undefined, r, c);
              const currentStyle = cell?.styleId ? stylePool.get(cell.styleId) : {};
              const newStyle = styleUpdater(currentStyle as CellStyle);
              const styleId = stylePool.getOrCreate(newStyle);
              workbook.setCell(undefined, r, c, { styleId });
            }
          }
        });
        // Trigger re-render
        setDimensionVersion(v => v + 1);
      }
    },
    [workbook]
  );

  // Helper function to apply format to selection
  const applyFormatToSelection = useCallback(
    (format: CellFormat) => {
      const selection = workbook.getSelection();
      if (selection.ranges.length > 0) {
        // Record history before format changes for undo/redo
        workbook.recordHistory();

        selection.ranges.forEach((range) => {
          const minRow = Math.min(range.startRow, range.endRow);
          const maxRow = Math.max(range.startRow, range.endRow);
          const minCol = Math.min(range.startCol, range.endCol);
          const maxCol = Math.max(range.startCol, range.endCol);
          for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
              const cell = workbook.getCell(undefined, r, c);
              const updatedCell = { ...cell, format };

              // If applying date formatting to a string value, try to parse it as a date
              if (format.type && ['date', 'time', 'datetime'].includes(format.type) &&
                  cell && typeof cell.value === 'string') {
                const dateSerial = parseDateString(cell.value);
                if (dateSerial !== null) {
                  updatedCell.value = dateSerial;
                }
              }

              workbook.setCell(undefined, r, c, updatedCell);
            }
          }
        });
        // Trigger re-render
        setDimensionVersion(v => v + 1);
      }
    },
    [workbook]
  );

  const editPosition = getEditOverlayPosition();
  
  // Container ref for checking if keyboard events are within the workbook
  const containerRef = useRef<HTMLDivElement>(null);

  // Track if the workbook was recently interacted with (for handling shortcuts when focus is on body)
  const wasRecentlyActiveRef = useRef(false);
  
  // Mark as active when user interacts with the workbook
  const handleContainerInteraction = useCallback(() => {
    wasRecentlyActiveRef.current = true;
  }, []);
  
  // Global keyboard handler for all shortcuts
  // This ensures keyboard shortcuts work even when focus is on toolbar elements
  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement;
      const isCtrl = e.ctrlKey || e.metaKey;
      const key = e.key.toLowerCase();
      const shiftKey = e.shiftKey;
      
      // Check if target is within our container
      const isInContainer = containerRef.current?.contains(target);
      
      // Also handle when focus is on body/document AND workbook was recently active
      // This covers the case when toolbar buttons lose focus after being clicked
      const isOnBody = target === document.body || target === document.documentElement;
      const shouldHandle = isInContainer || (isOnBody && wasRecentlyActiveRef.current);
      
      if (!shouldHandle) {
        return;
      }

      // Never hijack keys while the user is typing in a text field — the
      // formula bar, comment composer, dialogs, etc. own their own input.
      // Without this, Backspace/Delete would clear grid cells instead of
      // deleting characters.
      if (
        target.tagName === 'INPUT' ||
        target.tagName === 'TEXTAREA' ||
        target.isContentEditable
      ) {
        return;
      }

      // Skip if we're editing a cell (let the editor handle it)
      // Exception: Ctrl+Z/Y should still work for undo/redo during editing in some cases
      if (editingCell) {
        // Only allow undo/redo shortcuts during editing
        if (isCtrl && (key === 'z' || key === 'y')) {
          // Let the edit overlay handle these
        }
        return;
      }
      
      // Handle modifier key shortcuts (Ctrl/Cmd + key)
      if (isCtrl) {
        switch (key) {
          case 'z':
            e.preventDefault();
            if (shiftKey) {
              // Ctrl+Shift+Z to redo
              workbook.redo();
            } else {
              // Ctrl+Z to undo
              workbook.undo();
            }
            setDimensionVersion(v => v + 1);
            break;
            
          case 'y':
            // Ctrl+Y to redo
            e.preventDefault();
            workbook.redo();
            setDimensionVersion(v => v + 1);
            break;
            
          case 'c':
            // Ctrl+C to copy
            e.preventDefault();
            clipboardHandlersRef.current?.copy();
            break;
            
          case 'v':
            // Ctrl+V to paste
            e.preventDefault();
            clipboardHandlersRef.current?.paste();
            break;
            
          case 'x':
            // Ctrl+X to cut
            e.preventDefault();
            clipboardHandlersRef.current?.cut();
            setDimensionVersion(v => v + 1);
            break;
            
          case 'b':
            // Ctrl+B to toggle bold
            e.preventDefault();
            applyStyleToSelection((s: CellStyle) => ({ ...s, bold: !s?.bold }));
            break;
            
          case 'i':
            // Ctrl+I to toggle italic
            e.preventDefault();
            applyStyleToSelection((s: CellStyle) => ({ ...s, italic: !s?.italic }));
            break;
            
          case 'u':
            // Ctrl+U to toggle underline
            e.preventDefault();
            applyStyleToSelection((s: CellStyle) => ({
              ...s,
              textDecoration: s?.textDecoration === 'underline' ? 'none' : 'underline',
            }));
            break;
            
          case 'f':
            // Ctrl+F to open find & replace
            e.preventDefault();
            setFindReplace((fr) => ({ ...fr, isOpen: true }));
            break;

          case 'a':
            // Ctrl+A to select all - only if focus is in container (not body)
            if (isInContainer) {
              e.preventDefault();
              const sheet = workbook.getSheet();
              const newSelection: Selection = {
                ranges: [{
                  startRow: 0,
                  startCol: 0,
                  endRow: sheet.rowCount - 1,
                  endCol: sheet.colCount - 1,
                }],
                activeCell: { row: 0, col: 0 },
              };
              workbook.setSelection(newSelection);
              setDimensionVersion(v => v + 1);
            }
            break;
        }
        return;
      }
      
      // Handle Delete/Backspace to clear cells (non-modifier)
      if (key === 'delete' || key === 'backspace') {
        // Only handle if focus is explicitly in container (not when on body, to avoid conflicts)
        if (isInContainer) {
          const selection = workbook.getSelection();
          if (selection.ranges.length > 0) {
            e.preventDefault();
            workbook.batch(() => {
              selection.ranges.forEach((range) => {
                const minRow = Math.min(range.startRow, range.endRow);
                const maxRow = Math.max(range.startRow, range.endRow);
                const minCol = Math.min(range.startCol, range.endCol);
                const maxCol = Math.max(range.startCol, range.endCol);
                for (let r = minRow; r <= maxRow; r++) {
                  for (let c = minCol; c <= maxCol; c++) {
                    workbook.setCellValue(undefined, r, c, null);
                  }
                }
              });
            });
            setDimensionVersion(v => v + 1);
          }
        }
      }
    };
    
    // Reset "recently active" flag when focus moves outside the container
    const handleFocusOut = (e: FocusEvent) => {
      const relatedTarget = e.relatedTarget as HTMLElement | null;
      if (relatedTarget && !containerRef.current?.contains(relatedTarget)) {
        // Focus moved outside, but give a grace period for toolbar actions
        setTimeout(() => {
          if (!containerRef.current?.contains(document.activeElement)) {
            wasRecentlyActiveRef.current = false;
          }
        }, 100);
      }
    };
    
    const container = containerRef.current;
    
    document.addEventListener('keydown', handleGlobalKeyDown);
    container?.addEventListener('focusout', handleFocusOut);
    
    return () => {
      document.removeEventListener('keydown', handleGlobalKeyDown);
      container?.removeEventListener('focusout', handleFocusOut);
    };
  }, [workbook, editingCell, applyStyleToSelection]);

  return (
    <div
      ref={containerRef}
      className={className}
      style={{
        ...style,
        display: 'flex',
        flexDirection: 'column',
        width,
        height,
        border: '1px solid rgba(15, 23, 42, 0.08)',
        background: 'linear-gradient(180deg, #f8fafc 0%, #eef2f6 100%)',
        borderRadius: '14px',
        boxShadow: '0 1px 3px rgba(15, 23, 42, 0.04), 0 12px 32px -8px rgba(15, 23, 42, 0.12)',
        overflow: 'hidden',
        position: 'relative',
        fontFamily: '"Inter", "SF Pro Display", -apple-system, BlinkMacSystemFont, sans-serif',
      }}
      onMouseDown={handleContainerInteraction}
      onFocus={handleContainerInteraction}
    >
      {/* Toolbar */}
      <Toolbar
        onUndo={() => {
          workbook.undo();
          setDimensionVersion(v => v + 1);
        }}
        onRedo={() => {
          workbook.redo();
          setDimensionVersion(v => v + 1);
        }}
        onFreezeRows={(rows) => {
          workbook.recordHistory();
          const sheet = workbook.getSheet();
          sheet.setFrozenRows(rows);
          setDimensionVersion(v => v + 1);
        }}
        onFreezeCols={(cols) => {
          workbook.recordHistory();
          const sheet = workbook.getSheet();
          sheet.setFrozenCols(cols);
          setDimensionVersion(v => v + 1);
        }}
        onUnfreeze={() => {
          workbook.recordHistory();
          const sheet = workbook.getSheet();
          sheet.clearFreeze();
          setDimensionVersion(v => v + 1);
        }}
        frozenRows={workbook.getSheet().getFrozenRows()}
        frozenCols={workbook.getSheet().getFrozenCols()}
        activeColumn={activeCell?.col}
        onSortColumn={handleSortColumnByDirection}
        onFilterColumn={(column) => {
          const sheet = workbook.getSheet();
          const existingFilter = sheet.getFilters().get(column);
          setFilterModal({
            isOpen: true,
            column,
            existingFilter,
          });
        }}
        onToggleComments={() => setShowComments((v) => !v)}
        commentsActive={showComments}
        onBold={() => applyStyleToSelection((s) => ({ ...s, bold: !s?.bold }))}
        onItalic={() => applyStyleToSelection((s) => ({ ...s, italic: !s?.italic }))}
        onUnderline={() =>
          applyStyleToSelection((s) => ({
            ...s,
            textDecoration: s?.textDecoration === 'underline' ? 'none' : 'underline',
          }))
        }
        onStrikethrough={() =>
          applyStyleToSelection((s) => ({
            ...s,
            textDecoration: s?.textDecoration === 'line-through' ? 'none' : 'line-through',
          }))
        }
        onFontFamily={(fontFamily) => applyStyleToSelection((s) => ({ ...s, fontFamily }))}
        onFontSize={(fontSize) => applyStyleToSelection((s) => ({ ...s, fontSize }))}
        onFontColor={(fontColor) => applyStyleToSelection((s) => ({ ...s, fontColor }))}
        onBackgroundColor={(backgroundColor) => applyStyleToSelection((s) => ({ ...s, backgroundColor }))}
        onBorder={(border) => {
          applyStyleToSelection((s) => {
            const newStyle: CellStyle = { ...s };
            if (border === 'none') {
              delete newStyle.borderTop;
              delete newStyle.borderRight;
              delete newStyle.borderBottom;
              delete newStyle.borderLeft;
            } else if (border === 'all') {
              newStyle.borderTop = '1px solid #000000';
              newStyle.borderRight = '1px solid #000000';
              newStyle.borderBottom = '1px solid #000000';
              newStyle.borderLeft = '1px solid #000000';
            } else {
              const borderKey = `border${border.charAt(0).toUpperCase() + border.slice(1)}` as 'borderTop' | 'borderRight' | 'borderBottom' | 'borderLeft';
              newStyle[borderKey] = '1px solid #000000';
            }
            return newStyle;
          });
        }}
        onAlignLeft={() => applyStyleToSelection((s) => ({ ...s, textAlign: 'left' as const }))}
        onAlignCenter={() => applyStyleToSelection((s) => ({ ...s, textAlign: 'center' as const }))}
        onAlignRight={() => applyStyleToSelection((s) => ({ ...s, textAlign: 'right' as const }))}
        onVerticalAlign={(align) => applyStyleToSelection((s) => ({ ...s, verticalAlign: align }))}
        onTextWrap={() => applyStyleToSelection((s) => ({ ...s, textWrap: !s?.textWrap }))}
        onTextRotation={(angle) =>
          applyStyleToSelection((s) => ({ ...s, textRotation: angle }))
        }
        onFormatCurrency={() => applyFormatToSelection({ type: 'currency' })}
        onFormatPercentage={() => applyFormatToSelection({ type: 'percentage' })}
        onFormatNumber={() => applyFormatToSelection({ type: 'number' })}
        onMergeCells={() => {
          const sel = workbook.getSelection();
          if (sel.ranges.length === 0) return;
          const raw = sel.ranges[0];
          const range = {
            startRow: Math.min(raw.startRow, raw.endRow),
            endRow: Math.max(raw.startRow, raw.endRow),
            startCol: Math.min(raw.startCol, raw.endCol),
            endCol: Math.max(raw.startCol, raw.endCol),
          };
          const sheet = workbook.getSheet();
          // Toggle: unmerge if the selection touches a merge, else merge.
          const overlapsMerge = sheet.getMergedRegions().some((m) =>
            m.startRow <= range.endRow && m.endRow >= range.startRow &&
            m.startCol <= range.endCol && m.endCol >= range.startCol
          );
          if (overlapsMerge) {
            workbook.unmergeCells(range);
          } else {
            workbook.mergeCells(range);
          }
          setDimensionVersion(v => v + 1);
        }}
        onHyperlink={(url) => {
          const selection = workbook.getSelection();
          if (selection.ranges.length > 0) {
            // Record history before hyperlink change for undo/redo
            workbook.recordHistory();
            
            const range = selection.ranges[0];
            const minRow = Math.min(range.startRow, range.endRow);
            const minCol = Math.min(range.startCol, range.endCol);
            const cell = workbook.getCell(undefined, minRow, minCol);
            workbook.setCell(undefined, minRow, minCol, {
              ...cell,
              hyperlink: url,
            });
            setDimensionVersion(v => v + 1);
          }
        }}
        selectedFormat={
          activeCell
            ? (() => {
                const cell = workbook.getCell(undefined, activeCell.row, activeCell.col);
                const stylePool = workbook.getStylePool();
                const formatPool = workbook.getFormatPool();
                const style = cell?.styleId ? stylePool.get(cell.styleId) : {};
                const format = cell?.formatId ? formatPool.get(cell.formatId) : { type: 'text' as FormatType };
                return {
                  bold: style?.bold,
                  italic: style?.italic,
                  underline: style?.textDecoration === 'underline',
                  strikethrough: style?.textDecoration === 'line-through',
                  fontFamily: style?.fontFamily,
                  fontSize: style?.fontSize,
                  fontColor: style?.fontColor,
                  backgroundColor: style?.backgroundColor,
                  align: style?.textAlign,
                  verticalAlign: style?.verticalAlign,
                  textWrap: style?.textWrap,
                  textRotation: style?.textRotation,
                  format: format?.type,
                  hyperlink: cell?.hyperlink,
                };
              })()
            : undefined
        }
      />

      {/* Formula Bar */}
      <FormulaBar activeCell={activeCell} onFormulaChange={handleFormulaChange} />

      {/* Canvas Grid Area */}
      <div
        ref={canvasGridRef}
        style={{
          position: 'relative',
          height: canvasAreaHeight,
          flexShrink: 0,
        }}
      >
        <CanvasGrid
          width={canvasAreaWidth}
          height={canvasAreaHeight}
          rowHeight={rowHeight}
          colWidth={colWidth}
          headerHeight={headerHeight}
          activeCell={activeCell}
          onActiveCellChange={handleActiveCellChange}
          onCellEdit={handleCellEdit}
          onSelectionChange={handleSelectionChange}
          onContentChange={() => setDimensionVersion(v => v + 1)}
          dimensionVersion={dimensionVersion}
          editValue={editValue}
          isEditingFormula={isEditingFormula}
          onInsertCellReference={handleInsertCellReference}
          editingCell={editingCell}
          originalEditingSheetId={originalEditingSheetId ?? undefined}
          onContextMenu={handleContextMenu}
          onClipboardReady={handleClipboardReady}
          onNavigateReady={handleNavigateReady}
        />
        
        {/* Edit Overlay - only show when on the original sheet */}
        {editingCell && editPosition && !isOnDifferentSheet && (
          <EditOverlay
            ref={editOverlayRef}
            value={editValue}
            onChange={setEditValue}
            onCommit={commitEdit}
            onCancel={cancelEdit}
            x={editPosition.x}
            y={editPosition.y}
            width={editPosition.width}
            height={editPosition.height}
            isEditingFormula={isEditingFormula}
            cellFormat={editingCellFormat}
            interceptKeyDown={formulaAssist.onKeyDown}
            onSelectionChange={setCaret}
          />
        )}

        {/* Formula autocomplete / parameter help, anchored below the editor */}
        {editingCell && editPosition && !isOnDifferentSheet && formulaAssist.mode === 'autocomplete' && (
          <FormulaAutocomplete
            suggestions={formulaAssist.suggestions}
            highlightedIndex={formulaAssist.highlightedIndex}
            top={editPosition.y + editPosition.height}
            left={editPosition.x}
            onHover={formulaAssist.setHighlightedIndex}
            onPick={formulaAssist.accept}
          />
        )}
        {editingCell &&
          editPosition &&
          !isOnDifferentSheet &&
          formulaAssist.mode === 'signature' &&
          formulaAssist.signature && (
            <FormulaSignatureHint
              doc={formulaAssist.signature.doc}
              activeArg={formulaAssist.signature.activeArg}
              top={editPosition.y + editPosition.height}
              left={editPosition.x}
            />
          )}
        
        {/* Floating Formula Input - shown when editing formula on a different sheet */}
        {isEditingFormula && isOnDifferentSheet && editingCell && floatingInputPos && (
          <div
            className="floating-formula-input"
            style={{
              position: 'absolute',
              top: floatingInputPos.y,
              left: floatingInputPos.x,
              width: 280,
              zIndex: 1000,
              background: 'rgba(255, 255, 255, 0.98)',
              backdropFilter: 'blur(20px)',
              WebkitBackdropFilter: 'blur(20px)',
              border: '1px solid rgba(99, 102, 241, 0.35)',
              borderRadius: '12px',
              boxShadow:
                '0 10px 15px -3px rgba(15, 23, 42, 0.1), 0 20px 25px -5px rgba(15, 23, 42, 0.08), 0 0 0 1px rgba(15, 23, 42, 0.04)',
              overflow: 'hidden',
              fontFamily: '"Inter", "SF Pro Display", -apple-system, BlinkMacSystemFont, sans-serif',
            }}
          >
            {/* Drag handle */}
            <div
              onMouseDown={handleFloatingDragStart}
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                padding: '7px 12px',
                background: 'linear-gradient(135deg, #6366f1 0%, #818cf8 100%)',
                color: '#ffffff',
                cursor: isDragging ? 'grabbing' : 'grab',
                fontSize: '12px',
                fontWeight: 600,
                userSelect: 'none',
              }}
            >
              <span>
                {(() => {
                  const originalSheet = workbook.getSheet(originalEditingSheetId!);
                  return originalSheet?.name || 'Sheet';
                })()}!{columnIndexToLabel(editingCell.col)}{editingCell.row + 1}
              </span>
              <span style={{ opacity: 0.85, fontSize: '10px', fontWeight: 500 }}>⋮⋮ drag</span>
            </div>

            {/* Input area */}
            <div style={{ padding: '10px' }}>
              <input
                ref={floatingInputRef}
                type="text"
                value={editValue}
                onChange={(e) => {
                  lastInsertRef.current = null;
                  setEditValue(e.target.value);
                }}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    lastInsertRef.current = null;
                    commitEdit(e.currentTarget.value);
                    if (originalEditingSheetId) {
                      workbook.setActiveSheet(originalEditingSheetId);
                      setDimensionVersion(v => v + 1);
                    }
                  } else if (e.key === 'Escape') {
                    e.preventDefault();
                    lastInsertRef.current = null;
                    cancelEdit();
                    if (originalEditingSheetId) {
                      workbook.setActiveSheet(originalEditingSheetId);
                      setDimensionVersion(v => v + 1);
                    }
                  }
                }}
                onBlur={(e) => {
                  const relatedTarget = e.relatedTarget as HTMLElement;
                  if (relatedTarget?.tagName === 'CANVAS') {
                    setTimeout(() => {
                      if (floatingInputRef.current && isEditingFormula) {
                        floatingInputRef.current.focus();
                      }
                    }, 10);
                    return;
                  }
                  if (relatedTarget?.closest?.('.floating-formula-input')) {
                    return;
                  }
                }}
                onClick={(e) => e.stopPropagation()}
                onMouseDown={(e) => e.stopPropagation()}
                style={{
                  width: '100%',
                  border: '1px solid #6366f1',
                  borderRadius: '8px',
                  padding: '7px 10px',
                  fontSize: '13px',
                  outline: 'none',
                  boxSizing: 'border-box',
                  color: '#1e293b',
                  fontFamily: 'inherit',
                  boxShadow: '0 0 0 3px rgba(99, 102, 241, 0.12)',
                }}
                autoFocus
              />
              <div style={{ fontSize: '10px', color: '#94a3b8', marginTop: '6px' }}>
                Click cells to add. Enter ✓ | Esc ✗
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Sheet Tabs */}
      <div style={{ 
        position: 'absolute', 
        bottom: '28px', 
        left: 0, 
        right: 0, 
        zIndex: 10 
      }}>
        <SheetTabs
          onSheetSelect={(sheetId) => {
            workbook.setActiveSheet(sheetId);
            setDimensionVersion(v => v + 1);
          }}
        />
      </div>

      {/* Comments panel — floating drawer above the sheet tabs */}
      {showComments && (
        <div
          style={{
            position: 'absolute',
            top: toolbarHeight + formulaBarHeight,
            right: 0,
            bottom: 66,
            width: 320,
            zIndex: 30,
          }}
        >
          <CommentsPanel activeCell={activeCell} onClose={() => setShowComments(false)} />
        </div>
      )}

      {/* Cell Context Menu */}
      {contextMenu && contextMenu.type === 'cell' && (
        <ContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          onClose={handleContextMenuClose}
          onCopy={() => {
            clipboardHandlersRef.current?.copy();
            handleContextMenuClose();
          }}
          onPaste={() => {
            clipboardHandlersRef.current?.paste(contextMenu.cell);
            handleContextMenuClose();
          }}
          onCut={() => {
            clipboardHandlersRef.current?.cut();
            handleContextMenuClose();
          }}
          onDelete={() => {
            const selection = workbook.getSelection();
            if (selection.ranges.length > 0) {
              // Use batch to record single history entry for all deletions
              workbook.batch(() => {
                selection.ranges.forEach((range) => {
                  const minRow = Math.min(range.startRow, range.endRow);
                  const maxRow = Math.max(range.startRow, range.endRow);
                  const minCol = Math.min(range.startCol, range.endCol);
                  const maxCol = Math.max(range.startCol, range.endCol);
                  for (let r = minRow; r <= maxRow; r++) {
                    for (let c = minCol; c <= maxCol; c++) {
                      workbook.setCellValue(undefined, r, c, null);
                    }
                  }
                });
              });
              setDimensionVersion(v => v + 1);
            }
            handleContextMenuClose();
          }}
          onInsertRow={() => {
            const selection = workbook.getSelection();
            if (selection.ranges.length > 0) {
              workbook.recordHistory();
              const range = selection.ranges[0];
              const targetRow = Math.min(range.startRow, range.endRow);
              const sheet = workbook.getSheet();
              sheet.insertRows(targetRow, 1);
              setDimensionVersion(v => v + 1);
            }
            handleContextMenuClose();
          }}
          onInsertRowBelow={() => {
            const selection = workbook.getSelection();
            if (selection.ranges.length > 0) {
              workbook.recordHistory();
              const range = selection.ranges[0];
              const maxRow = Math.max(range.startRow, range.endRow);
              const sheet = workbook.getSheet();
              sheet.insertRows(maxRow + 1, 1);
              setDimensionVersion(v => v + 1);
            }
            handleContextMenuClose();
          }}
          onInsertColumn={() => {
            const selection = workbook.getSelection();
            if (selection.ranges.length > 0) {
              workbook.recordHistory();
              const range = selection.ranges[0];
              const targetCol = Math.min(range.startCol, range.endCol);
              const sheet = workbook.getSheet();
              sheet.insertCols(targetCol, 1);
              setDimensionVersion(v => v + 1);
            }
            handleContextMenuClose();
          }}
          onInsertColumnRight={() => {
            const selection = workbook.getSelection();
            if (selection.ranges.length > 0) {
              workbook.recordHistory();
              const range = selection.ranges[0];
              const maxCol = Math.max(range.startCol, range.endCol);
              const sheet = workbook.getSheet();
              sheet.insertCols(maxCol + 1, 1);
              setDimensionVersion(v => v + 1);
            }
            handleContextMenuClose();
          }}
          onDeleteRow={() => {
            const selection = workbook.getSelection();
            if (selection.ranges.length > 0) {
              workbook.recordHistory();
              const range = selection.ranges[0];
              const minRow = Math.min(range.startRow, range.endRow);
              const maxRow = Math.max(range.startRow, range.endRow);
              const sheet = workbook.getSheet();
              sheet.deleteRows(minRow, maxRow - minRow + 1);
              setDimensionVersion(v => v + 1);
            }
            handleContextMenuClose();
          }}
          onDeleteColumn={() => {
            const selection = workbook.getSelection();
            if (selection.ranges.length > 0) {
              workbook.recordHistory();
              const range = selection.ranges[0];
              const minCol = Math.min(range.startCol, range.endCol);
              const maxCol = Math.max(range.startCol, range.endCol);
              const sheet = workbook.getSheet();
              sheet.deleteCols(minCol, maxCol - minCol + 1);
              setDimensionVersion(v => v + 1);
            }
            handleContextMenuClose();
          }}
          onFormat={() => {
            // Get current format from selection
            const selection = workbook.getSelection();
            let currentFormat: CellFormat | undefined;
            let sampleValue = 1234.56;

            if (selection.ranges.length > 0) {
              // Get format from first cell in selection
              const range = selection.ranges[0];
              const cell = workbook.getCell(undefined, range.startRow, range.startCol);
              const formatPool = workbook.getFormatPool();
              currentFormat = cell?.formatId ? formatPool.get(cell.formatId) : undefined;

              // Get sample value from first cell
              if (cell?.value && typeof cell.value === 'number') {
                sampleValue = cell.value;
              }
            }

            setFormatModal({
              isOpen: true,
              currentFormat,
              sampleValue,
            });
            handleContextMenuClose();
          }}
          onComment={() => {
            setActiveCell(contextMenu.cell);
            setShowComments(true);
            handleContextMenuClose();
          }}
        />
      )}

      {/* Column Header Context Menu */}
      {contextMenu && contextMenu.type === 'column' && (
        <HeaderContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          type="column"
          onClose={handleContextMenuClose}
          isHidden={workbook.getSheet().isColHidden(contextMenu.index)}
          hasHiddenAdjacent={
            (contextMenu.index > 0 && workbook.getSheet().isColHidden(contextMenu.index - 1)) ||
            (contextMenu.index < workbook.getSheet().colCount - 1 && workbook.getSheet().isColHidden(contextMenu.index + 1))
          }
          onHide={() => {
            workbook.recordHistory();
            workbook.getSheet().hideCol(contextMenu.index);
            setDimensionVersion(v => v + 1);
            handleContextMenuClose();
          }}
          onShow={() => {
            workbook.recordHistory();
            workbook.getSheet().showCol(contextMenu.index);
            setDimensionVersion(v => v + 1);
            handleContextMenuClose();
          }}
          onUnhideAdjacent={() => {
            workbook.recordHistory();
            const sheet = workbook.getSheet();
            // Unhide columns around the selected column
            const startCol = Math.max(0, contextMenu.index - 1);
            const endCol = Math.min(sheet.colCount - 1, contextMenu.index + 1);
            sheet.showColsInRange(startCol, endCol);
            setDimensionVersion(v => v + 1);
            handleContextMenuClose();
          }}
          onInsert={() => {
            workbook.recordHistory();
            workbook.getSheet().insertCols(contextMenu.index, 1);
            setDimensionVersion(v => v + 1);
            handleContextMenuClose();
          }}
          onDelete={() => {
            workbook.recordHistory();
            workbook.getSheet().deleteCols(contextMenu.index, 1);
            setDimensionVersion(v => v + 1);
            handleContextMenuClose();
          }}
          onSortAsc={() => {
            handleSortColumnByDirection(contextMenu.index, 'asc');
            handleContextMenuClose();
          }}
          onSortDesc={() => {
            handleSortColumnByDirection(contextMenu.index, 'desc');
            handleContextMenuClose();
          }}
          onFilter={() => {
            const sheet = workbook.getSheet();
            const existingFilter = sheet.getFilters().get(contextMenu.index);
            setFilterModal({
              isOpen: true,
              column: contextMenu.index,
              existingFilter,
            });
            handleContextMenuClose();
          }}
        />
      )}

      {/* Row Header Context Menu */}
      {contextMenu && contextMenu.type === 'row' && (
        <HeaderContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          type="row"
          onClose={handleContextMenuClose}
          isHidden={workbook.getSheet().isRowHidden(contextMenu.index)}
          hasHiddenAdjacent={
            (contextMenu.index > 0 && workbook.getSheet().isRowHidden(contextMenu.index - 1)) ||
            (contextMenu.index < workbook.getSheet().rowCount - 1 && workbook.getSheet().isRowHidden(contextMenu.index + 1))
          }
          onHide={() => {
            workbook.recordHistory();
            workbook.getSheet().hideRow(contextMenu.index);
            setDimensionVersion(v => v + 1);
            handleContextMenuClose();
          }}
          onShow={() => {
            workbook.recordHistory();
            workbook.getSheet().showRow(contextMenu.index);
            setDimensionVersion(v => v + 1);
            handleContextMenuClose();
          }}
          onUnhideAdjacent={() => {
            workbook.recordHistory();
            const sheet = workbook.getSheet();
            // Unhide rows around the selected row
            const startRow = Math.max(0, contextMenu.index - 1);
            const endRow = Math.min(sheet.rowCount - 1, contextMenu.index + 1);
            sheet.showRowsInRange(startRow, endRow);
            setDimensionVersion(v => v + 1);
            handleContextMenuClose();
          }}
          onInsert={() => {
            workbook.recordHistory();
            workbook.getSheet().insertRows(contextMenu.index, 1);
            setDimensionVersion(v => v + 1);
            handleContextMenuClose();
          }}
          onDelete={() => {
            workbook.recordHistory();
            workbook.getSheet().deleteRows(contextMenu.index, 1);
            setDimensionVersion(v => v + 1);
            handleContextMenuClose();
          }}
        />
      )}

      {/* Filter Modal */}
      {filterModal && (
        <FilterModal
          isOpen={filterModal.isOpen}
          sheet={workbook.getSheet()}
          column={filterModal.column}
          existingFilter={filterModal.existingFilter}
          onClose={() => setFilterModal(null)}
          onApply={(filter) => {
            workbook.setFilter(filter.column, filter);
            setDimensionVersion(v => v + 1);
            setFilterModal(null);
          }}
          onClear={() => {
            workbook.clearFilter(filterModal.column);
            setDimensionVersion(v => v + 1);
            setFilterModal(null);
          }}
        />
      )}

      {/* Format Cells Modal */}
      {formatModal && (
        <FormatCellsModal
          isOpen={formatModal.isOpen}
          currentFormat={formatModal.currentFormat}
          sampleValue={formatModal.sampleValue}
          onClose={() => setFormatModal(null)}
          onApply={(format) => {
            applyFormatToSelection(format);
            setFormatModal(null);
          }}
        />
      )}

      {/* Find & Replace */}
      {findReplace.isOpen && (
        <FindReplaceModal
          query={findReplace.query}
          replacement={findReplace.replacement}
          matchCase={findReplace.matchCase}
          wholeCell={findReplace.wholeCell}
          searchFormulas={findReplace.searchFormulas}
          matchCount={findReplaceMatches.length}
          activeMatchIndex={activeMatchIndex}
          onChange={handleFindReplaceChange}
          onFindNext={handleFindNext}
          onFindPrev={handleFindPrev}
          onReplace={handleReplace}
          onReplaceAll={handleReplaceAll}
          onClose={handleCloseFindReplace}
        />
      )}
    </div>
  );
});

export default WorkbookCanvas;

