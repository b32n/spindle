import React, { useState, useEffect, useRef, useCallback, memo } from 'react';
import { EditorView } from 'prosemirror-view';
import { docsSchema, createCommands } from '@b32nio/spindle-docs-core';

interface TableSizePickerProps {
  /** ProseMirror editor view */
  editorView: EditorView | null;
  /** Whether the picker is open */
  isOpen: boolean;
  /** Callback when picker is closed */
  onClose: () => void;
  /** Position to anchor the picker */
  anchorPosition?: { top: number; left: number };
}

const MAX_ROWS = 10;
const MAX_COLS = 10;
const DEFAULT_PREVIEW_ROWS = 5;
const DEFAULT_PREVIEW_COLS = 5;

/**
 * Grid-based table size picker (similar to Google Docs / Word)
 */
export const TableSizePicker = memo(function TableSizePicker({
  editorView,
  isOpen,
  onClose,
  anchorPosition,
}: TableSizePickerProps) {
  const [hoverRow, setHoverRow] = useState(0);
  const [hoverCol, setHoverCol] = useState(0);
  const [visibleRows, setVisibleRows] = useState(DEFAULT_PREVIEW_ROWS);
  const [visibleCols, setVisibleCols] = useState(DEFAULT_PREVIEW_COLS);
  const pickerRef = useRef<HTMLDivElement>(null);
  
  const commands = createCommands(docsSchema);
  
  // Reset state when picker opens
  useEffect(() => {
    if (isOpen) {
      setHoverRow(0);
      setHoverCol(0);
      setVisibleRows(DEFAULT_PREVIEW_ROWS);
      setVisibleCols(DEFAULT_PREVIEW_COLS);
    }
  }, [isOpen]);
  
  // Handle click outside to close
  useEffect(() => {
    if (!isOpen) return;
    
    const handleClickOutside = (e: MouseEvent) => {
      if (pickerRef.current && !pickerRef.current.contains(e.target as Node)) {
        onClose();
      }
    };
    
    // Delay adding listener to prevent immediate close
    const timer = setTimeout(() => {
      document.addEventListener('mousedown', handleClickOutside);
    }, 0);
    
    return () => {
      clearTimeout(timer);
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isOpen, onClose]);
  
  // Handle escape key
  useEffect(() => {
    if (!isOpen) return;
    
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        onClose();
      }
    };
    
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose]);
  
  // Handle cell hover - expand grid if needed
  const handleCellHover = useCallback((row: number, col: number) => {
    setHoverRow(row);
    setHoverCol(col);
    
    // Expand visible grid when hovering near edges
    if (row >= visibleRows - 1 && visibleRows < MAX_ROWS) {
      setVisibleRows(Math.min(visibleRows + 1, MAX_ROWS));
    }
    if (col >= visibleCols - 1 && visibleCols < MAX_COLS) {
      setVisibleCols(Math.min(visibleCols + 1, MAX_COLS));
    }
  }, [visibleRows, visibleCols]);
  
  // Handle cell click - insert table
  const handleCellClick = useCallback((rows: number, cols: number) => {
    if (!editorView) return;
    
    const { state, dispatch } = editorView;
    commands.insertTable(rows, cols)(state, dispatch, editorView);
    
    editorView.focus();
    onClose();
  }, [editorView, commands, onClose]);
  
  if (!isOpen) return null;
  
  // Calculate position
  const style: React.CSSProperties = {
    position: 'absolute',
    zIndex: 10000,
    ...(anchorPosition
      ? { top: anchorPosition.top, left: anchorPosition.left }
      : { top: '100%', left: 0 }),
  };
  
  return (
    <div
      ref={pickerRef}
      style={{
        ...style,
        backgroundColor: '#ffffff',
        borderRadius: '8px',
        boxShadow: '0 4px 16px rgba(0, 0, 0, 0.15)',
        padding: '12px',
        marginTop: '4px',
      }}
    >
      {/* Size label */}
      <div
        style={{
          textAlign: 'center',
          marginBottom: '8px',
          fontSize: '13px',
          color: '#5f6368',
          fontWeight: 500,
        }}
      >
        {hoverRow > 0 && hoverCol > 0
          ? `${hoverRow} × ${hoverCol}`
          : 'Select table size'}
      </div>
      
      {/* Grid */}
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: `repeat(${visibleCols}, 20px)`,
          gridTemplateRows: `repeat(${visibleRows}, 20px)`,
          gap: '2px',
        }}
        onMouseLeave={() => {
          setHoverRow(0);
          setHoverCol(0);
        }}
      >
        {Array.from({ length: visibleRows }).map((_, rowIndex) =>
          Array.from({ length: visibleCols }).map((_, colIndex) => {
            const row = rowIndex + 1;
            const col = colIndex + 1;
            const isHighlighted = row <= hoverRow && col <= hoverCol;
            
            return (
              <div
                key={`${row}-${col}`}
                onMouseEnter={() => handleCellHover(row, col)}
                onClick={() => handleCellClick(row, col)}
                style={{
                  width: '18px',
                  height: '18px',
                  border: '1px solid',
                  borderColor: isHighlighted ? '#1a73e8' : '#dadce0',
                  backgroundColor: isHighlighted ? '#e8f0fe' : '#ffffff',
                  borderRadius: '2px',
                  cursor: 'pointer',
                  transition: 'all 0.1s ease',
                }}
              />
            );
          })
        )}
      </div>
      
      {/* Insert button for custom sizes (optional enhancement) */}
      <div
        style={{
          marginTop: '8px',
          paddingTop: '8px',
          borderTop: '1px solid #e8eaed',
        }}
      >
        <button
          onClick={() => {
            // For now, just insert a 3x3 table if no selection
            if (hoverRow > 0 && hoverCol > 0) {
              handleCellClick(hoverRow, hoverCol);
            } else {
              handleCellClick(3, 3);
            }
          }}
          style={{
            width: '100%',
            padding: '8px',
            fontSize: '13px',
            color: '#1a73e8',
            backgroundColor: 'transparent',
            border: '1px solid #dadce0',
            borderRadius: '4px',
            cursor: 'pointer',
            fontWeight: 500,
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.backgroundColor = '#f8f9fa';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.backgroundColor = 'transparent';
          }}
        >
          {hoverRow > 0 && hoverCol > 0
            ? `Insert ${hoverRow}×${hoverCol} table`
            : 'Insert 3×3 table'}
        </button>
      </div>
    </div>
  );
});

