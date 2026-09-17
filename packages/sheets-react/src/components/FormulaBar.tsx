import React, { memo, useCallback, useState, useEffect, useRef } from 'react';
import { useWorkbook } from '../context/WorkbookContext';
import { columnIndexToLabel, excelDateToJS, formatJSDate } from '@b32nio/spindle-sheets-core';
import type { Cell, WorkbookImpl } from '@b32nio/spindle-sheets-core';
import { FormulaAutocomplete } from './FormulaAutocomplete';
import { FormulaSignatureHint } from './FormulaSignatureHint';
import { useFormulaAssist } from '../hooks/useFormulaAssist';

interface FormulaBarProps {
  activeCell: { row: number; col: number } | null;
  onFormulaChange?: (formula: string) => void;
  /** Disables the input entirely, independent of the per-cell spill read-only state below. */
  readOnly?: boolean;
}

/**
 * The text shown in the formula bar for a cell. Date-formatted cells render
 * their value using the cell's date pattern (so the user sees "1/12/2020",
 * not the underlying serial number).
 */
function getFormulaBarText(cell: Cell | undefined, workbook: WorkbookImpl): string {
  if (!cell) return '';
  if (cell.formula) return cell.formula;
  if (cell.value === null || cell.value === undefined) return '';
  if (typeof cell.value === 'number' && cell.formatId) {
    const format = workbook.getFormatPool().get(cell.formatId);
    if (format && (format.type === 'date' || format.type === 'datetime')) {
      try {
        return formatJSDate(excelDateToJS(cell.value), format.dateFormat || 'MM/DD/YYYY');
      } catch {
        // fall through to raw value
      }
    }
  }
  return String(cell.value);
}

export const FormulaBar = memo(function FormulaBar({
  activeCell,
  onFormulaChange,
  readOnly: workbookReadOnly = false,
}: FormulaBarProps) {
  const { workbook } = useWorkbook();
  const [inputValue, setInputValue] = useState('');
  const [isEditing, setIsEditing] = useState(false);
  const [spillReadOnly, setSpillReadOnly] = useState(false);
  const [caret, setCaret] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);

  // Disabled either because the workbook itself is read-only, or because this
  // cell is a spilled (non-anchor) cell whose formula lives on the anchor.
  const readOnly = workbookReadOnly || spillReadOnly;

  // Update input value when active cell changes
  useEffect(() => {
    if (!activeCell) {
      setInputValue('');
      setIsEditing(false);
      return;
    }

    if (!isEditing) {
      const anchor = workbook.getSpillAnchor(undefined, activeCell.row, activeCell.col);
      const spilledNonAnchor =
        anchor !== undefined && (anchor.row !== activeCell.row || anchor.col !== activeCell.col);
      if (spilledNonAnchor) {
        // A spilled cell shows the anchor's formula, read-only.
        const anchorCell = workbook.getCell(undefined, anchor.row, anchor.col);
        setInputValue(anchorCell?.formula || '');
        setSpillReadOnly(true);
      } else {
        const cell = workbook.getCell(undefined, activeCell.row, activeCell.col);
        setInputValue(getFormulaBarText(cell, workbook));
        setSpillReadOnly(false);
      }
    }
  }, [activeCell, workbook, isEditing]);

  // Formula autocomplete + parameter help for the formula bar.
  const onFormulaAccept = useCallback((nextValue: string, nextCaret: number) => {
    setInputValue(nextValue);
    setIsEditing(true);
    setCaret(nextCaret);
    requestAnimationFrame(() => {
      const input = inputRef.current;
      if (input) {
        input.focus();
        input.setSelectionRange(nextCaret, nextCaret);
      }
    });
  }, []);
  const formulaAssist = useFormulaAssist({
    value: inputValue,
    caret,
    enabled: isEditing && inputValue.startsWith('=') && !readOnly,
    onAccept: onFormulaAccept,
  });

  const reportSelection = useCallback((e: React.SyntheticEvent<HTMLInputElement>) => {
    setCaret(e.currentTarget.selectionStart ?? 0);
  }, []);

  const handleInputChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    setInputValue(e.target.value);
    setIsEditing(true);
    setCaret(e.target.selectionStart ?? e.target.value.length);
  }, []);

  const handleInputFocus = useCallback(() => {
    setIsEditing(true);
  }, []);

  const handleInputBlur = useCallback(() => {
    setIsEditing(false);
    if (activeCell && onFormulaChange && !readOnly) {
      onFormulaChange(inputValue);
    }
  }, [activeCell, inputValue, onFormulaChange, readOnly]);

  const handleInputKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLInputElement>) => {
      // The formula-assist popup gets first refusal on navigation keys.
      if (formulaAssist.onKeyDown(e)) return;
      if (e.key === 'Enter') {
        e.preventDefault();
        inputRef.current?.blur();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        setIsEditing(false);
        if (activeCell) {
          const cell = workbook.getCell(undefined, activeCell.row, activeCell.col);
          setInputValue(getFormulaBarText(cell, workbook));
        }
      }
    },
    [activeCell, workbook, formulaAssist]
  );

  const cellReference = activeCell
    ? `${columnIndexToLabel(activeCell.col)}${activeCell.row + 1}`
    : '';

  return (
    <div
      className="formula-bar"
      style={{
        display: 'flex',
        alignItems: 'center',
        height: '32px',
        backgroundColor: 'transparent',
        padding: '0 14px',
        gap: '8px',
        fontFamily: '"Inter", "SF Pro Display", -apple-system, BlinkMacSystemFont, sans-serif',
      }}
    >
      <div
        className="cell-reference"
        style={{
          minWidth: '72px',
          fontSize: '12px',
          fontWeight: 600,
          color: '#475569',
          textAlign: 'center',
          padding: '5px 10px',
          backgroundColor: 'rgba(255, 255, 255, 0.75)',
          border: '1px solid rgba(15, 23, 42, 0.1)',
          borderRadius: '8px',
        }}
      >
        {cellReference}
      </div>
      <div
        className="formula-input-container"
        style={{
          flex: 1,
          position: 'relative',
          display: 'flex',
          alignItems: 'center',
          border: '1px solid rgba(15, 23, 42, 0.1)',
          borderRadius: '8px',
          padding: '0 10px',
          backgroundColor: 'rgba(255, 255, 255, 0.8)',
          transition: 'border-color 0.15s ease, box-shadow 0.15s ease',
        }}
      >
        <span
          style={{
            fontSize: '14px',
            color: '#94a3b8',
            marginRight: '6px',
            fontWeight: 500,
          }}
        >
          =
        </span>
        <input
          ref={inputRef}
          type="text"
          value={inputValue}
          readOnly={readOnly}
          onChange={handleInputChange}
          onFocus={(e) => {
            handleInputFocus();
            const parent = e.currentTarget.parentElement!;
            parent.style.borderColor = '#6366f1';
            parent.style.boxShadow = '0 0 0 3px rgba(99, 102, 241, 0.12)';
          }}
          onBlur={(e) => {
            handleInputBlur();
            const parent = e.currentTarget.parentElement!;
            parent.style.borderColor = 'rgba(15, 23, 42, 0.1)';
            parent.style.boxShadow = 'none';
          }}
          onKeyDown={handleInputKeyDown}
          onKeyUp={reportSelection}
          onMouseUp={reportSelection}
          placeholder="Enter formula or value"
          style={{
            flex: 1,
            border: 'none',
            outline: 'none',
            fontSize: '13px',
            padding: '4px 0',
            backgroundColor: 'transparent',
            color: readOnly ? '#94a3b8' : '#1e293b',
            fontStyle: readOnly ? 'italic' : 'normal',
            fontFamily: 'inherit',
          }}
        />
        {formulaAssist.mode === 'autocomplete' && (
          <FormulaAutocomplete
            suggestions={formulaAssist.suggestions}
            highlightedIndex={formulaAssist.highlightedIndex}
            top={36}
            left={0}
            onHover={formulaAssist.setHighlightedIndex}
            onPick={formulaAssist.accept}
          />
        )}
        {formulaAssist.mode === 'signature' && formulaAssist.signature && (
          <FormulaSignatureHint
            doc={formulaAssist.signature.doc}
            activeArg={formulaAssist.signature.activeArg}
            top={36}
            left={0}
          />
        )}
      </div>
    </div>
  );
});

