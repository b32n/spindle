import React, { memo, useRef, useEffect, useCallback, useState, useImperativeHandle, forwardRef } from 'react';
import type { CellFormat } from '@b32nio/spindle-sheets-core';
import { excelDateToJS, dateStringToExcelSerial, dateTimeStringToExcelSerial } from '@b32nio/spindle-sheets-core';

export interface EditOverlayRef {
  insertAtCursor: (text: string, replaceExisting?: boolean) => void;
  getCursorPosition: () => number;
  setSelection: (start: number, end?: number) => void;
  focus: () => void;
}

export interface EditOverlayProps {
  value: string;
  onChange: (value: string) => void;
  onCommit: (value: string, moveToNextCell?: boolean) => void;
  onCancel: () => void;
  x: number;
  y: number;
  width: number;
  height: number;
  minWidth?: number;
  fontSize?: number;
  fontFamily?: string;
  onFocus?: () => void;
  onBlurCapture?: () => void;
  isEditingFormula?: boolean;
  cellFormat?: CellFormat;
  /** Intercepts a keydown before the editor acts on it; return true if consumed. */
  interceptKeyDown?: (e: React.KeyboardEvent<HTMLInputElement>) => boolean;
  /** Reports the caret offset whenever it moves. */
  onSelectionChange?: (caret: number) => void;
}

export const EditOverlay = memo(forwardRef<EditOverlayRef, EditOverlayProps>(function EditOverlay({
  value,
  onChange,
  onCommit,
  onCancel,
  x,
  y,
  width,
  height,
  minWidth = 100,
  fontSize = 11,
  fontFamily = 'Arial',
  onFocus,
  onBlurCapture,
  isEditingFormula = false,
  cellFormat,
  interceptKeyDown,
  onSelectionChange,
}, ref) {
  const inputRef = useRef<HTMLInputElement>(null);
  const dateInputRef = useRef<HTMLInputElement>(null);
  const timeInputRef = useRef<HTMLInputElement>(null);
  const datetimeInputRef = useRef<HTMLInputElement>(null);
  const [inputWidth, setInputWidth] = useState(Math.max(width, minWidth));
  const measureRef = useRef<HTMLSpanElement>(null);

  // Determine if we should show date/time pickers
  const showDatePicker = cellFormat?.type === 'date';
  const showTimePicker = cellFormat?.type === 'time';
  const showDateTimePicker = cellFormat?.type === 'datetime';

  // Convert Excel date serial to HTML date/time input format
  const getPickerValue = useCallback(() => {
    if (!cellFormat) return '';

    // Try to parse the current value as a number (Excel date serial)
    const numValue = parseFloat(value);
    if (isNaN(numValue)) return '';

    const jsDate = excelDateToJS(numValue);

    if (showDatePicker) {
      return jsDate.toISOString().split('T')[0]; // YYYY-MM-DD
    } else if (showTimePicker) {
      return jsDate.toTimeString().slice(0, 5); // HH:MM
    } else if (showDateTimePicker) {
      return jsDate.toISOString().slice(0, 16); // YYYY-MM-DDTHH:MM
    }

    return '';
  }, [value, cellFormat, showDatePicker, showTimePicker, showDateTimePicker]);

  // Convert an HTML date/time input value back to an Excel serial number.
  // getPickerValue / excelDateToJS work in UTC (serial 0 = 1899-12-30 UTC), so
  // this inverse must build the serial in UTC too — a local-time round trip
  // drifts by a whole day in negative-UTC timezones (the picker showed the right
  // date but committing it decremented the stored serial).
  const convertPickerToSerial = useCallback((pickerValue: string) => {
    if (!pickerValue) return '';
    const DAY_MS = 24 * 60 * 60 * 1000;

    if (showDatePicker) {
      return String(dateStringToExcelSerial(pickerValue));
    } else if (showTimePicker) {
      // For time-only, assume today's date (behavior unchanged).
      const today = new Date().toISOString().split('T')[0];
      const jsDate = new Date(today + 'T' + pickerValue + ':00');
      return String(Math.floor((jsDate.getTime() - new Date(1899, 11, 30).getTime()) / DAY_MS));
    } else if (showDateTimePicker) {
      return String(dateTimeStringToExcelSerial(pickerValue));
    }
    return pickerValue;
  }, [showDatePicker, showTimePicker, showDateTimePicker]);
  
  // Track the last inserted reference position for replacement
  const lastInsertRef = useRef<{ start: number; end: number } | null>(null);
  
  // Expose methods via ref
  useImperativeHandle(ref, () => ({
    insertAtCursor: (text: string, replaceExisting = false) => {
      if (!inputRef.current) return;
      
      const input = inputRef.current;
      const currentValue = input.value;
      
      let start: number;
      let end: number;
      
      if (replaceExisting && lastInsertRef.current) {
        // Replace the previously inserted reference
        start = lastInsertRef.current.start;
        end = lastInsertRef.current.end;
      } else {
        // Insert at cursor position
        start = input.selectionStart ?? currentValue.length;
        end = input.selectionEnd ?? start;
      }
      
      // Create new value with inserted text
      const newValue = currentValue.slice(0, start) + text + currentValue.slice(end);
      
      // Update the value
      onChange(newValue);
      
      // Track the inserted reference position
      lastInsertRef.current = { start, end: start + text.length };
      
      // Update cursor position after React updates the input
      requestAnimationFrame(() => {
        if (inputRef.current) {
          const newCursorPos = start + text.length;
          inputRef.current.focus();
          inputRef.current.setSelectionRange(newCursorPos, newCursorPos);
        }
      });
    },
    getCursorPosition: () => {
      return inputRef.current?.selectionStart ?? 0;
    },
    setSelection: (start: number, end: number = start) => {
      // Deferred so it lands after React commits any pending value change.
      requestAnimationFrame(() => {
        const input = inputRef.current;
        if (input) {
          input.focus();
          input.setSelectionRange(start, end);
        }
      });
    },
    focus: () => {
      const inputToFocus = showDateTimePicker ? datetimeInputRef.current :
                          showDatePicker ? dateInputRef.current :
                          showTimePicker ? timeInputRef.current :
                          inputRef.current;
      inputToFocus?.focus();
    },
  }), [onChange]);
  
  // Focus appropriate input on mount
  useEffect(() => {
    const inputToFocus = showDateTimePicker ? datetimeInputRef.current :
                        showDatePicker ? dateInputRef.current :
                        showTimePicker ? timeInputRef.current :
                        inputRef.current;

    if (inputToFocus) {
      inputToFocus.focus();
      // For text input, place cursor at end
      if (inputToFocus === inputRef.current && 'setSelectionRange' in inputToFocus) {
        const len = value.length;
        inputToFocus.setSelectionRange(len, len);
      }
    }
  }, [showDatePicker, showTimePicker, showDateTimePicker]);
  
  // Clear last insert ref when value changes externally (not via insertAtCursor)
  useEffect(() => {
    // If the value doesn't match what we expect from last insert, clear it
    if (lastInsertRef.current) {
      const expectedEnd = lastInsertRef.current.end;
      if (expectedEnd > value.length) {
        lastInsertRef.current = null;
      }
    }
  }, [value]);
  
  // Auto-expand input width based on content
  useEffect(() => {
    if (measureRef.current) {
      const measuredWidth = measureRef.current.offsetWidth;
      setInputWidth(Math.max(width, minWidth, measuredWidth + 16)); // 16px padding
    }
  }, [value, width, minWidth]);
  
  // Handle key events
  // Use e.currentTarget.value to get the most current value from the DOM
  // rather than relying on the prop value which may be stale due to React's batching
  const handleKeyDown = useCallback((e: React.KeyboardEvent<HTMLInputElement>) => {
    // The formula-assist popup gets first refusal on navigation keys.
    if (interceptKeyDown?.(e)) return;
    if (e.key === 'Enter') {
      e.preventDefault();
      lastInsertRef.current = null;
      onCommit(e.currentTarget.value, true);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      lastInsertRef.current = null;
      onCancel();
    } else if (e.key === 'Tab') {
      e.preventDefault();
      lastInsertRef.current = null;
      onCommit(e.currentTarget.value);
      // Let parent handle tab navigation
    }
  }, [onCommit, onCancel, interceptKeyDown]);

  // Handle input change
  const handleChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    // Clear last insert ref when user types manually
    lastInsertRef.current = null;
    onChange(e.target.value);
    onSelectionChange?.(e.target.selectionStart ?? e.target.value.length);
  }, [onChange, onSelectionChange]);

  // Report caret moves (arrows, clicks) so formula-assist tracks the position.
  const reportSelection = useCallback((e: React.SyntheticEvent<HTMLInputElement>) => {
    onSelectionChange?.(e.currentTarget.selectionStart ?? 0);
  }, [onSelectionChange]);

  // Handle date/time picker changes
  const handlePickerChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const serialValue = convertPickerToSerial(e.target.value);
    onChange(serialValue);
  }, [convertPickerToSerial, onChange]);
  
  // Handle blur (commit on focus loss)
  const handleBlur = useCallback((e: React.FocusEvent<HTMLInputElement>) => {
    const relatedTarget = e.relatedTarget as HTMLElement;
    
    // Check if the focus is going to a canvas (likely clicking on a cell for formula reference)
    if (relatedTarget?.tagName === 'CANVAS') {
      // Don't commit - we're selecting cells for a formula
      onBlurCapture?.();
      return;
    }
    
    // If editing a formula, don't commit on blur - user might be switching sheets
    // for cross-sheet references. Only commit via Enter key or explicit action.
    if (isEditingFormula) {
      onBlurCapture?.();
      return;
    }
    
    lastInsertRef.current = null;
    onCommit(e.currentTarget.value);
  }, [onCommit, onBlurCapture, isEditingFormula]);
  
  return (
    <>
      {/* Hidden element for measuring text width */}
      <span
        ref={measureRef}
        style={{
          position: 'absolute',
          visibility: 'hidden',
          whiteSpace: 'pre',
          fontSize,
          fontFamily,
        }}
      >
        {value}
      </span>

      {/* Date/Time pickers for formatted cells */}
      {showDatePicker && (
        <input
          ref={dateInputRef}
          type="date"
          value={getPickerValue()}
          onChange={handlePickerChange}
          onKeyDown={handleKeyDown}
          onBlur={handleBlur}
          onFocus={onFocus}
          style={{
            position: 'absolute',
            left: x,
            top: y,
            width: Math.max(width, 120), // Minimum width for date picker
            height: height - 2,
            padding: '0 4px',
            margin: 0,
            border: '1px solid #1a73e8',
            borderRadius: '4px',
            outline: 'none',
            fontSize,
            fontFamily,
            backgroundColor: '#ffffff',
            boxSizing: 'border-box',
            boxShadow: '0 1px 3px rgba(0, 0, 0, 0.12)',
            transition: 'box-shadow 0.15s ease',
            zIndex: 100,
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.boxShadow = '0 2px 6px rgba(26, 115, 232, 0.2)';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.boxShadow = '0 1px 3px rgba(0, 0, 0, 0.12)';
          }}
        />
      )}

      {showTimePicker && (
        <input
          ref={timeInputRef}
          type="time"
          value={getPickerValue()}
          onChange={handlePickerChange}
          onKeyDown={handleKeyDown}
          onBlur={handleBlur}
          onFocus={onFocus}
          style={{
            position: 'absolute',
            left: x,
            top: y,
            width: Math.max(width, 100), // Minimum width for time picker
            height: height - 2,
            padding: '0 4px',
            margin: 0,
            border: '1px solid #1a73e8',
            borderRadius: '4px',
            outline: 'none',
            fontSize,
            fontFamily,
            backgroundColor: '#ffffff',
            boxSizing: 'border-box',
            boxShadow: '0 1px 3px rgba(0, 0, 0, 0.12)',
            transition: 'box-shadow 0.15s ease',
            zIndex: 100,
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.boxShadow = '0 2px 6px rgba(26, 115, 232, 0.2)';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.boxShadow = '0 1px 3px rgba(0, 0, 0, 0.12)';
          }}
        />
      )}

      {showDateTimePicker && (
        <input
          ref={datetimeInputRef}
          type="datetime-local"
          value={getPickerValue()}
          onChange={handlePickerChange}
          onKeyDown={handleKeyDown}
          onBlur={handleBlur}
          onFocus={onFocus}
          style={{
            position: 'absolute',
            left: x,
            top: y,
            width: Math.max(width, 180), // Minimum width for datetime picker
            height: height - 2,
            padding: '0 4px',
            margin: 0,
            border: '1px solid #1a73e8',
            borderRadius: '4px',
            outline: 'none',
            fontSize,
            fontFamily,
            backgroundColor: '#ffffff',
            boxSizing: 'border-box',
            boxShadow: '0 1px 3px rgba(0, 0, 0, 0.12)',
            transition: 'box-shadow 0.15s ease',
            zIndex: 100,
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.boxShadow = '0 2px 6px rgba(26, 115, 232, 0.2)';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.boxShadow = '0 1px 3px rgba(0, 0, 0, 0.12)';
          }}
        />
      )}

      {/* Edit input - shown for non-date/time cells, or alongside date/time pickers for manual editing */}
      {(!showDatePicker && !showTimePicker && !showDateTimePicker) && (
        <input
          ref={inputRef}
          type="text"
          value={value}
          onChange={handleChange}
          onKeyDown={handleKeyDown}
          onKeyUp={reportSelection}
          onMouseUp={reportSelection}
          onBlur={handleBlur}
          onFocus={onFocus}
          style={{
            position: 'absolute',
            left: x,
            top: y,
            width: inputWidth,
            height: height - 2, // Account for border
            minWidth,
            padding: '0 4px',
            margin: 0,
            border: '1px solid #1a73e8',
            borderRadius: '4px',
            outline: 'none',
            fontSize,
            fontFamily,
            lineHeight: `${height - 4}px`,
            backgroundColor: '#ffffff',
            boxSizing: 'border-box',
            boxShadow: '0 1px 3px rgba(0, 0, 0, 0.12)',
            transition: 'box-shadow 0.15s ease',
            zIndex: 100,
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.boxShadow = '0 2px 6px rgba(26, 115, 232, 0.2)';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.boxShadow = '0 1px 3px rgba(0, 0, 0, 0.12)';
          }}
        />
      )}

      {/* Text input alongside date/time pickers for manual editing */}
      {(showDatePicker || showTimePicker || showDateTimePicker) && (
        <input
          ref={inputRef}
          type="text"
          value={value}
          onChange={handleChange}
          onKeyDown={handleKeyDown}
          onKeyUp={reportSelection}
          onMouseUp={reportSelection}
          onBlur={handleBlur}
          onFocus={onFocus}
          placeholder="Manual entry..."
          style={{
            position: 'absolute',
            left: x + (showDatePicker ? 130 : showTimePicker ? 110 : 190) + 10,
            top: y,
            width: Math.max(width - (showDatePicker ? 130 : showTimePicker ? 110 : 190) - 20, 100),
            height: height - 2, // Account for border
            padding: '0 4px',
            margin: 0,
            border: '1px solid #1a73e8',
            borderRadius: '4px',
            outline: 'none',
            fontSize,
            fontFamily,
            lineHeight: `${height - 4}px`,
            backgroundColor: '#ffffff',
            boxSizing: 'border-box',
            boxShadow: '0 1px 3px rgba(0, 0, 0, 0.12)',
            transition: 'box-shadow 0.15s ease',
            zIndex: 100,
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.boxShadow = '0 2px 6px rgba(26, 115, 232, 0.2)';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.boxShadow = '0 1px 3px rgba(0, 0, 0, 0.12)';
          }}
        />
      )}
    </>
  );
}));

export default EditOverlay;

