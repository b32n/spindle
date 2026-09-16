import React, { memo, useState, useEffect, useRef } from 'react';
import type { ColumnFilter, FilterCriteria, Sheet } from '@b32nio/spindle-sheets-core';
import { FilterManager } from '@b32nio/spindle-sheets-core';

interface FilterModalProps {
  isOpen: boolean;
  sheet: Sheet;
  column: number;
  existingFilter?: ColumnFilter;
  onClose: () => void;
  onApply: (filter: ColumnFilter) => void;
  onClear: () => void;
}

type FilterType = 'text' | 'number' | 'date';

interface FilterOption {
  value: string;
  label: string;
}

const TEXT_FILTER_OPTIONS: FilterOption[] = [
  { value: 'equals', label: 'Equals' },
  { value: 'notEquals', label: 'Does not equal' },
  { value: 'contains', label: 'Contains' },
  { value: 'notContains', label: 'Does not contain' },
  { value: 'startsWith', label: 'Starts with' },
  { value: 'endsWith', label: 'Ends with' },
  { value: 'custom', label: 'Custom...' },
];

const NUMBER_FILTER_OPTIONS: FilterOption[] = [
  { value: 'equals', label: 'Equals' },
  { value: 'notEquals', label: 'Does not equal' },
  { value: 'greaterThan', label: 'Greater than' },
  { value: 'lessThan', label: 'Less than' },
  { value: 'greaterThanOrEqual', label: 'Greater than or equal to' },
  { value: 'lessThanOrEqual', label: 'Less than or equal to' },
  { value: 'between', label: 'Between' },
  { value: 'custom', label: 'Custom...' },
];

const DATE_FILTER_OPTIONS: FilterOption[] = [
  { value: 'equals', label: 'Equals' },
  { value: 'notEquals', label: 'Does not equal' },
  { value: 'greaterThan', label: 'After' },
  { value: 'lessThan', label: 'Before' },
  { value: 'between', label: 'Between' },
  { value: 'custom', label: 'Custom...' },
];

export const FilterModal = memo(function FilterModal({
  isOpen,
  sheet,
  column,
  existingFilter,
  onClose,
  onApply,
  onClear,
}: FilterModalProps) {
  const [filterType, setFilterType] = useState<FilterType>('text');
  const [criteriaType, setCriteriaType] = useState<string>('equals');
  const [textValue, setTextValue] = useState('');
  const [numberValue, setNumberValue] = useState('');
  const [minValue, setMinValue] = useState('');
  const [maxValue, setMaxValue] = useState('');
  const [selectedValues, setSelectedValues] = useState<Set<string | number | boolean>>(new Set());
  const [uniqueValues, setUniqueValues] = useState<Set<string | number | boolean>>(new Set());

  const firstInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isOpen) {
      // Reset form state
      const detectedType = FilterManager.detectColumnType(sheet, column);
      setFilterType(detectedType);

      // Load unique values for the column
      const values = FilterManager.getUniqueColumnValues(sheet, column);
      setUniqueValues(values);

      // Load existing filter if any
      if (existingFilter) {
        setFilterType(existingFilter.type);
        const criteria = existingFilter.criteria;

        switch (criteria.type) {
          case 'equals':
          case 'notEquals':
          case 'contains':
          case 'notContains':
          case 'startsWith':
          case 'endsWith':
            setCriteriaType(criteria.type);
            setTextValue(String(criteria.value));
            break;
          case 'greaterThan':
          case 'lessThan':
          case 'greaterThanOrEqual':
          case 'lessThanOrEqual':
            setCriteriaType(criteria.type);
            setNumberValue(String(criteria.value));
            break;
          case 'between':
            setCriteriaType(criteria.type);
            setMinValue(String(criteria.min));
            setMaxValue(String(criteria.max));
            break;
          case 'custom':
            setCriteriaType(criteria.type);
            setSelectedValues(new Set(criteria.values));
            break;
        }
      } else {
        // Reset to defaults
        setCriteriaType('equals');
        setTextValue('');
        setNumberValue('');
        setMinValue('');
        setMaxValue('');
        setSelectedValues(new Set());
      }

      // Focus first input after modal opens
      setTimeout(() => {
        firstInputRef.current?.focus();
      }, 0);
    }
  }, [isOpen, sheet, column, existingFilter]);

  const getFilterOptions = (): FilterOption[] => {
    switch (filterType) {
      case 'number':
        return NUMBER_FILTER_OPTIONS;
      case 'date':
        return DATE_FILTER_OPTIONS;
      default:
        return TEXT_FILTER_OPTIONS;
    }
  };

  const handleApply = () => {
    let criteria: FilterCriteria;

    switch (criteriaType) {
      case 'equals':
      case 'notEquals':
        criteria = {
          type: criteriaType as 'equals' | 'notEquals',
          value: filterType === 'number' ? Number(numberValue) : textValue,
        };
        break;
      case 'contains':
      case 'notContains':
      case 'startsWith':
      case 'endsWith':
        criteria = {
          type: criteriaType as 'contains' | 'notContains' | 'startsWith' | 'endsWith',
          value: textValue,
        };
        break;
      case 'greaterThan':
      case 'lessThan':
      case 'greaterThanOrEqual':
      case 'lessThanOrEqual':
        criteria = {
          type: criteriaType as 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual',
          value: Number(numberValue),
        };
        break;
      case 'between':
        criteria = {
          type: 'between',
          min: Number(minValue),
          max: Number(maxValue),
        };
        break;
      case 'custom':
        criteria = {
          type: 'custom',
          values: new Set(Array.from(selectedValues).filter((value) => typeof value === 'string' || typeof value === 'number')),
        };
        break;
      default:
        return; // Invalid criteria
    }

    const filter: ColumnFilter = {
      column,
      type: filterType,
      criteria,
    };

    onApply(filter);
    onClose();
  };

  const handleValueToggle = (value: string | number | boolean) => {
    const newSelected = new Set(selectedValues);
    if (newSelected.has(value)) {
      newSelected.delete(value);
    } else {
      newSelected.add(value);
    }
    setSelectedValues(newSelected);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      onClose();
    } else if (e.key === 'Enter' && canApply()) {
      handleApply();
    }
  };

  const canApply = (): boolean => {
    switch (criteriaType) {
      case 'equals':
      case 'notEquals':
        return filterType === 'number' ? numberValue.trim() !== '' : textValue.trim() !== '';
      case 'contains':
      case 'notContains':
      case 'startsWith':
      case 'endsWith':
        return textValue.trim() !== '';
      case 'greaterThan':
      case 'lessThan':
      case 'greaterThanOrEqual':
      case 'lessThanOrEqual':
        return numberValue.trim() !== '' && !isNaN(Number(numberValue));
      case 'between':
        return minValue.trim() !== '' && maxValue.trim() !== '' &&
               !isNaN(Number(minValue)) && !isNaN(Number(maxValue));
      case 'custom':
        return selectedValues.size > 0;
      default:
        return false;
    }
  };

  const renderCriteriaInput = () => {
    switch (criteriaType) {
      case 'equals':
      case 'notEquals':
        return filterType === 'number' ? (
          <input
            ref={firstInputRef}
            type="number"
            value={numberValue}
            onChange={(e) => setNumberValue(e.target.value)}
            placeholder="Enter value"
            style={inputStyle}
            onKeyDown={handleKeyDown}
          />
        ) : (
          <input
            ref={firstInputRef}
            type="text"
            value={textValue}
            onChange={(e) => setTextValue(e.target.value)}
            placeholder="Enter text"
            style={inputStyle}
            onKeyDown={handleKeyDown}
          />
        );

      case 'contains':
      case 'notContains':
      case 'startsWith':
      case 'endsWith':
        return (
          <input
            ref={firstInputRef}
            type="text"
            value={textValue}
            onChange={(e) => setTextValue(e.target.value)}
            placeholder="Enter text"
            style={inputStyle}
            onKeyDown={handleKeyDown}
          />
        );

      case 'greaterThan':
      case 'lessThan':
      case 'greaterThanOrEqual':
      case 'lessThanOrEqual':
        return (
          <input
            ref={firstInputRef}
            type="number"
            value={numberValue}
            onChange={(e) => setNumberValue(e.target.value)}
            placeholder="Enter number"
            style={inputStyle}
            onKeyDown={handleKeyDown}
          />
        );

      case 'between':
        return (
          <div style={{ display: 'flex', gap: '8px' }}>
            <input
              ref={firstInputRef}
              type="number"
              value={minValue}
              onChange={(e) => setMinValue(e.target.value)}
              placeholder="Min"
              style={{ ...inputStyle, flex: 1 }}
              onKeyDown={handleKeyDown}
            />
            <input
              type="number"
              value={maxValue}
              onChange={(e) => setMaxValue(e.target.value)}
              placeholder="Max"
              style={{ ...inputStyle, flex: 1 }}
              onKeyDown={handleKeyDown}
            />
          </div>
        );

      case 'custom':
        return (
          <div style={{ maxHeight: '200px', overflowY: 'auto', border: '1px solid #e8eaed', borderRadius: '4px' }}>
            {Array.from(uniqueValues).map((value) => (
              <label key={String(value)} style={checkboxLabelStyle}>
                <input
                  type="checkbox"
                  checked={selectedValues.has(value)}
                  onChange={() => handleValueToggle(value)}
                  style={checkboxStyle}
                />
                <span style={checkboxTextStyle}>{String(value)}</span>
              </label>
            ))}
          </div>
        );

      default:
        return null;
    }
  };

  if (!isOpen) return null;

  return (
    <div
      style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: 'rgba(0, 0, 0, 0.5)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 10000,
      }}
      onClick={onClose}
      onKeyDown={handleKeyDown}
    >
      <div
        style={{
          backgroundColor: '#ffffff',
          borderRadius: '8px',
          boxShadow: '0 4px 16px rgba(0, 0, 0, 0.2)',
          padding: '24px',
          minWidth: '400px',
          maxWidth: '90vw',
          maxHeight: '80vh',
          overflow: 'auto',
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <h2
          style={{
            margin: '0 0 16px 0',
            fontSize: '18px',
            fontWeight: 500,
            color: '#202124',
          }}
        >
          Filter Column {column + 1}
        </h2>

        <div style={{ marginBottom: '16px' }}>
          <label style={labelStyle}>Filter by:</label>
          <select
            value={criteriaType}
            onChange={(e) => setCriteriaType(e.target.value)}
            style={selectStyle}
          >
            {getFilterOptions().map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>

        <div style={{ marginBottom: '24px' }}>
          {renderCriteriaInput()}
        </div>

        <div
          style={{
            display: 'flex',
            gap: '8px',
            justifyContent: 'flex-end',
          }}
        >
          <button
            type="button"
            onClick={onClear}
            style={secondaryButtonStyle}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#f1f3f4';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = '#ffffff';
            }}
          >
            Clear Filter
          </button>
          <button
            type="button"
            onClick={onClose}
            style={secondaryButtonStyle}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#f1f3f4';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = '#ffffff';
            }}
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={handleApply}
            disabled={!canApply()}
            style={{
              ...primaryButtonStyle,
              opacity: canApply() ? 1 : 0.5,
              cursor: canApply() ? 'pointer' : 'not-allowed',
            }}
            onMouseEnter={(e) => {
              if (canApply()) {
                e.currentTarget.style.backgroundColor = '#1557b0';
              }
            }}
            onMouseLeave={(e) => {
              if (canApply()) {
                e.currentTarget.style.backgroundColor = '#1a73e8';
              }
            }}
          >
            Apply Filter
          </button>
        </div>
      </div>
    </div>
  );
});

// Styles
const labelStyle: React.CSSProperties = {
  display: 'block',
  marginBottom: '4px',
  fontSize: '13px',
  color: '#5f6368',
  fontWeight: 500,
};

const inputStyle: React.CSSProperties = {
  width: '100%',
  padding: '8px 12px',
  border: '1px solid #e8eaed',
  borderRadius: '4px',
  fontSize: '14px',
  boxSizing: 'border-box',
};

const selectStyle: React.CSSProperties = {
  width: '100%',
  padding: '8px 12px',
  border: '1px solid #e8eaed',
  borderRadius: '4px',
  fontSize: '14px',
  backgroundColor: '#ffffff',
  boxSizing: 'border-box',
};

const checkboxLabelStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  padding: '4px 8px',
  cursor: 'pointer',
  borderBottom: '1px solid #f1f3f4',
};

const checkboxStyle: React.CSSProperties = {
  marginRight: '8px',
};

const checkboxTextStyle: React.CSSProperties = {
  fontSize: '13px',
  color: '#202124',
};

const secondaryButtonStyle: React.CSSProperties = {
  padding: '8px 16px',
  border: '1px solid #e8eaed',
  borderRadius: '4px',
  backgroundColor: '#ffffff',
  color: '#5f6368',
  cursor: 'pointer',
  fontSize: '13px',
  fontWeight: 500,
};

const primaryButtonStyle: React.CSSProperties = {
  padding: '8px 16px',
  border: 'none',
  borderRadius: '4px',
  backgroundColor: '#1a73e8',
  color: '#ffffff',
  cursor: 'pointer',
  fontSize: '13px',
  fontWeight: 500,
};
