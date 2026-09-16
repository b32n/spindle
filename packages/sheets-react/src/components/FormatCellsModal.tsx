import React, { memo, useState, useEffect, useRef } from 'react';
import type { CellFormat, FormatType } from '@b32nio/spindle-sheets-core';
import { formatNumber, getDefaultFormatForType } from '@b32nio/spindle-sheets-core';

interface FormatCellsModalProps {
  isOpen: boolean;
  currentFormat?: CellFormat;
  sampleValue?: number;
  onClose: () => void;
  onApply: (format: CellFormat) => void;
}

type FormatCategory = 'number' | 'currency' | 'datetime' | 'custom';

const CURRENCY_OPTIONS = [
  { code: 'USD', symbol: '$', name: 'US Dollar' },
  { code: 'EUR', symbol: '€', name: 'Euro' },
  { code: 'GBP', symbol: '£', name: 'British Pound' },
  { code: 'JPY', symbol: '¥', name: 'Japanese Yen' },
  { code: 'CAD', symbol: 'C$', name: 'Canadian Dollar' },
  { code: 'AUD', symbol: 'A$', name: 'Australian Dollar' },
  { code: 'CHF', symbol: 'CHF', name: 'Swiss Franc' },
  { code: 'CNY', symbol: '¥', name: 'Chinese Yuan' },
  { code: 'INR', symbol: '₹', name: 'Indian Rupee' },
  { code: 'KRW', symbol: '₩', name: 'Korean Won' },
];

const DATE_FORMAT_OPTIONS = [
  { value: 'MM/DD/YYYY', label: 'MM/DD/YYYY (12/31/2023)' },
  { value: 'DD-MM-YYYY', label: 'DD-MM-YYYY (31-12-2023)' },
  { value: 'YYYY-MM-DD', label: 'YYYY-MM-DD (2023-12-31)' },
  { value: 'Month DD YYYY', label: 'Month DD YYYY (December 31 2023)' },
];

const TIME_FORMAT_OPTIONS = [
  { value: 'HH:mm:ss', label: 'HH:mm:ss (14:30:25)' },
  { value: 'h:mm AM/PM', label: 'h:mm AM/PM (2:30 PM)' },
  { value: 'HH:mm', label: 'HH:mm (14:30)' },
];

export const FormatCellsModal = memo(function FormatCellsModal({
  isOpen,
  currentFormat,
  sampleValue = 1234.56,
  onClose,
  onApply,
}: FormatCellsModalProps) {
  const [activeCategory, setActiveCategory] = useState<FormatCategory>('number');
  const [format, setFormat] = useState<CellFormat>(currentFormat || {});
  const [customPattern, setCustomPattern] = useState('');

  const firstInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isOpen) {
      // Initialize format state
      if (currentFormat) {
        setFormat({ ...currentFormat });
        if (currentFormat.type === 'custom' && currentFormat.pattern) {
          setCustomPattern(currentFormat.pattern);
        }
      } else {
        setFormat({ type: 'number' }); // Default to number format
        setCustomPattern('');
      }

      // Determine initial category based on current format
      if (currentFormat?.type) {
        if (['currency', 'accounting'].includes(currentFormat.type)) {
          setActiveCategory('currency');
        } else if (['date', 'time', 'datetime'].includes(currentFormat.type)) {
          setActiveCategory('datetime');
        } else if (currentFormat.type === 'custom') {
          setActiveCategory('custom');
        } else {
          setActiveCategory('number');
        }
      }

      // Focus first input
      setTimeout(() => {
        firstInputRef.current?.focus();
      }, 0);
    }
  }, [isOpen, currentFormat]);

  const handleApply = () => {
    let finalFormat = { ...format };

    // Handle custom pattern
    if (activeCategory === 'custom' && customPattern.trim()) {
      finalFormat = {
        ...finalFormat,
        type: 'custom',
        pattern: customPattern.trim(),
      };
    }

    onApply(finalFormat);
    onClose();
  };

  const handleFormatTypeChange = (type: FormatType) => {
    const defaultFormat = getDefaultFormatForType(type);
    setFormat(defaultFormat);
  };

  const handleCategoryChange = (category: FormatCategory) => {
    setActiveCategory(category);

    // When switching categories, ensure the format type is appropriate for the category
    const currentType = format.type;
    let newType: FormatType;

    switch (category) {
      case 'currency':
        newType = currentType === 'accounting' ? 'accounting' : 'currency';
        break;
      case 'datetime':
        if (currentType && ['date', 'time', 'datetime'].includes(currentType)) {
          newType = currentType;
        } else {
          newType = 'date';
        }
        break;
      case 'custom':
        newType = 'custom';
        break;
      case 'number':
      default:
        if (currentType && ['number', 'percentage', 'scientific', 'fraction'].includes(currentType)) {
          newType = currentType;
        } else {
          newType = 'number';
        }
        break;
    }

    // Only change the format type if it's different
    if (newType !== currentType) {
      const defaultFormat = getDefaultFormatForType(newType);
      setFormat(prev => ({ ...defaultFormat, ...prev, type: newType }));
    }
  };

  const updateFormat = (updates: Partial<CellFormat>) => {
    setFormat(prev => ({ ...prev, ...updates }));
  };

  const getSampleFormattedValue = (): string => {
    try {
      const testFormat = activeCategory === 'custom' && customPattern.trim()
        ? { ...format, type: 'custom' as const, pattern: customPattern.trim() }
        : format;

      return formatNumber(sampleValue, testFormat as CellFormat);
    } catch {
      return 'Invalid format';
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      onClose();
    } else if (e.key === 'Enter' && !e.shiftKey) {
      handleApply();
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
          minWidth: '500px',
          maxWidth: '90vw',
          maxHeight: '80vh',
          overflow: 'auto',
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <h2
          style={{
            margin: '0 0 20px 0',
            fontSize: '18px',
            fontWeight: 500,
            color: '#202124',
          }}
        >
          Format Cells
        </h2>

        {/* Category Tabs */}
        <div style={{ display: 'flex', marginBottom: '20px', borderBottom: '1px solid #e8eaed' }}>
          {[
            { key: 'number', label: 'Number' },
            { key: 'currency', label: 'Currency' },
            { key: 'datetime', label: 'Date/Time' },
            { key: 'custom', label: 'Custom' },
          ].map((category) => (
            <button
              key={category.key}
              onClick={() => handleCategoryChange(category.key as FormatCategory)}
              style={{
                padding: '8px 16px',
                border: 'none',
                backgroundColor: activeCategory === category.key ? '#f1f3f4' : 'transparent',
                borderBottom: activeCategory === category.key ? '2px solid #1a73e8' : 'none',
                color: activeCategory === category.key ? '#1a73e8' : '#5f6368',
                fontSize: '14px',
                fontWeight: 500,
                cursor: 'pointer',
              }}
            >
              {category.label}
            </button>
          ))}
        </div>

        {/* Format Preview */}
        <div style={{ marginBottom: '20px', padding: '12px', backgroundColor: '#f8f9fa', borderRadius: '4px' }}>
          <div style={{ fontSize: '12px', color: '#5f6368', marginBottom: '4px' }}>Sample:</div>
          <div style={{ fontSize: '16px', fontFamily: 'monospace', color: '#202124' }}>
            {getSampleFormattedValue()}
          </div>
        </div>

        {/* Category Content */}
        <div style={{ marginBottom: '24px' }}>
          {activeCategory === 'number' && (
            <NumberFormatOptions
              format={format}
              onChange={updateFormat}
              onTypeChange={handleFormatTypeChange}
            />
          )}
          {activeCategory === 'currency' && (
            <CurrencyFormatOptions
              format={format}
              onChange={updateFormat}
              onTypeChange={handleFormatTypeChange}
            />
          )}
          {activeCategory === 'datetime' && (
            <DateTimeFormatOptions
              format={format}
              onChange={updateFormat}
              onTypeChange={handleFormatTypeChange}
            />
          )}
          {activeCategory === 'custom' && (
            <CustomFormatOptions
              pattern={customPattern}
              onPatternChange={setCustomPattern}
            />
          )}
        </div>

        {/* Action Buttons */}
        <div
          style={{
            display: 'flex',
            gap: '8px',
            justifyContent: 'flex-end',
          }}
        >
          <button
            type="button"
            onClick={onClose}
            style={{
              padding: '8px 16px',
              border: '1px solid #e8eaed',
              borderRadius: '4px',
              backgroundColor: '#ffffff',
              color: '#5f6368',
              cursor: 'pointer',
              fontSize: '13px',
              fontWeight: 500,
            }}
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
            style={{
              padding: '8px 16px',
              border: 'none',
              borderRadius: '4px',
              backgroundColor: '#1a73e8',
              color: '#ffffff',
              cursor: 'pointer',
              fontSize: '13px',
              fontWeight: 500,
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#1557b0';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = '#1a73e8';
            }}
          >
            Apply
          </button>
        </div>
      </div>
    </div>
  );
});

// Number Format Options Component
const NumberFormatOptions = memo(function NumberFormatOptions({
  format,
  onChange,
  onTypeChange,
}: {
  format: CellFormat;
  onChange: (updates: Partial<CellFormat>) => void;
  onTypeChange: (type: FormatType) => void;
}) {
  return (
    <div>
      <div style={{ marginBottom: '16px' }}>
        <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
          Format Type
        </label>
        <select
          value={format.type || 'number'}
          onChange={(e) => onTypeChange(e.target.value as FormatType)}
          style={{
            width: '100%',
            padding: '8px 12px',
            border: '1px solid #e8eaed',
            borderRadius: '4px',
            fontSize: '14px',
          }}
        >
          <option value="number">Number</option>
          <option value="percentage">Percentage</option>
          <option value="scientific">Scientific</option>
          <option value="fraction">Fraction</option>
          <option value="text">Text</option>
        </select>
      </div>

      {format.type === 'number' && (
        <>
          <div style={{ marginBottom: '16px' }}>
            <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
              Decimal Places
            </label>
            <input
              type="number"
              min="0"
              max="30"
              value={format.decimalPlaces ?? 2}
              onChange={(e) => onChange({ decimalPlaces: parseInt(e.target.value) || 0 })}
              style={{
                width: '100%',
                padding: '8px 12px',
                border: '1px solid #e8eaed',
                borderRadius: '4px',
                fontSize: '14px',
              }}
            />
          </div>
          <div style={{ marginBottom: '16px' }}>
            <label style={{ display: 'flex', alignItems: 'center', fontSize: '14px', fontWeight: 500 }}>
              <input
                type="checkbox"
                checked={format.useThousandsSeparator ?? true}
                onChange={(e) => onChange({ useThousandsSeparator: e.target.checked })}
                style={{ marginRight: '8px' }}
              />
              Use thousands separator
            </label>
          </div>
          <div style={{ marginBottom: '16px' }}>
            <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
              Negative Numbers
            </label>
            <select
              value={format.negativeFormat || 'minus'}
              onChange={(e) => onChange({ negativeFormat: e.target.value as 'minus' | 'parentheses' | 'red' })}
              style={{
                width: '100%',
                padding: '8px 12px',
                border: '1px solid #e8eaed',
                borderRadius: '4px',
                fontSize: '14px',
              }}
            >
              <option value="minus">-123</option>
              <option value="parentheses">(123)</option>
              <option value="red">123 (red)</option>
            </select>
          </div>
        </>
      )}

      {format.type === 'percentage' && (
        <div style={{ marginBottom: '16px' }}>
          <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
            Decimal Places
          </label>
          <input
            type="number"
            min="0"
            max="30"
            value={format.decimalPlaces ?? 2}
            onChange={(e) => onChange({ decimalPlaces: parseInt(e.target.value) || 0 })}
            style={{
              width: '100%',
              padding: '8px 12px',
              border: '1px solid #e8eaed',
              borderRadius: '4px',
              fontSize: '14px',
            }}
          />
        </div>
      )}

      {format.type === 'scientific' && (
        <div style={{ marginBottom: '16px' }}>
          <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
            Decimal Places
          </label>
          <input
            type="number"
            min="0"
            max="30"
            value={format.decimalPlaces ?? 2}
            onChange={(e) => onChange({ decimalPlaces: parseInt(e.target.value) || 0 })}
            style={{
              width: '100%',
              padding: '8px 12px',
              border: '1px solid #e8eaed',
              borderRadius: '4px',
              fontSize: '14px',
            }}
          />
        </div>
      )}

      {format.type === 'fraction' && (
        <div style={{ marginBottom: '16px' }}>
          <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
            Fraction Type
          </label>
          <select
            value={format.fractionType || 'upToOne'}
            onChange={(e) => onChange({ fractionType: e.target.value as CellFormat['fractionType'] })}
            style={{
              width: '100%',
              padding: '8px 12px',
              border: '1px solid #e8eaed',
              borderRadius: '4px',
              fontSize: '14px',
            }}
          >
            <option value="upToOne">Up to one digit (1/2)</option>
            <option value="upToTwo">Up to two digits (1/10)</option>
            <option value="upToThree">Up to three digits (1/100)</option>
            <option value="asHalves">As halves (1/2)</option>
            <option value="asQuarters">As quarters (1/4)</option>
            <option value="asEighths">As eighths (1/8)</option>
            <option value="asSixteenths">As sixteenths (1/16)</option>
            <option value="asTenths">As tenths (1/10)</option>
            <option value="asHundredths">As hundredths (1/100)</option>
          </select>
        </div>
      )}
    </div>
  );
});

// Currency Format Options Component
const CurrencyFormatOptions = memo(function CurrencyFormatOptions({
  format,
  onChange,
  onTypeChange,
}: {
  format: CellFormat;
  onChange: (updates: Partial<CellFormat>) => void;
  onTypeChange: (type: FormatType) => void;
}) {
  return (
    <div>
      <div style={{ marginBottom: '16px' }}>
        <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
          Format Type
        </label>
        <select
          value={format.type || 'currency'}
          onChange={(e) => onTypeChange(e.target.value as FormatType)}
          style={{
            width: '100%',
            padding: '8px 12px',
            border: '1px solid #e8eaed',
            borderRadius: '4px',
            fontSize: '14px',
          }}
        >
          <option value="currency">Currency</option>
          <option value="accounting">Accounting</option>
        </select>
      </div>

      <div style={{ marginBottom: '16px' }}>
        <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
          Currency
        </label>
        <select
          value={format.currencyCode || 'USD'}
          onChange={(e) => onChange({ currencyCode: e.target.value })}
          style={{
            width: '100%',
            padding: '8px 12px',
            border: '1px solid #e8eaed',
            borderRadius: '4px',
            fontSize: '14px',
          }}
        >
          {CURRENCY_OPTIONS.map((currency) => (
            <option key={currency.code} value={currency.code}>
              {currency.symbol} {currency.name} ({currency.code})
            </option>
          ))}
        </select>
      </div>

      <div style={{ marginBottom: '16px' }}>
        <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
          Decimal Places
        </label>
        <input
          type="number"
          min="0"
          max="30"
          value={format.decimalPlaces ?? 2}
          onChange={(e) => onChange({ decimalPlaces: parseInt(e.target.value) || 0 })}
          style={{
            width: '100%',
            padding: '8px 12px',
            border: '1px solid #e8eaed',
            borderRadius: '4px',
            fontSize: '14px',
          }}
        />
      </div>

      <div style={{ marginBottom: '16px' }}>
        <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
          Negative Numbers
        </label>
        <select
          value={format.negativeFormat || 'minus'}
          onChange={(e) => onChange({ negativeFormat: e.target.value as 'minus' | 'parentheses' | 'red' })}
          style={{
            width: '100%',
            padding: '8px 12px',
            border: '1px solid #e8eaed',
            borderRadius: '4px',
            fontSize: '14px',
          }}
        >
          <option value="minus">-$123</option>
          <option value="parentheses">($123)</option>
          <option value="red">$123 (red)</option>
        </select>
      </div>
    </div>
  );
});

// Date/Time Format Options Component
const DateTimeFormatOptions = memo(function DateTimeFormatOptions({
  format,
  onChange,
  onTypeChange,
}: {
  format: CellFormat;
  onChange: (updates: Partial<CellFormat>) => void;
  onTypeChange: (type: FormatType) => void;
}) {
  return (
    <div>
      <div style={{ marginBottom: '16px' }}>
        <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
          Format Type
        </label>
        <select
          value={format.type || 'date'}
          onChange={(e) => onTypeChange(e.target.value as FormatType)}
          style={{
            width: '100%',
            padding: '8px 12px',
            border: '1px solid #e8eaed',
            borderRadius: '4px',
            fontSize: '14px',
          }}
        >
          <option value="date">Date</option>
          <option value="time">Time</option>
          <option value="datetime">Date and Time</option>
          <option value="duration">Duration</option>
        </select>
      </div>

      {(format.type === 'date' || format.type === 'datetime') && (
        <div style={{ marginBottom: '16px' }}>
          <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
            Date Format
          </label>
          <select
            value={format.dateFormat || 'MM/DD/YYYY'}
            onChange={(e) => onChange({ dateFormat: e.target.value })}
            style={{
              width: '100%',
              padding: '8px 12px',
              border: '1px solid #e8eaed',
              borderRadius: '4px',
              fontSize: '14px',
            }}
          >
            {DATE_FORMAT_OPTIONS.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>
      )}

      {(format.type === 'time' || format.type === 'datetime') && (
        <div style={{ marginBottom: '16px' }}>
          <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
            Time Format
          </label>
          <select
            value={format.timeFormat || 'HH:mm:ss'}
            onChange={(e) => onChange({ timeFormat: e.target.value })}
            style={{
              width: '100%',
              padding: '8px 12px',
              border: '1px solid #e8eaed',
              borderRadius: '4px',
              fontSize: '14px',
            }}
          >
            {TIME_FORMAT_OPTIONS.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>
      )}

      {format.type === 'duration' && (
        <div style={{ marginBottom: '16px' }}>
          <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
            Duration Unit
          </label>
          <select
            value={format.durationFormat || 'hours'}
            onChange={(e) => onChange({ durationFormat: e.target.value as CellFormat['durationFormat'] })}
            style={{
              width: '100%',
              padding: '8px 12px',
              border: '1px solid #e8eaed',
              borderRadius: '4px',
              fontSize: '14px',
            }}
          >
            <option value="hours">Hours (h)</option>
            <option value="minutes">Minutes (m)</option>
            <option value="seconds">Seconds (s)</option>
            <option value="milliseconds">Milliseconds (ms)</option>
          </select>
        </div>
      )}
    </div>
  );
});

// Custom Format Options Component
const CustomFormatOptions = memo(function CustomFormatOptions({
  pattern,
  onPatternChange,
}: {
  pattern: string;
  onPatternChange: (pattern: string) => void;
}) {
  const commonPatterns = [
    { pattern: '#,##0.00', description: 'Number with commas and 2 decimals' },
    { pattern: '0.0%', description: 'Percentage with 1 decimal' },
    { pattern: 'MM/DD/YYYY', description: 'Date format' },
    { pattern: '$#,##0.00', description: 'Currency format' },
    { pattern: '0.00E+00', description: 'Scientific notation' },
  ];

  return (
    <div>
      <div style={{ marginBottom: '16px' }}>
        <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
          Custom Format Pattern
        </label>
        <input
          type="text"
          value={pattern}
          onChange={(e) => onPatternChange(e.target.value)}
          placeholder="Enter format pattern (e.g., #,##0.00)"
          style={{
            width: '100%',
            padding: '8px 12px',
            border: '1px solid #e8eaed',
            borderRadius: '4px',
            fontSize: '14px',
            fontFamily: 'monospace',
          }}
        />
      </div>

      <div style={{ marginBottom: '16px' }}>
        <label style={{ display: 'block', marginBottom: '8px', fontSize: '14px', fontWeight: 500 }}>
          Common Patterns
        </label>
        <div style={{ display: 'grid', gap: '4px' }}>
          {commonPatterns.map((item) => (
            <button
              key={item.pattern}
              type="button"
              onClick={() => onPatternChange(item.pattern)}
              style={{
                padding: '6px 8px',
                border: '1px solid #e8eaed',
                borderRadius: '4px',
                backgroundColor: '#ffffff',
                color: '#5f6368',
                cursor: 'pointer',
                fontSize: '12px',
                textAlign: 'left',
                fontFamily: 'monospace',
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.backgroundColor = '#f1f3f4';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.backgroundColor = '#ffffff';
              }}
            >
              <div style={{ fontWeight: 500 }}>{item.pattern}</div>
              <div style={{ fontSize: '11px', color: '#9aa0a6' }}>{item.description}</div>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
});
