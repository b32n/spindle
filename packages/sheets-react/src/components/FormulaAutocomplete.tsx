import { memo, useEffect, useRef } from 'react';
import type { FunctionDoc } from '@b32nio/spindle-sheets-core';

interface FormulaAutocompleteProps {
  suggestions: FunctionDoc[];
  highlightedIndex: number;
  /** Absolute position (within the editor's positioned container). */
  top: number;
  left: number;
  onHover: (index: number) => void;
  onPick: (index: number) => void;
}

/** Dropdown of function-name suggestions shown while typing a formula. */
export const FormulaAutocomplete = memo(function FormulaAutocomplete({
  suggestions,
  highlightedIndex,
  top,
  left,
  onHover,
  onPick,
}: FormulaAutocompleteProps) {
  const listRef = useRef<HTMLDivElement>(null);

  // Keep the highlighted row visible as the keyboard moves through the list.
  useEffect(() => {
    const row = listRef.current?.children[highlightedIndex] as HTMLElement | undefined;
    row?.scrollIntoView({ block: 'nearest' });
  }, [highlightedIndex]);

  if (suggestions.length === 0) return null;

  return (
    <div
      ref={listRef}
      style={{
        position: 'absolute',
        top,
        left,
        zIndex: 1100,
        minWidth: '280px',
        maxHeight: '240px',
        overflowY: 'auto',
        padding: '4px',
        background: '#ffffff',
        border: '1px solid rgba(15, 23, 42, 0.12)',
        borderRadius: '8px',
        boxShadow: '0 8px 24px rgba(15, 23, 42, 0.16)',
        fontFamily: '"Inter", "SF Pro Display", -apple-system, BlinkMacSystemFont, sans-serif',
        fontSize: '13px',
      }}
    >
      {suggestions.map((doc, index) => (
        <div
          key={doc.name}
          onMouseEnter={() => onHover(index)}
          // mousedown (not click) so the editor input never loses focus.
          onMouseDown={(e) => {
            e.preventDefault();
            onPick(index);
          }}
          style={{
            display: 'flex',
            alignItems: 'baseline',
            gap: '8px',
            padding: '4px 8px',
            borderRadius: '6px',
            cursor: 'pointer',
            whiteSpace: 'nowrap',
            background: index === highlightedIndex ? '#eef2ff' : 'transparent',
          }}
        >
          <span style={{ fontWeight: 600, color: '#1e293b' }}>{doc.name}</span>
          <span
            style={{
              color: '#94a3b8',
              fontSize: '11px',
              overflow: 'hidden',
              textOverflow: 'ellipsis',
            }}
          >
            {doc.description}
          </span>
        </div>
      ))}
    </div>
  );
});
