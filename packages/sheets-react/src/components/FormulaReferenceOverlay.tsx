import { memo, useMemo } from 'react';
import type { Sheet } from '@b32nio/spindle-sheets-core';
import type { FormulaRange } from '@b32nio/spindle-sheets-core';

interface FormulaReferenceOverlayProps {
  formulaRanges: FormulaRange[];
  sheet: Sheet;
  getRowTop: (row: number) => number;
  getColLeft: (col: number) => number;
}

// Color palette for highlighting different cell references (like Google Sheets)
// Generate random colors for each range
const HIGHLIGHT_COLORS = [
  { border: '#4285F4', background: 'rgba(66, 133, 244, 0.15)' }, // Blue
  { border: '#EA4335', background: 'rgba(234, 67, 53, 0.15)' }, // Red
  { border: '#FBBC04', background: 'rgba(251, 188, 4, 0.15)' }, // Yellow
  { border: '#34A853', background: 'rgba(52, 168, 83, 0.15)' }, // Green
  { border: '#9C27B0', background: 'rgba(156, 39, 176, 0.15)' }, // Purple
  { border: '#FF9800', background: 'rgba(255, 152, 0, 0.15)' }, // Orange
  { border: '#00BCD4', background: 'rgba(0, 188, 212, 0.15)' }, // Cyan
  { border: '#E91E63', background: 'rgba(233, 30, 99, 0.15)' }, // Pink
];

// Generate a random color for ranges beyond the palette
function getRandomColor(index: number): { border: string; background: string } {
  if (index < HIGHLIGHT_COLORS.length) {
    return HIGHLIGHT_COLORS[index];
  }
  
  // Generate a random color for ranges beyond the palette
  const hue = (index * 137.508) % 360; // Golden angle approximation for better distribution
  const saturation = 60 + (index % 20); // Vary saturation between 60-80%
  const lightness = 50 + (index % 10); // Vary lightness between 50-60%
  
  return {
    border: `hsl(${hue}, ${saturation}%, ${lightness}%)`,
    background: `hsla(${hue}, ${saturation}%, ${lightness}%, 0.15)`,
  };
}

export const FormulaReferenceOverlay = memo(function FormulaReferenceOverlay({
  formulaRanges,
  sheet,
  getRowTop,
  getColLeft,
}: FormulaReferenceOverlayProps) {
  const rangePositions = useMemo(() => {
    return formulaRanges.map((range, index) => {
      // Calculate the top-left corner of the range
      const top = getRowTop(range.startRow);
      const left = getColLeft(range.startCol);
      
      // Calculate the width and height of the range
      let width = 0;
      let height = 0;
      
      for (let r = range.startRow; r <= range.endRow; r++) {
        height += sheet.getRowHeight(r);
      }
      
      for (let c = range.startCol; c <= range.endCol; c++) {
        width += sheet.getColWidth(c);
      }
      
      const color = getRandomColor(index);

      return {
        top,
        left,
        width,
        height,
        color,
        key: `${range.startRow}:${range.startCol}:${range.endRow}:${range.endCol}:${index}`,
      };
    });
  }, [formulaRanges, sheet, getRowTop, getColLeft]);

  if (rangePositions.length === 0) {
    return null;
  }

  return (
    <>
      {rangePositions.map(({ top, left, width, height, color, key }) => (
        <div
          key={key}
          className="formula-reference-overlay"
          style={{
            position: 'absolute',
            top,
            left,
            width,
            height,
            border: `2px solid ${color.border}`,
            backgroundColor: color.background,
            pointerEvents: 'none',
            boxSizing: 'border-box',
            zIndex: 20, // Above selection overlay
          }}
        />
      ))}
    </>
  );
});

