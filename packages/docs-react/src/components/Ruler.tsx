import React, { memo, useState, useCallback, useRef, useEffect } from 'react';
import type { PageConfig, PageMargins } from '@b32nio/spindle-docs-core';
import { useDocument } from '../context/DocumentContext';

interface RulerProps {
  pageConfig: PageConfig;
  onMarginsChange?: (margins: PageMargins) => void;
}

/**
 * Horizontal ruler with margin handles
 */
export const Ruler = memo(function Ruler({
  pageConfig,
  onMarginsChange,
}: RulerProps) {
  const { zoom } = useDocument();
  const rulerRef = useRef<HTMLDivElement>(null);
  
  const [dragging, setDragging] = useState<'left' | 'right' | null>(null);
  const [dragStart, setDragStart] = useState({ x: 0, margin: 0 });
  
  const scale = zoom / 100;
  
  // Calculate effective page width based on orientation
  const effectiveWidth = pageConfig.orientation === 'landscape'
    ? pageConfig.size.h
    : pageConfig.size.w;
  
  const rulerWidth = effectiveWidth * scale;
  const leftMargin = pageConfig.margins.left * scale;
  const rightMargin = pageConfig.margins.right * scale;
  
  // Generate tick marks
  const ticks: Array<{ position: number; major: boolean; label?: string }> = [];
  const pixelsPerInch = 96 * scale;
  const totalInches = effectiveWidth / 96;
  
  for (let inch = 0; inch <= totalInches; inch += 0.125) {
    const position = inch * pixelsPerInch;
    const isMajor = inch % 1 === 0;
    const isHalf = inch % 0.5 === 0 && !isMajor;
    
    if (isMajor || isHalf || inch % 0.25 === 0) {
      ticks.push({
        position,
        major: isMajor,
        label: isMajor ? String(inch) : undefined,
      });
    }
  }
  
  const handleMouseDown = useCallback((side: 'left' | 'right', e: React.MouseEvent) => {
    e.preventDefault();
    setDragging(side);
    setDragStart({
      x: e.clientX,
      margin: side === 'left' ? pageConfig.margins.left : pageConfig.margins.right,
    });
  }, [pageConfig.margins]);
  
  const handleMouseMove = useCallback((e: MouseEvent) => {
    if (!dragging) return;
    
    const delta = (e.clientX - dragStart.x) / scale;
    let newMargin = dragStart.margin + (dragging === 'left' ? delta : -delta);
    
    // Clamp margin values
    newMargin = Math.max(24, Math.min(effectiveWidth / 2 - 48, newMargin));
    
    const newMargins = { ...pageConfig.margins };
    if (dragging === 'left') {
      newMargins.left = Math.round(newMargin);
    } else {
      newMargins.right = Math.round(newMargin);
    }
    
    onMarginsChange?.(newMargins);
  }, [dragging, dragStart, scale, effectiveWidth, pageConfig.margins, onMarginsChange]);
  
  const handleMouseUp = useCallback(() => {
    setDragging(null);
  }, []);
  
  useEffect(() => {
    if (!dragging) return;
    
    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [dragging, handleMouseMove, handleMouseUp]);
  
  return (
    <div
      ref={rulerRef}
      className="ruler"
      style={{
        width: rulerWidth,
        height: 24 * scale,
        backgroundColor: '#f8f9fa',
        borderBottom: '1px solid #e8eaed',
        position: 'relative',
        overflow: 'hidden',
        userSelect: 'none',
      }}
    >
      {/* Margin areas (darker) */}
      <div
        style={{
          position: 'absolute',
          left: 0,
          top: 0,
          width: leftMargin,
          height: '100%',
          backgroundColor: '#e8eaed',
        }}
      />
      <div
        style={{
          position: 'absolute',
          right: 0,
          top: 0,
          width: rightMargin,
          height: '100%',
          backgroundColor: '#e8eaed',
        }}
      />
      
      {/* Tick marks */}
      {ticks.map((tick, index) => (
        <div
          key={index}
          style={{
            position: 'absolute',
            left: tick.position,
            bottom: 0,
            width: 1,
            height: tick.major ? 12 * scale : tick.label ? 8 * scale : 4 * scale,
            backgroundColor: '#5f6368',
          }}
        >
          {tick.label && (
            <span
              style={{
                position: 'absolute',
                bottom: '100%',
                left: '50%',
                transform: 'translateX(-50%)',
                fontSize: 9 * scale,
                color: '#5f6368',
                whiteSpace: 'nowrap',
              }}
            >
              {tick.label}
            </span>
          )}
        </div>
      ))}
      
      {/* Left margin handle */}
      <div
        onMouseDown={(e) => handleMouseDown('left', e)}
        style={{
          position: 'absolute',
          left: leftMargin - 6 * scale,
          top: '50%',
          transform: 'translateY(-50%)',
          width: 12 * scale,
          height: 16 * scale,
          backgroundColor: dragging === 'left' ? '#1a73e8' : '#5f6368',
          borderRadius: 2 * scale,
          cursor: 'ew-resize',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          transition: dragging ? 'none' : 'background-color 0.15s',
        }}
        onMouseEnter={(e) => {
          if (!dragging) {
            e.currentTarget.style.backgroundColor = '#1a73e8';
          }
        }}
        onMouseLeave={(e) => {
          if (!dragging) {
            e.currentTarget.style.backgroundColor = '#5f6368';
          }
        }}
      >
        <div
          style={{
            width: 2 * scale,
            height: 8 * scale,
            backgroundColor: '#ffffff',
            borderRadius: 1 * scale,
          }}
        />
      </div>
      
      {/* Right margin handle */}
      <div
        onMouseDown={(e) => handleMouseDown('right', e)}
        style={{
          position: 'absolute',
          right: rightMargin - 6 * scale,
          top: '50%',
          transform: 'translateY(-50%)',
          width: 12 * scale,
          height: 16 * scale,
          backgroundColor: dragging === 'right' ? '#1a73e8' : '#5f6368',
          borderRadius: 2 * scale,
          cursor: 'ew-resize',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          transition: dragging ? 'none' : 'background-color 0.15s',
        }}
        onMouseEnter={(e) => {
          if (!dragging) {
            e.currentTarget.style.backgroundColor = '#1a73e8';
          }
        }}
        onMouseLeave={(e) => {
          if (!dragging) {
            e.currentTarget.style.backgroundColor = '#5f6368';
          }
        }}
      >
        <div
          style={{
            width: 2 * scale,
            height: 8 * scale,
            backgroundColor: '#ffffff',
            borderRadius: 1 * scale,
          }}
        />
      </div>
    </div>
  );
});

