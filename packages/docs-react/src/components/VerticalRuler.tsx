import React, { memo, useState, useCallback, useRef, useEffect, useMemo } from 'react';
import type { PageConfig, PageMargins } from '@b32nio/spindle-docs-core';
import { useDocument } from '../context/DocumentContext';

interface VerticalRulerProps {
  pageConfig: PageConfig;
  onMarginsChange?: (margins: PageMargins) => void;
  /** The page index the ruler should align to */
  pageIndex: number;
  /** Height of each page in pixels (unscaled) */
  pageHeight: number;
  /** Gap between pages */
  pageGap?: number;
  /** Reference to the scroll container to sync with */
  scrollContainerRef?: React.RefObject<HTMLElement>;
}

/**
 * Vertical ruler with top/bottom margin handles
 * Uses CSS transform for GPU-accelerated positioning during scroll
 */
export const VerticalRuler = memo(function VerticalRuler({
  pageConfig,
  onMarginsChange,
  pageIndex,
  pageHeight: unscaledPageHeight,
  pageGap = 24,
  scrollContainerRef,
}: VerticalRulerProps) {
  const { zoom } = useDocument();
  const rulerRef = useRef<HTMLDivElement>(null);
  
  const [dragging, setDragging] = useState<'top' | 'bottom' | null>(null);
  const [dragStart, setDragStart] = useState({ y: 0, margin: 0 });
  
  const scale = zoom / 100;
  const pageHeight = unscaledPageHeight * scale;
  const scaledPageGap = pageGap;
  
  // Calculate effective page height based on orientation
  const effectiveHeight = pageConfig.orientation === 'landscape'
    ? pageConfig.size.w
    : pageConfig.size.h;
  
  const rulerHeight = effectiveHeight * scale;
  const rulerWidth = 24 * scale;
  const topMargin = pageConfig.margins.top * scale;
  const bottomMargin = pageConfig.margins.bottom * scale;
  
  // Calculate the absolute position of the active page
  const activePageTop = pageIndex * (pageHeight + scaledPageGap) + 24; // +24 for container padding
  
  // Sync ruler position with scroll using CSS transform (GPU-accelerated)
  useEffect(() => {
    const container = scrollContainerRef?.current;
    const ruler = rulerRef.current;
    if (!container || !ruler) return;
    
    let rafId: number | null = null;
    let lastScrollTop = -1;
    
    const updatePosition = () => {
      rafId = null;
      const scrollTop = container.scrollTop;
      
      // Skip if scroll position hasn't changed
      if (scrollTop === lastScrollTop) return;
      lastScrollTop = scrollTop;
      
      // Use transform for GPU-accelerated positioning
      const offset = activePageTop - scrollTop;
      ruler.style.transform = `translateY(${offset}px)`;
    };
    
    // Initial position
    updatePosition();
    
    const handleScroll = () => {
      // Debounce with rAF - only one update per frame
      if (rafId === null) {
        rafId = requestAnimationFrame(updatePosition);
      }
    };
    
    container.addEventListener('scroll', handleScroll, { passive: true });
    return () => {
      container.removeEventListener('scroll', handleScroll);
      if (rafId !== null) {
        cancelAnimationFrame(rafId);
      }
    };
  }, [scrollContainerRef, activePageTop]);
  
  // Memoize tick marks to avoid recalculation on every render
  const ticks = useMemo(() => {
    const result: Array<{ position: number; major: boolean; label?: string }> = [];
    const pixelsPerInch = 96 * scale;
    const totalInches = effectiveHeight / 96;
    
    for (let inch = 0; inch <= totalInches; inch += 0.125) {
      const position = inch * pixelsPerInch;
      const isMajor = inch % 1 === 0;
      const isHalf = inch % 0.5 === 0 && !isMajor;
      
      if (isMajor || isHalf || inch % 0.25 === 0) {
        result.push({
          position,
          major: isMajor,
          label: isMajor ? String(inch) : undefined,
        });
      }
    }
    return result;
  }, [scale, effectiveHeight]);
  
  const handleMouseDown = useCallback((side: 'top' | 'bottom', e: React.MouseEvent) => {
    e.preventDefault();
    setDragging(side);
    setDragStart({
      y: e.clientY,
      margin: side === 'top' ? pageConfig.margins.top : pageConfig.margins.bottom,
    });
  }, [pageConfig.margins]);
  
  const handleMouseMove = useCallback((e: MouseEvent) => {
    if (!dragging) return;
    
    const delta = (e.clientY - dragStart.y) / scale;
    let newMargin = dragStart.margin + (dragging === 'top' ? delta : -delta);
    
    // Clamp margin values
    newMargin = Math.max(24, Math.min(effectiveHeight / 2 - 48, newMargin));
    
    const newMargins = { ...pageConfig.margins };
    if (dragging === 'top') {
      newMargins.top = Math.round(newMargin);
    } else {
      newMargins.bottom = Math.round(newMargin);
    }
    
    onMarginsChange?.(newMargins);
  }, [dragging, dragStart, scale, effectiveHeight, pageConfig.margins, onMarginsChange]);
  
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
      className="vertical-ruler"
      style={{
        position: 'absolute',
        left: 0,
        top: 0,
        width: rulerWidth,
        height: rulerHeight,
        backgroundColor: '#f8f9fa',
        borderRight: '1px solid #e8eaed',
        overflow: 'hidden',
        userSelect: 'none',
        zIndex: 50,
        willChange: 'transform',
        contain: 'layout style paint',
      }}
    >
      {/* Margin areas (darker) */}
      <div
        style={{
          position: 'absolute',
          left: 0,
          top: 0,
          width: '100%',
          height: topMargin,
          backgroundColor: '#e8eaed',
        }}
      />
      <div
        style={{
          position: 'absolute',
          left: 0,
          bottom: 0,
          width: '100%',
          height: bottomMargin,
          backgroundColor: '#e8eaed',
        }}
      />
      
      {/* Tick marks */}
      {ticks.map((tick, index) => (
        <div
          key={index}
          style={{
            position: 'absolute',
            top: tick.position,
            right: 0,
            height: 1,
            width: tick.major ? 12 * scale : tick.label ? 8 * scale : 4 * scale,
            backgroundColor: '#5f6368',
          }}
        >
          {tick.label && (
            <span
              style={{
                position: 'absolute',
                right: '100%',
                top: '50%',
                transform: 'translateY(-50%) rotate(-90deg)',
                transformOrigin: 'right center',
                fontSize: 9 * scale,
                color: '#5f6368',
                whiteSpace: 'nowrap',
                marginRight: 2 * scale,
              }}
            >
              {tick.label}
            </span>
          )}
        </div>
      ))}
      
      {/* Top margin handle */}
      <div
        onMouseDown={(e) => handleMouseDown('top', e)}
        style={{
          position: 'absolute',
          top: topMargin - 6 * scale,
          left: '50%',
          transform: 'translateX(-50%)',
          width: 16 * scale,
          height: 12 * scale,
          backgroundColor: dragging === 'top' ? '#1a73e8' : '#5f6368',
          borderRadius: 2 * scale,
          cursor: 'ns-resize',
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
            width: 8 * scale,
            height: 2 * scale,
            backgroundColor: '#ffffff',
            borderRadius: 1 * scale,
          }}
        />
      </div>
      
      {/* Bottom margin handle */}
      <div
        onMouseDown={(e) => handleMouseDown('bottom', e)}
        style={{
          position: 'absolute',
          bottom: bottomMargin - 6 * scale,
          left: '50%',
          transform: 'translateX(-50%)',
          width: 16 * scale,
          height: 12 * scale,
          backgroundColor: dragging === 'bottom' ? '#1a73e8' : '#5f6368',
          borderRadius: 2 * scale,
          cursor: 'ns-resize',
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
            width: 8 * scale,
            height: 2 * scale,
            backgroundColor: '#ffffff',
            borderRadius: 1 * scale,
          }}
        />
      </div>
    </div>
  );
});
