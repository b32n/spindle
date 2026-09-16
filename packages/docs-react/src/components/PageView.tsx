import React, { memo, useMemo } from 'react';
import { PAGE_SIZES } from '@b32nio/spindle-docs-core';
import type { Section, PageConfig } from '@b32nio/spindle-docs-core';
import { useDocument } from '../context/DocumentContext';

interface PageViewProps {
  section: Section;
  pageNumber: number;
  children: React.ReactNode;
}

/**
 * Renders a single page with proper dimensions based on PageConfig
 */
export const PageView = memo(function PageView({
  section,
  pageNumber,
  children,
}: PageViewProps) {
  const { zoom } = useDocument();
  const { pageConfig } = section;
  
  // Calculate effective dimensions based on orientation
  const effectiveSize = useMemo(() => {
    if (pageConfig.orientation === 'landscape') {
      return { w: pageConfig.size.h, h: pageConfig.size.w };
    }
    return pageConfig.size;
  }, [pageConfig.size, pageConfig.orientation]);
  
  // Calculate content area dimensions
  const contentArea = useMemo(() => ({
    width: effectiveSize.w - pageConfig.margins.left - pageConfig.margins.right,
    height: effectiveSize.h - pageConfig.margins.top - pageConfig.margins.bottom,
    top: pageConfig.margins.top,
    left: pageConfig.margins.left,
  }), [effectiveSize, pageConfig.margins]);
  
  const scale = zoom / 100;

  return (
    <div
      className="page-view"
      style={{
        width: effectiveSize.w * scale,
        height: effectiveSize.h * scale,
        backgroundColor: '#ffffff',
        boxShadow: '0 1px 3px rgba(0, 0, 0, 0.12), 0 1px 2px rgba(0, 0, 0, 0.24)',
        position: 'relative',
        overflow: 'hidden',
        marginBottom: 24 * scale,
        transformOrigin: 'top left',
      }}
    >
      {/* Page content area */}
      <div
        className="page-content"
        style={{
          position: 'absolute',
          top: pageConfig.margins.top * scale,
          left: pageConfig.margins.left * scale,
          width: contentArea.width * scale,
          height: contentArea.height * scale,
          overflow: 'hidden',
        }}
      >
        {children}
      </div>
      
      {/* Page number */}
      <div
        className="page-number"
        style={{
          position: 'absolute',
          bottom: (pageConfig.margins.bottom / 2 - 8) * scale,
          left: 0,
          right: 0,
          textAlign: 'center',
          fontSize: 10 * scale,
          color: '#5f6368',
          fontFamily: 'Arial, sans-serif',
        }}
      >
        {pageNumber}
      </div>
      
      {/* Margin guides (visible in edit mode) */}
      <div
        className="margin-guides"
        style={{
          position: 'absolute',
          top: pageConfig.margins.top * scale,
          left: pageConfig.margins.left * scale,
          right: pageConfig.margins.right * scale,
          bottom: pageConfig.margins.bottom * scale,
          border: '1px dashed rgba(0, 0, 0, 0.08)',
          pointerEvents: 'none',
        }}
      />
    </div>
  );
});

/**
 * Get the page size name from dimensions
 */
export function getPageSizeName(config: PageConfig): string {
  const effectiveSize = config.orientation === 'landscape'
    ? { w: config.size.h, h: config.size.w }
    : config.size;
    
  for (const [name, size] of Object.entries(PAGE_SIZES)) {
    const s = size as { w: number; h: number };
    if (s.w === effectiveSize.w && s.h === effectiveSize.h) {
      return name;
    }
    // Check landscape
    if (s.h === effectiveSize.w && s.w === effectiveSize.h) {
      return name;
    }
  }
  return 'Custom';
}

