import { memo, useState, useEffect } from 'react';
import type { PageConfig, PageSizeKey } from '@b32nio/spindle-docs-core';
import { PAGE_SIZES } from '@b32nio/spindle-docs-core';

interface PageSetupModalProps {
  isOpen: boolean;
  pageConfig: PageConfig;
  onClose: () => void;
  onConfirm: (config: PageConfig) => void;
}

/**
 * Modal for configuring page settings
 */
export const PageSetupModal = memo(function PageSetupModal({
  isOpen,
  pageConfig,
  onClose,
  onConfirm,
}: PageSetupModalProps) {
  const [config, setConfig] = useState<PageConfig>(pageConfig);
  const [selectedSize, setSelectedSize] = useState<PageSizeKey | 'custom'>('LETTER');
  
  // Determine current page size name
  useEffect(() => {
    const effectiveSize = config.orientation === 'landscape'
      ? { w: config.size.h, h: config.size.w }
      : config.size;
    
    let found: PageSizeKey | 'custom' = 'custom';
    for (const [name, size] of Object.entries(PAGE_SIZES)) {
      if (size.w === effectiveSize.w && size.h === effectiveSize.h) {
        found = name as PageSizeKey;
        break;
      }
    }
    setSelectedSize(found);
  }, [config.size, config.orientation]);
  
  useEffect(() => {
    setConfig(pageConfig);
  }, [pageConfig, isOpen]);
  
  if (!isOpen) return null;
  
  const handleSizeChange = (sizeName: PageSizeKey) => {
    const size = PAGE_SIZES[sizeName];
    setConfig(prev => ({
      ...prev,
      size: { ...size },
    }));
    setSelectedSize(sizeName);
  };
  
  const handleOrientationChange = (orientation: 'portrait' | 'landscape') => {
    setConfig(prev => ({
      ...prev,
      orientation,
    }));
  };
  
  const handleMarginChange = (side: 'top' | 'right' | 'bottom' | 'left', value: number) => {
    setConfig(prev => ({
      ...prev,
      margins: {
        ...prev.margins,
        [side]: Math.max(0, value),
      },
    }));
  };
  
  const handleConfirm = () => {
    onConfirm(config);
    onClose();
  };
  
  const pageSizeOptions: Array<{ key: PageSizeKey; label: string; dimensions: string }> = [
    { key: 'LETTER', label: 'Letter', dimensions: '8.5" × 11"' },
    { key: 'LEGAL', label: 'Legal', dimensions: '8.5" × 14"' },
    { key: 'A4', label: 'A4', dimensions: '210mm × 297mm' },
    { key: 'A5', label: 'A5', dimensions: '148mm × 210mm' },
    { key: 'TABLOID', label: 'Tabloid', dimensions: '11" × 17"' },
    { key: 'EXECUTIVE', label: 'Executive', dimensions: '7.25" × 10.5"' },
    { key: 'B5', label: 'B5', dimensions: '182mm × 257mm' },
  ];
  
  // Convert pixels to inches for display
  const toInches = (px: number) => (px / 96).toFixed(2);
  const fromInches = (inches: string) => Math.round(parseFloat(inches) * 96);
  
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
    >
      <div
        style={{
          backgroundColor: '#ffffff',
          borderRadius: '8px',
          boxShadow: '0 4px 24px rgba(0, 0, 0, 0.2)',
          width: '480px',
          maxWidth: '90vw',
          maxHeight: '90vh',
          overflow: 'auto',
        }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div
          style={{
            padding: '16px 24px',
            borderBottom: '1px solid #e8eaed',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
          }}
        >
          <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 500, color: '#202124' }}>
            Page Setup
          </h2>
          <button
            onClick={onClose}
            style={{
              border: 'none',
              background: 'none',
              fontSize: '20px',
              cursor: 'pointer',
              color: '#5f6368',
              padding: '4px',
            }}
          >
            ×
          </button>
        </div>
        
        {/* Content */}
        <div style={{ padding: '24px' }}>
          {/* Page Size */}
          <div style={{ marginBottom: '24px' }}>
            <label
              style={{
                display: 'block',
                fontSize: '13px',
                fontWeight: 500,
                color: '#5f6368',
                marginBottom: '8px',
              }}
            >
              Page Size
            </label>
            <select
              value={selectedSize}
              onChange={(e) => handleSizeChange(e.target.value as PageSizeKey)}
              style={{
                width: '100%',
                padding: '10px 12px',
                fontSize: '14px',
                border: '1px solid #e8eaed',
                borderRadius: '4px',
                backgroundColor: '#ffffff',
                color: '#202124',
                cursor: 'pointer',
              }}
            >
              {pageSizeOptions.map((option) => (
                <option key={option.key} value={option.key}>
                  {option.label} ({option.dimensions})
                </option>
              ))}
            </select>
          </div>
          
          {/* Orientation */}
          <div style={{ marginBottom: '24px' }}>
            <label
              style={{
                display: 'block',
                fontSize: '13px',
                fontWeight: 500,
                color: '#5f6368',
                marginBottom: '8px',
              }}
            >
              Orientation
            </label>
            <div style={{ display: 'flex', gap: '12px' }}>
              <button
                onClick={() => handleOrientationChange('portrait')}
                style={{
                  flex: 1,
                  padding: '12px',
                  border: `2px solid ${config.orientation === 'portrait' ? '#1a73e8' : '#e8eaed'}`,
                  borderRadius: '4px',
                  backgroundColor: config.orientation === 'portrait' ? '#e8f0fe' : '#ffffff',
                  cursor: 'pointer',
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  gap: '8px',
                }}
              >
                <div
                  style={{
                    width: '32px',
                    height: '44px',
                    border: '2px solid',
                    borderColor: config.orientation === 'portrait' ? '#1a73e8' : '#5f6368',
                    borderRadius: '2px',
                  }}
                />
                <span
                  style={{
                    fontSize: '13px',
                    color: config.orientation === 'portrait' ? '#1a73e8' : '#5f6368',
                    fontWeight: config.orientation === 'portrait' ? 500 : 400,
                  }}
                >
                  Portrait
                </span>
              </button>
              <button
                onClick={() => handleOrientationChange('landscape')}
                style={{
                  flex: 1,
                  padding: '12px',
                  border: `2px solid ${config.orientation === 'landscape' ? '#1a73e8' : '#e8eaed'}`,
                  borderRadius: '4px',
                  backgroundColor: config.orientation === 'landscape' ? '#e8f0fe' : '#ffffff',
                  cursor: 'pointer',
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  gap: '8px',
                }}
              >
                <div
                  style={{
                    width: '44px',
                    height: '32px',
                    border: '2px solid',
                    borderColor: config.orientation === 'landscape' ? '#1a73e8' : '#5f6368',
                    borderRadius: '2px',
                  }}
                />
                <span
                  style={{
                    fontSize: '13px',
                    color: config.orientation === 'landscape' ? '#1a73e8' : '#5f6368',
                    fontWeight: config.orientation === 'landscape' ? 500 : 400,
                  }}
                >
                  Landscape
                </span>
              </button>
            </div>
          </div>
          
          {/* Margins */}
          <div style={{ marginBottom: '24px' }}>
            <label
              style={{
                display: 'block',
                fontSize: '13px',
                fontWeight: 500,
                color: '#5f6368',
                marginBottom: '8px',
              }}
            >
              Margins (inches)
            </label>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
              <div>
                <label style={{ fontSize: '12px', color: '#5f6368', display: 'block', marginBottom: '4px' }}>
                  Top
                </label>
                <input
                  type="number"
                  step="0.1"
                  min="0"
                  max="5"
                  value={toInches(config.margins.top)}
                  onChange={(e) => handleMarginChange('top', fromInches(e.target.value))}
                  style={{
                    width: '100%',
                    padding: '8px 12px',
                    fontSize: '14px',
                    border: '1px solid #e8eaed',
                    borderRadius: '4px',
                    boxSizing: 'border-box',
                  }}
                />
              </div>
              <div>
                <label style={{ fontSize: '12px', color: '#5f6368', display: 'block', marginBottom: '4px' }}>
                  Bottom
                </label>
                <input
                  type="number"
                  step="0.1"
                  min="0"
                  max="5"
                  value={toInches(config.margins.bottom)}
                  onChange={(e) => handleMarginChange('bottom', fromInches(e.target.value))}
                  style={{
                    width: '100%',
                    padding: '8px 12px',
                    fontSize: '14px',
                    border: '1px solid #e8eaed',
                    borderRadius: '4px',
                    boxSizing: 'border-box',
                  }}
                />
              </div>
              <div>
                <label style={{ fontSize: '12px', color: '#5f6368', display: 'block', marginBottom: '4px' }}>
                  Left
                </label>
                <input
                  type="number"
                  step="0.1"
                  min="0"
                  max="5"
                  value={toInches(config.margins.left)}
                  onChange={(e) => handleMarginChange('left', fromInches(e.target.value))}
                  style={{
                    width: '100%',
                    padding: '8px 12px',
                    fontSize: '14px',
                    border: '1px solid #e8eaed',
                    borderRadius: '4px',
                    boxSizing: 'border-box',
                  }}
                />
              </div>
              <div>
                <label style={{ fontSize: '12px', color: '#5f6368', display: 'block', marginBottom: '4px' }}>
                  Right
                </label>
                <input
                  type="number"
                  step="0.1"
                  min="0"
                  max="5"
                  value={toInches(config.margins.right)}
                  onChange={(e) => handleMarginChange('right', fromInches(e.target.value))}
                  style={{
                    width: '100%',
                    padding: '8px 12px',
                    fontSize: '14px',
                    border: '1px solid #e8eaed',
                    borderRadius: '4px',
                    boxSizing: 'border-box',
                  }}
                />
              </div>
            </div>
          </div>
          
          {/* Preview */}
          <div style={{ marginBottom: '24px' }}>
            <label
              style={{
                display: 'block',
                fontSize: '13px',
                fontWeight: 500,
                color: '#5f6368',
                marginBottom: '8px',
              }}
            >
              Preview
            </label>
            <div
              style={{
                display: 'flex',
                justifyContent: 'center',
                padding: '16px',
                backgroundColor: '#f8f9fa',
                borderRadius: '4px',
              }}
            >
              <div
                style={{
                  width: config.orientation === 'landscape' ? 120 : 90,
                  height: config.orientation === 'landscape' ? 90 : 120,
                  backgroundColor: '#ffffff',
                  border: '1px solid #e8eaed',
                  boxShadow: '0 1px 3px rgba(0, 0, 0, 0.1)',
                  position: 'relative',
                }}
              >
                {/* Content area */}
                <div
                  style={{
                    position: 'absolute',
                    top: `${(config.margins.top / (config.orientation === 'landscape' ? config.size.h : config.size.w)) * 100}%`,
                    left: `${(config.margins.left / (config.orientation === 'landscape' ? config.size.h : config.size.w)) * 100}%`,
                    right: `${(config.margins.right / (config.orientation === 'landscape' ? config.size.h : config.size.w)) * 100}%`,
                    bottom: `${(config.margins.bottom / (config.orientation === 'landscape' ? config.size.h : config.size.w)) * 100}%`,
                    border: '1px dashed #1a73e8',
                    backgroundColor: 'rgba(26, 115, 232, 0.05)',
                  }}
                />
              </div>
            </div>
          </div>
        </div>
        
        {/* Footer */}
        <div
          style={{
            padding: '16px 24px',
            borderTop: '1px solid #e8eaed',
            display: 'flex',
            justifyContent: 'flex-end',
            gap: '12px',
          }}
        >
          <button
            onClick={onClose}
            style={{
              padding: '10px 24px',
              fontSize: '14px',
              fontWeight: 500,
              border: '1px solid #e8eaed',
              borderRadius: '4px',
              backgroundColor: '#ffffff',
              color: '#5f6368',
              cursor: 'pointer',
            }}
          >
            Cancel
          </button>
          <button
            onClick={handleConfirm}
            style={{
              padding: '10px 24px',
              fontSize: '14px',
              fontWeight: 500,
              border: 'none',
              borderRadius: '4px',
              backgroundColor: '#1a73e8',
              color: '#ffffff',
              cursor: 'pointer',
            }}
          >
            Apply
          </button>
        </div>
      </div>
    </div>
  );
});

