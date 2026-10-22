import { createPortal } from 'react-dom';
import React, { useState, useEffect, useRef, useCallback, memo } from 'react';
import { EditorView } from 'prosemirror-view';
import { docsSchema, createCommands } from '@b32nio/spindle-docs-core';
import { useImageHooks, readImageFile, storeImageFile, resolveImageUrl, ImageError } from '@b32nio/spindle-shared/react';

interface ImageDialogProps {
  /** ProseMirror editor view */
  editorView: EditorView | null;
  /** Whether the dialog is open */
  isOpen: boolean;
  /** Callback when dialog is closed */
  onClose: () => void;
}

/**
 * Dialog for inserting images via URL or file upload
 */
export const ImageDialog = memo(function ImageDialog({
  editorView,
  isOpen,
  onClose,
}: ImageDialogProps) {
  const [url, setUrl] = useState('');
  const [alt, setAlt] = useState('');
  const [width, setWidth] = useState('');
  const [height, setHeight] = useState('');
  const [error, setError] = useState('');
  const [preview, setPreview] = useState<string | null>(null);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [isDragOver, setIsDragOver] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const [activeTab, setActiveTab] = useState<'upload' | 'url'>('upload');
  const hooks = useImageHooks();
  // Set while a local file is staged (its data URL is in `url`); uploaded via the host on Insert.
  const [pendingFile, setPendingFile] = useState<File | null>(null);
  // The data URL `resolveImage` produced for the typed URL.
  const [resolved, setResolved] = useState<string | null>(null);
  
  const urlInputRef = useRef<HTMLInputElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const dialogRef = useRef<HTMLDivElement>(null);
  const dropZoneRef = useRef<HTMLDivElement>(null);
  
  const commands = createCommands(docsSchema);
  
  // Reset state when dialog opens
  useEffect(() => {
    if (isOpen) {
      setUrl('');
      setAlt('');
      setWidth('');
      setHeight('');
      setError('');
      setPreview(null);
      setIsDragOver(false);
      setIsConverting(false);
      setActiveTab('upload');
      setPendingFile(null);
      setResolved(null);
    }
  }, [isOpen]);
  
  // Handle click outside to close
  useEffect(() => {
    if (!isOpen) return;
    
    const handleClickOutside = (e: MouseEvent) => {
      if (dialogRef.current && !dialogRef.current.contains(e.target as Node)) {
        onClose();
      }
    };
    
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [isOpen, onClose]);
  
  // Handle escape key
  useEffect(() => {
    if (!isOpen) return;
    
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        onClose();
      }
    };
    
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose]);
  
  // Handle file selection (from input or drop)
  const handleFile = useCallback(async (file: File) => {
    setError('');
    setPreviewLoading(true);
    try {
      const dataUrl = await readImageFile(hooks, file);
      setPreview(dataUrl);
      setUrl(dataUrl);
      setPendingFile(file);
      setPreviewLoading(false);

      // Use filename as alt text if empty
      if (!alt) {
        const name = file.name.replace(/\.[^/.]+$/, '');
        setAlt(name);
      }
    } catch (err) {
      setError(err instanceof ImageError ? err.message : 'Failed to read image file');
      setPreviewLoading(false);
    }
  }, [alt, hooks]);

  // Handle file input change
  const handleFileChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      handleFile(file);
    }
  }, [handleFile]);
  
  // Handle drag events
  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(true);
  }, []);
  
  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
  }, []);
  
  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
    
    const file = e.dataTransfer.files?.[0];
    if (file) {
      handleFile(file);
    }
  }, [handleFile]);
  
  // Fetch the typed URL through the host's resolveImage hook and preview the result.
  const validateAndPreview = useCallback(async (urlString: string): Promise<string | null> => {
    if (!urlString.trim()) {
      setError('');
      setPreview(null);
      setResolved(null);
      return null;
    }

    let normalizedUrl = urlString.trim();
    if (!normalizedUrl.startsWith('http://') && !normalizedUrl.startsWith('https://') && !normalizedUrl.startsWith('data:')) {
      normalizedUrl = 'https://' + normalizedUrl;
    }

    setPreviewLoading(true);
    try {
      const dataUrl = await resolveImageUrl(hooks, normalizedUrl);
      setPreview(dataUrl);
      setResolved(dataUrl);
      setError('');
      return dataUrl;
    } catch (err) {
      setPreview(null);
      setResolved(null);
      setError(err instanceof ImageError ? err.message : 'Could not load image. Please check the URL.');
      return null;
    } finally {
      setPreviewLoading(false);
    }
  }, [hooks]);

  // Handle URL blur - trigger preview
  const handleUrlBlur = useCallback(() => {
    if (url.trim() && !url.startsWith('data:')) {
      validateAndPreview(url);
    }
  }, [url, validateAndPreview]);
  
  // Handle form submission
  const handleSubmit = useCallback(async (e: React.FormEvent) => {
    e.preventDefault();
    
    if (!editorView) return;
    
    if (!url.trim()) {
      setError('Please upload an image or provide a URL');
      return;
    }
    
    setIsConverting(true);
    
    try {
      let src: string;
      if (activeTab === 'url') {
        const dataUrl = resolved ?? (await validateAndPreview(url));
        if (!dataUrl) return; // validateAndPreview already showed the error
        src = dataUrl;
      } else if (pendingFile) {
        src = await storeImageFile(hooks, pendingFile, url);
      } else {
        src = url.trim();
      }

      const { state, dispatch } = editorView;
      
      // Build image attributes
      const attrs: { src: string; alt?: string; width?: number; height?: number } = {
        src,
      };
      
      if (alt.trim()) {
        attrs.alt = alt.trim();
      }
      
      const parsedWidth = parseInt(width, 10);
      const parsedHeight = parseInt(height, 10);
      
      if (!isNaN(parsedWidth) && parsedWidth > 0) {
        attrs.width = parsedWidth;
      }
      if (!isNaN(parsedHeight) && parsedHeight > 0) {
        attrs.height = parsedHeight;
      }
      
      // Insert image
      commands.insertImage(attrs)(state, dispatch, editorView);
      
      editorView.focus();
      onClose();
    } catch (err) {
      setError(err instanceof ImageError ? err.message : 'Failed to process image');
    } finally {
      setIsConverting(false);
    }
  }, [editorView, url, alt, width, height, commands, onClose, activeTab, resolved, pendingFile, hooks, validateAndPreview]);
  
  if (!isOpen) return null;
  
  const tabStyle = (isActive: boolean) => ({
    padding: '10px 20px',
    fontSize: '14px',
    fontWeight: 500,
    color: isActive ? '#1a73e8' : '#5f6368',
    backgroundColor: 'transparent',
    border: 'none',
    borderBottom: isActive ? '2px solid #1a73e8' : '2px solid transparent',
    cursor: 'pointer',
    transition: 'all 0.2s',
  });
  
  return createPortal(
    <div
      style={{
        position: 'fixed',
        inset: 0,
        backgroundColor: 'rgba(0, 0, 0, 0.3)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 10000,
      }}
    >
      <div
        ref={dialogRef}
        style={{
          backgroundColor: '#ffffff',
          borderRadius: '8px',
          boxShadow: '0 4px 24px rgba(0, 0, 0, 0.2)',
          width: '520px',
          maxWidth: '90vw',
          maxHeight: '90vh',
          overflow: 'auto',
        }}
      >
        <h2 style={{ margin: 0, padding: '20px 24px 0', fontSize: '18px', fontWeight: 500, color: '#202124' }}>
          Insert Image
        </h2>
        
        {/* Tabs */}
        <div style={{ display: 'flex', borderBottom: '1px solid #e8eaed', marginTop: '8px' }}>
          <button
            type="button"
            style={tabStyle(activeTab === 'upload')}
            onClick={() => setActiveTab('upload')}
          >
            Upload
          </button>
          {hooks.resolveImage && <button
            type="button"
            style={tabStyle(activeTab === 'url')}
            onClick={() => {
              setActiveTab('url');
              setTimeout(() => urlInputRef.current?.focus(), 0);
            }}
          >
            By URL
          </button>}
        </div>
        
        <form onSubmit={handleSubmit} style={{ padding: '20px 24px 24px' }}>
          {/* Upload Tab */}
          {activeTab === 'upload' && (
            <div style={{ marginBottom: '16px' }}>
              {/* Drop Zone */}
              <div
                ref={dropZoneRef}
                onClick={() => fileInputRef.current?.click()}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                style={{
                  border: `2px dashed ${isDragOver ? '#1a73e8' : '#dadce0'}`,
                  borderRadius: '8px',
                  padding: preview ? '16px' : '40px 20px',
                  textAlign: 'center',
                  cursor: 'pointer',
                  backgroundColor: isDragOver ? 'rgba(26, 115, 232, 0.05)' : '#fafafa',
                  transition: 'all 0.2s',
                }}
              >
                {previewLoading ? (
                  <div style={{ color: '#5f6368', fontSize: '14px' }}>
                    Loading...
                  </div>
                ) : preview ? (
                  <div>
                    <img
                      src={preview}
                      alt="Preview"
                      style={{
                        maxWidth: '100%',
                        maxHeight: '200px',
                        objectFit: 'contain',
                        borderRadius: '4px',
                      }}
                    />
                    <p style={{ margin: '12px 0 0', fontSize: '13px', color: '#5f6368' }}>
                      Click or drag to replace
                    </p>
                  </div>
                ) : (
                  <>
                    <div style={{ marginBottom: '12px' }}>
                      <svg width="48" height="48" viewBox="0 0 48 48" fill="none" style={{ opacity: 0.5 }}>
                        <path
                          d="M40 32V8c0-2.2-1.8-4-4-4H12c-2.2 0-4 1.8-4 4v32c0 2.2 1.8 4 4 4h24c2.2 0 4-1.8 4-4v-8"
                          stroke="#5f6368"
                          strokeWidth="2"
                          strokeLinecap="round"
                        />
                        <circle cx="18" cy="16" r="3" fill="#5f6368" />
                        <path
                          d="M8 36l8-8 4 4 8-8 12 12"
                          stroke="#5f6368"
                          strokeWidth="2"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        />
                      </svg>
                    </div>
                    <p style={{ margin: 0, fontSize: '14px', color: '#202124', fontWeight: 500 }}>
                      Drag and drop an image here
                    </p>
                    <p style={{ margin: '8px 0 0', fontSize: '13px', color: '#5f6368' }}>
                      or click to browse
                    </p>
                    <p style={{ margin: '8px 0 0', fontSize: '12px', color: '#9aa0a6' }}>
                      PNG, JPG, GIF, WebP up to {Math.round((hooks.maxImageBytes ?? 10 * 1024 * 1024) / (1024 * 1024))}MB
                    </p>
                    {!hooks.resolveImage && (
                      <p style={{ margin: '8px 0 0', fontSize: '12px', color: '#9aa0a6' }}>
                        Images from a web address are turned off. Upload a file or paste one instead.
                      </p>
                    )}
                  </>
                )}
              </div>
              
              <input
                ref={fileInputRef}
                type="file"
                accept="image/*"
                onChange={handleFileChange}
                style={{ display: 'none' }}
              />
            </div>
          )}
          
          {/* URL Tab */}
          {activeTab === 'url' && (
            <>
              {/* URL Input */}
              <div style={{ marginBottom: '16px' }}>
                <label
                  htmlFor="image-url"
                  style={{
                    display: 'block',
                    marginBottom: '6px',
                    fontSize: '13px',
                    fontWeight: 500,
                    color: '#5f6368',
                  }}
                >
                  Image URL
                </label>
                <input
                  ref={urlInputRef}
                  id="image-url"
                  type="text"
                  value={url.startsWith('data:') ? '' : url}
                  onChange={(e) => {
                    setUrl(e.target.value);
                    setPreview(null);
                    setResolved(null);
                  }}
                  onBlur={handleUrlBlur}
                  placeholder="https://example.com/image.png"
                  style={{
                    width: '100%',
                    padding: '10px 12px',
                    fontSize: '14px',
                    border: error && activeTab === 'url' ? '1px solid #d93025' : '1px solid #dadce0',
                    borderRadius: '4px',
                    outline: 'none',
                    boxSizing: 'border-box',
                  }}
                  onFocus={(e) => {
                    e.target.style.borderColor = error ? '#d93025' : '#1a73e8';
                    e.target.style.boxShadow = `0 0 0 2px ${error ? 'rgba(217, 48, 37, 0.2)' : 'rgba(26, 115, 232, 0.2)'}`;
                  }}
                />
                <p style={{ margin: '6px 0 0', fontSize: '12px', color: '#5f6368' }}>
                  Images will be converted to base64 for offline access
                </p>
              </div>
              
              {/* Preview */}
              {(preview || previewLoading) && activeTab === 'url' && (
                <div style={{ marginBottom: '16px' }}>
                  <label
                    style={{
                      display: 'block',
                      marginBottom: '6px',
                      fontSize: '13px',
                      fontWeight: 500,
                      color: '#5f6368',
                    }}
                  >
                    Preview
                  </label>
                  <div
                    style={{
                      border: '1px solid #dadce0',
                      borderRadius: '4px',
                      padding: '8px',
                      backgroundColor: '#f8f9fa',
                      minHeight: '100px',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                    }}
                  >
                    {previewLoading ? (
                      <span style={{ color: '#5f6368', fontSize: '13px' }}>Loading preview...</span>
                    ) : preview ? (
                      <img
                        src={preview}
                        alt="Preview"
                        style={{
                          maxWidth: '100%',
                          maxHeight: '200px',
                          objectFit: 'contain',
                        }}
                      />
                    ) : null}
                  </div>
                </div>
              )}
            </>
          )}
          
          {/* Error */}
          {error && (
            <p style={{ margin: '0 0 16px', fontSize: '13px', color: '#d93025' }}>
              {error}
            </p>
          )}
          
          {/* Alt Text */}
          <div style={{ marginBottom: '16px' }}>
            <label
              htmlFor="image-alt"
              style={{
                display: 'block',
                marginBottom: '6px',
                fontSize: '13px',
                fontWeight: 500,
                color: '#5f6368',
              }}
            >
              Alt text (optional)
            </label>
            <input
              id="image-alt"
              type="text"
              value={alt}
              onChange={(e) => setAlt(e.target.value)}
              placeholder="Describe the image for accessibility"
              style={{
                width: '100%',
                padding: '10px 12px',
                fontSize: '14px',
                border: '1px solid #dadce0',
                borderRadius: '4px',
                outline: 'none',
                boxSizing: 'border-box',
              }}
              onFocus={(e) => {
                e.target.style.borderColor = '#1a73e8';
                e.target.style.boxShadow = '0 0 0 2px rgba(26, 115, 232, 0.2)';
              }}
              onBlur={(e) => {
                e.target.style.borderColor = '#dadce0';
                e.target.style.boxShadow = 'none';
              }}
            />
          </div>
          
          {/* Dimensions */}
          <div style={{ display: 'flex', gap: '16px', marginBottom: '16px' }}>
            <div style={{ flex: 1 }}>
              <label
                htmlFor="image-width"
                style={{
                  display: 'block',
                  marginBottom: '6px',
                  fontSize: '13px',
                  fontWeight: 500,
                  color: '#5f6368',
                }}
              >
                Width (px, optional)
              </label>
              <input
                id="image-width"
                type="number"
                value={width}
                onChange={(e) => setWidth(e.target.value)}
                placeholder="Auto"
                min="1"
                style={{
                  width: '100%',
                  padding: '10px 12px',
                  fontSize: '14px',
                  border: '1px solid #dadce0',
                  borderRadius: '4px',
                  outline: 'none',
                  boxSizing: 'border-box',
                }}
                onFocus={(e) => {
                  e.target.style.borderColor = '#1a73e8';
                  e.target.style.boxShadow = '0 0 0 2px rgba(26, 115, 232, 0.2)';
                }}
                onBlur={(e) => {
                  e.target.style.borderColor = '#dadce0';
                  e.target.style.boxShadow = 'none';
                }}
              />
            </div>
            <div style={{ flex: 1 }}>
              <label
                htmlFor="image-height"
                style={{
                  display: 'block',
                  marginBottom: '6px',
                  fontSize: '13px',
                  fontWeight: 500,
                  color: '#5f6368',
                }}
              >
                Height (px, optional)
              </label>
              <input
                id="image-height"
                type="number"
                value={height}
                onChange={(e) => setHeight(e.target.value)}
                placeholder="Auto"
                min="1"
                style={{
                  width: '100%',
                  padding: '10px 12px',
                  fontSize: '14px',
                  border: '1px solid #dadce0',
                  borderRadius: '4px',
                  outline: 'none',
                  boxSizing: 'border-box',
                }}
                onFocus={(e) => {
                  e.target.style.borderColor = '#1a73e8';
                  e.target.style.boxShadow = '0 0 0 2px rgba(26, 115, 232, 0.2)';
                }}
                onBlur={(e) => {
                  e.target.style.borderColor = '#dadce0';
                  e.target.style.boxShadow = 'none';
                }}
              />
            </div>
          </div>
          
          {/* Buttons */}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '24px' }}>
            <button
              type="button"
              onClick={onClose}
              style={{
                padding: '10px 16px',
                fontSize: '14px',
                fontWeight: 500,
                color: '#5f6368',
                backgroundColor: 'transparent',
                border: '1px solid #dadce0',
                borderRadius: '4px',
                cursor: 'pointer',
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.backgroundColor = '#f1f3f4';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.backgroundColor = 'transparent';
              }}
            >
              Cancel
            </button>
            
            <button
              type="submit"
              disabled={!url.trim() || !!error || isConverting}
              style={{
                padding: '10px 24px',
                fontSize: '14px',
                fontWeight: 500,
                color: '#ffffff',
                backgroundColor: (!url.trim() || !!error || isConverting) ? '#9aa0a6' : '#1a73e8',
                border: 'none',
                borderRadius: '4px',
                cursor: (!url.trim() || !!error || isConverting) ? 'not-allowed' : 'pointer',
              }}
              onMouseEnter={(e) => {
                if (url.trim() && !error && !isConverting) {
                  e.currentTarget.style.backgroundColor = '#1557b0';
                }
              }}
              onMouseLeave={(e) => {
                if (url.trim() && !error && !isConverting) {
                  e.currentTarget.style.backgroundColor = '#1a73e8';
                }
              }}
            >
              {isConverting ? 'Processing...' : 'Insert'}
            </button>
          </div>
        </form>
      </div>
    </div>
  , document.body);
});
