import React, { useState, useEffect, useRef, useCallback, memo } from 'react';
import { EditorView } from 'prosemirror-view';
import { docsSchema, createCommands } from '@b32nio/spindle-docs-core';

interface LinkDialogProps {
  /** ProseMirror editor view */
  editorView: EditorView | null;
  /** Whether the dialog is open */
  isOpen: boolean;
  /** Callback when dialog is closed */
  onClose: () => void;
  /** Initial link URL (for editing existing links) */
  initialUrl?: string;
  /** Initial link text (selected text in editor) */
  initialText?: string;
  /** Whether editing an existing link */
  isEditing?: boolean;
}

/**
 * Dialog for inserting or editing links
 */
export const LinkDialog = memo(function LinkDialog({
  editorView,
  isOpen,
  onClose,
  initialUrl = '',
  initialText = '',
  isEditing = false,
}: LinkDialogProps) {
  const [url, setUrl] = useState(initialUrl);
  const [text, setText] = useState(initialText);
  const [error, setError] = useState('');
  const urlInputRef = useRef<HTMLInputElement>(null);
  const dialogRef = useRef<HTMLDivElement>(null);
  
  const commands = createCommands(docsSchema);
  
  // Reset state when dialog opens
  useEffect(() => {
    if (isOpen) {
      setUrl(initialUrl);
      setText(initialText);
      setError('');
      // Focus URL input when opening
      setTimeout(() => urlInputRef.current?.focus(), 0);
    }
  }, [isOpen, initialUrl, initialText]);
  
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
  
  // Validate URL
  const validateUrl = useCallback((urlString: string): boolean => {
    if (!urlString.trim()) {
      setError('URL is required');
      return false;
    }
    
    // Allow relative URLs and common protocols
    const validProtocols = ['http:', 'https:', 'mailto:', 'tel:'];
    try {
      // If it doesn't have a protocol, try adding https
      let testUrl = urlString;
      if (!urlString.includes('://') && !urlString.startsWith('mailto:') && !urlString.startsWith('tel:')) {
        testUrl = 'https://' + urlString;
      }
      const parsed = new URL(testUrl);
      if (!validProtocols.includes(parsed.protocol)) {
        setError('Invalid URL protocol');
        return false;
      }
    } catch {
      // Allow relative URLs (starting with / or not having protocol)
      if (!urlString.startsWith('/') && urlString.includes('.')) {
        // Likely a domain without protocol - that's okay
      } else if (!urlString.startsWith('/')) {
        setError('Invalid URL');
        return false;
      }
    }
    
    setError('');
    return true;
  }, []);
  
  // Handle form submission
  const handleSubmit = useCallback((e: React.FormEvent) => {
    e.preventDefault();
    
    if (!editorView) return;
    if (!validateUrl(url)) return;
    
    // Normalize URL - add https:// if no protocol
    let normalizedUrl = url.trim();
    if (!normalizedUrl.includes('://') && !normalizedUrl.startsWith('mailto:') && !normalizedUrl.startsWith('tel:') && !normalizedUrl.startsWith('/')) {
      normalizedUrl = 'https://' + normalizedUrl;
    }
    
    const { state, dispatch } = editorView;
    const { from, empty } = state.selection;
    
    // If there's no selection and we have link text, insert the text first
    if (empty && text.trim()) {
      // Insert text and then add link mark
      const tr = state.tr.insertText(text.trim(), from);
      const newTo = from + text.trim().length;
      tr.addMark(from, newTo, docsSchema.marks.link.create({ href: normalizedUrl }));
      dispatch(tr);
    } else if (!empty) {
      // Apply link to selection
      commands.insertLink({ href: normalizedUrl })(state, dispatch, editorView);
    } else {
      // No selection and no text - just insert the URL as link text
      const tr = state.tr.insertText(normalizedUrl, from);
      const newTo = from + normalizedUrl.length;
      tr.addMark(from, newTo, docsSchema.marks.link.create({ href: normalizedUrl }));
      dispatch(tr);
    }
    
    editorView.focus();
    onClose();
  }, [editorView, url, text, validateUrl, commands, onClose]);
  
  // Handle remove link
  const handleRemoveLink = useCallback(() => {
    if (!editorView) return;
    
    const { state, dispatch } = editorView;
    commands.removeLink()(state, dispatch, editorView);
    editorView.focus();
    onClose();
  }, [editorView, commands, onClose]);
  
  if (!isOpen) return null;
  
  return (
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
          width: '400px',
          maxWidth: '90vw',
          padding: '24px',
        }}
      >
        <h2 style={{ margin: '0 0 16px', fontSize: '18px', fontWeight: 500, color: '#202124' }}>
          {isEditing ? 'Edit Link' : 'Insert Link'}
        </h2>
        
        <form onSubmit={handleSubmit}>
          {/* URL Input */}
          <div style={{ marginBottom: '16px' }}>
            <label
              htmlFor="link-url"
              style={{
                display: 'block',
                marginBottom: '6px',
                fontSize: '13px',
                fontWeight: 500,
                color: '#5f6368',
              }}
            >
              URL
            </label>
            <input
              ref={urlInputRef}
              id="link-url"
              type="text"
              value={url}
              onChange={(e) => {
                setUrl(e.target.value);
                if (error) validateUrl(e.target.value);
              }}
              placeholder="https://example.com"
              style={{
                width: '100%',
                padding: '10px 12px',
                fontSize: '14px',
                border: error ? '1px solid #d93025' : '1px solid #dadce0',
                borderRadius: '4px',
                outline: 'none',
                boxSizing: 'border-box',
              }}
              onFocus={(e) => {
                e.target.style.borderColor = error ? '#d93025' : '#1a73e8';
                e.target.style.boxShadow = `0 0 0 2px ${error ? 'rgba(217, 48, 37, 0.2)' : 'rgba(26, 115, 232, 0.2)'}`;
              }}
              onBlur={(e) => {
                e.target.style.borderColor = error ? '#d93025' : '#dadce0';
                e.target.style.boxShadow = 'none';
              }}
            />
            {error && (
              <p style={{ margin: '4px 0 0', fontSize: '12px', color: '#d93025' }}>
                {error}
              </p>
            )}
          </div>
          
          {/* Text Input (only when no selection) */}
          {!initialText && (
            <div style={{ marginBottom: '16px' }}>
              <label
                htmlFor="link-text"
                style={{
                  display: 'block',
                  marginBottom: '6px',
                  fontSize: '13px',
                  fontWeight: 500,
                  color: '#5f6368',
                }}
              >
                Text to display (optional)
              </label>
              <input
                id="link-text"
                type="text"
                value={text}
                onChange={(e) => setText(e.target.value)}
                placeholder="Link text"
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
          )}
          
          {/* Buttons */}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '24px' }}>
            {isEditing && (
              <button
                type="button"
                onClick={handleRemoveLink}
                style={{
                  padding: '10px 16px',
                  fontSize: '14px',
                  fontWeight: 500,
                  color: '#d93025',
                  backgroundColor: 'transparent',
                  border: '1px solid #d93025',
                  borderRadius: '4px',
                  cursor: 'pointer',
                  marginRight: 'auto',
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.backgroundColor = 'rgba(217, 48, 37, 0.08)';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.backgroundColor = 'transparent';
                }}
              >
                Remove Link
              </button>
            )}
            
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
              style={{
                padding: '10px 24px',
                fontSize: '14px',
                fontWeight: 500,
                color: '#ffffff',
                backgroundColor: '#1a73e8',
                border: 'none',
                borderRadius: '4px',
                cursor: 'pointer',
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.backgroundColor = '#1557b0';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.backgroundColor = '#1a73e8';
              }}
            >
              {isEditing ? 'Update' : 'Apply'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
});

