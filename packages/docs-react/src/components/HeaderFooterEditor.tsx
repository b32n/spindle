/**
 * HeaderFooterEditor - Google Docs-style inline editing for headers and footers
 * 
 * Architecture:
 * - Creates a separate ProseMirror editor for the header/footer content
 * - Renders in the actual header/footer position on the page
 * - The main document is dimmed and non-interactive while editing
 * - Full toolbar support for formatting, images, etc.
 */

import { memo, useState, useCallback, useEffect, useRef } from 'react';
import { EditorState, Transaction } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { Schema, Node as PmNode, type Mark } from 'prosemirror-model';
import { TextSelection } from 'prosemirror-state';
import { baseKeymap, toggleMark } from 'prosemirror-commands';
import { keymap } from 'prosemirror-keymap';
import { history, undo, redo } from 'prosemirror-history';
import type { HeaderFooterContent, HeaderFooterParagraph, DynamicFieldType, HeaderFooterImageRun } from '@b32nio/spindle-docs-core';
import { 
  FileText, CalendarDays, Clock, Hash, Layers,
  Bold, Italic, Underline, 
  AlignLeft, AlignCenter, AlignRight,
  Image, Palette, ChevronDown, Upload
} from 'lucide-react';
import { ColorPicker } from './ColorPicker';

// ============================================================================
// Types
// ============================================================================

export interface HeaderFooterEditorProps {
  /** Type of area being edited */
  type: 'header' | 'footer';
  /** Current content */
  content: HeaderFooterContent;
  /** Called when content changes */
  onChange: (content: HeaderFooterContent) => void;
  /** Whether editing is active */
  isEditing: boolean;
  /** Called when user wants to stop editing */
  onClose: () => void;
  /** Page index (for first page logic) */
  pageIndex: number;
  /** Total pages */
  totalPages: number;
  /** Document title */
  documentTitle?: string;
  /** Width of the content area in pixels */
  contentWidth: number;
  /** Scale factor */
  scale: number;
  /** Position relative to the page */
  position: {
    top?: number;
    bottom?: number;
    left: number;
    pageY: number;
  };
  /** Main editor view (for toolbar integration, if needed) */
  mainEditorView?: EditorView | null;
}

// ============================================================================
// Simple Schema for Header/Footer
// ============================================================================

// Schema that supports formatting, images, and alignment
const headerFooterSchema = new Schema({
  nodes: {
    doc: { content: 'paragraph+' },
    paragraph: {
      content: 'inline*',
      group: 'block',
      attrs: { alignment: { default: 'center' } },
      parseDOM: [{ tag: 'p' }],
      toDOM(node) {
        return ['p', { style: `text-align: ${node.attrs.alignment}; margin: 0;` }, 0];
      },
    },
    text: { group: 'inline' },
    dynamicField: {
      group: 'inline',
      inline: true,
      atom: true,
      attrs: { fieldType: { default: 'pageNumber' } },
      parseDOM: [{
        tag: 'span.dynamic-field',
        getAttrs(dom) {
          return { fieldType: (dom as HTMLElement).dataset.fieldType || 'pageNumber' };
        },
      }],
      toDOM(node) {
        const labels: Record<string, string> = {
          pageNumber: '{PAGE}',
          totalPages: '{PAGES}',
          date: '{DATE}',
          time: '{TIME}',
          title: '{TITLE}',
        };
        return ['span', {
          class: 'dynamic-field',
          'data-field-type': node.attrs.fieldType,
          contenteditable: 'false',
          style: 'background: #e8f0fe; padding: 0 4px; border-radius: 3px; color: #1967d2; font-family: monospace; font-size: 0.9em;',
        }, labels[node.attrs.fieldType] || `{${node.attrs.fieldType.toUpperCase()}}`];
      },
    },
    image: {
      group: 'inline',
      inline: true,
      attrs: {
        src: {},
        alt: { default: '' },
        width: { default: 24 },
        height: { default: 24 },
      },
      parseDOM: [{
        tag: 'img[src]',
        getAttrs(dom) {
          const el = dom as HTMLImageElement;
          return {
            src: el.getAttribute('src'),
            alt: el.getAttribute('alt') || '',
            width: parseInt(el.getAttribute('width') || '24', 10),
            height: parseInt(el.getAttribute('height') || '24', 10),
          };
        },
      }],
      toDOM(node) {
        return ['img', {
          src: node.attrs.src,
          alt: node.attrs.alt,
          width: node.attrs.width,
          height: node.attrs.height,
          style: 'vertical-align: middle; max-height: 32px;',
        }];
      },
    },
  },
  marks: {
    bold: {
      parseDOM: [
        { tag: 'strong' },
        { tag: 'b' },
        { style: 'font-weight', getAttrs: (value) => (value as string).match(/^(bold|[7-9]\d\d)$/) && null },
      ],
      toDOM() { return ['strong', 0]; },
    },
    italic: {
      parseDOM: [{ tag: 'em' }, { tag: 'i' }, { style: 'font-style=italic' }],
      toDOM() { return ['em', 0]; },
    },
    underline: {
      parseDOM: [{ tag: 'u' }, { style: 'text-decoration=underline' }],
      toDOM() { return ['u', 0]; },
    },
    textColor: {
      attrs: { color: {} },
      parseDOM: [{
        tag: 'span[data-color]',
        getAttrs(dom) {
          return { color: (dom as HTMLElement).dataset.color };
        },
      }, {
        style: 'color',
        getAttrs(value) {
          return { color: value as string };
        },
      }],
      toDOM(mark) {
        return ['span', { style: `color: ${mark.attrs.color}`, 'data-color': mark.attrs.color }, 0];
      },
    },
  },
});

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Convert HeaderFooterContent to ProseMirror document
 */
function contentToPmDoc(content: HeaderFooterContent, pageIndex: number): PmNode {
  const isFirstPage = pageIndex === 0;
  const blocks = content.differentFirstPage && isFirstPage
    ? (content.firstPageBlocks ?? [])
    : content.blocks;
  
  if (blocks.length === 0) {
    // Create empty paragraph
    return headerFooterSchema.node('doc', null, [
      headerFooterSchema.node('paragraph', { alignment: 'center' }, []),
    ]);
  }
  
  const paragraphs = blocks.map(block => {
    const inlineContent: PmNode[] = [];
    
    for (const item of block.content) {
      if (item.type === 'text') {
        if (item.text) {
          const marks: Mark[] = [];
          if (item.bold) marks.push(headerFooterSchema.mark('bold'));
          if (item.italic) marks.push(headerFooterSchema.mark('italic'));
          if (item.underline) marks.push(headerFooterSchema.mark('underline'));
          if (item.color) marks.push(headerFooterSchema.mark('textColor', { color: item.color }));
          inlineContent.push(headerFooterSchema.text(item.text, marks));
        }
      } else if (item.type === 'dynamicField') {
        inlineContent.push(headerFooterSchema.node('dynamicField', { fieldType: item.fieldType }));
      } else if (item.type === 'image') {
        inlineContent.push(headerFooterSchema.node('image', {
          src: item.src,
          alt: item.alt || '',
          width: item.width || 24,
          height: item.height || 24,
        }));
      }
    }
    
    return headerFooterSchema.node('paragraph', { alignment: block.alignment || 'center' }, inlineContent);
  });
  
  return headerFooterSchema.node('doc', null, paragraphs);
}

/**
 * Convert ProseMirror document to HeaderFooterContent
 */
function pmDocToContent(doc: PmNode, originalContent: HeaderFooterContent, pageIndex: number): HeaderFooterContent {
  const blocks: HeaderFooterParagraph[] = [];
  
  doc.forEach((node) => {
    if (node.type.name === 'paragraph') {
      const content: HeaderFooterParagraph['content'] = [];
      
      node.forEach((child) => {
        if (child.isText) {
          const colorMark = child.marks.find(m => m.type.name === 'textColor');
          content.push({
            type: 'text',
            text: child.text || '',
            bold: child.marks.some(m => m.type.name === 'bold'),
            italic: child.marks.some(m => m.type.name === 'italic'),
            underline: child.marks.some(m => m.type.name === 'underline'),
            color: colorMark?.attrs.color,
          });
        } else if (child.type.name === 'dynamicField') {
          content.push({
            type: 'dynamicField',
            fieldType: child.attrs.fieldType as DynamicFieldType,
          });
        } else if (child.type.name === 'image') {
          const imgRun: HeaderFooterImageRun = {
            type: 'image',
            src: child.attrs.src,
            alt: child.attrs.alt,
            width: child.attrs.width,
            height: child.attrs.height,
          };
          content.push(imgRun);
        }
      });
      
      // If paragraph is empty, add empty text
      if (content.length === 0) {
        content.push({ type: 'text', text: '' });
      }
      
      blocks.push({
        type: 'paragraph',
        alignment: node.attrs.alignment || 'center',
        content,
      });
    }
  });
  
  const isFirstPage = pageIndex === 0;
  
  if (originalContent.differentFirstPage && isFirstPage) {
    return {
      ...originalContent,
      firstPageBlocks: blocks,
    };
  }
  
  return {
    ...originalContent,
    blocks,
  };
}

// ============================================================================
// Toolbar Button Component
// ============================================================================

interface ToolbarButtonProps {
  title: string;
  onClick: () => void;
  active?: boolean;
  children: React.ReactNode;
}

const ToolbarButton = memo(function ToolbarButton({ title, onClick, active, children }: ToolbarButtonProps) {
  return (
    <button
      type="button"
      title={title}
      onClick={(e) => {
        e.preventDefault();
        e.stopPropagation();
        onClick();
      }}
      style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        width: 26,
        height: 26,
        background: active ? '#e8f0fe' : 'transparent',
        border: 'none',
        borderRadius: 4,
        cursor: 'pointer',
        color: active ? '#1967d2' : '#5f6368',
        transition: 'all 0.15s',
        flexShrink: 0,
      }}
      onMouseEnter={(e) => {
        if (!active) e.currentTarget.style.background = '#f1f3f4';
      }}
      onMouseLeave={(e) => {
        if (!active) e.currentTarget.style.background = 'transparent';
      }}
    >
      {children}
    </button>
  );
});

// ============================================================================
// Color Dropdown Component (uses full ColorPicker)
// ============================================================================

interface ColorDropdownProps {
  onColorSelect: (color: string) => void;
  currentColor?: string;
}

const ColorDropdown = memo(function ColorDropdown({ onColorSelect, currentColor }: ColorDropdownProps) {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);
  
  return (
    <div ref={dropdownRef} style={{ position: 'relative', flexShrink: 0 }}>
      <button
        type="button"
        title="Text Color"
        onClick={(e) => {
          e.preventDefault();
          e.stopPropagation();
          setIsOpen(!isOpen);
        }}
        style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          width: 26,
          height: 26,
          background: 'transparent',
          border: 'none',
          borderRadius: 4,
          cursor: 'pointer',
          color: '#5f6368',
        }}
        onMouseEnter={(e) => {
          e.currentTarget.style.background = '#f1f3f4';
        }}
        onMouseLeave={(e) => {
          e.currentTarget.style.background = 'transparent';
        }}
      >
        <Palette size={14} />
      </button>
      
      {isOpen && (
        <ColorPicker
          currentColor={currentColor}
          onColorSelect={(color) => {
            onColorSelect(color);
            setIsOpen(false);
          }}
          onClose={() => setIsOpen(false)}
        />
      )}
    </div>
  );
});

// ============================================================================
// Image Upload Dialog for Header/Footer
// ============================================================================

interface HFImageDialogProps {
  isOpen: boolean;
  onClose: () => void;
  onInsert: (src: string, width: number, height: number) => void;
}

/**
 * Convert a File to base64 data URI
 */
function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      if (typeof reader.result === 'string') {
        resolve(reader.result);
      } else {
        reject(new Error('Failed to convert file to base64'));
      }
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

const HFImageDialog = memo(function HFImageDialog({ isOpen, onClose, onInsert }: HFImageDialogProps) {
  const [preview, setPreview] = useState<string | null>(null);
  const [width, setWidth] = useState(32);
  const [height, setHeight] = useState(32);
  const [isDragOver, setIsDragOver] = useState(false);
  const [error, setError] = useState('');
  const [activeTab, setActiveTab] = useState<'upload' | 'url'>('upload');
  const [url, setUrl] = useState('');
  
  const fileInputRef = useRef<HTMLInputElement>(null);
  const dialogRef = useRef<HTMLDivElement>(null);
  
  // Reset when dialog opens
  useEffect(() => {
    if (isOpen) {
      setPreview(null);
      setWidth(32);
      setHeight(32);
      setError('');
      setActiveTab('upload');
      setUrl('');
    }
  }, [isOpen]);
  
  // Handle click outside
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
  
  // Handle file
  const handleFile = useCallback(async (file: File) => {
    if (!file.type.startsWith('image/')) {
      setError('Please select an image file');
      return;
    }
    if (file.size > 5 * 1024 * 1024) {
      setError('Image file is too large (max 5MB for headers/footers)');
      return;
    }
    setError('');
    try {
      const base64 = await fileToBase64(file);
      setPreview(base64);
      
      // Get image dimensions
      const img = new window.Image();
      img.onload = () => {
        // Scale to max 32px height for headers/footers
        const scale = Math.min(1, 32 / img.height);
        setWidth(Math.round(img.width * scale));
        setHeight(Math.round(img.height * scale));
      };
      img.src = base64;
    } catch {
      setError('Failed to read image file');
    }
  }, []);
  
  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file);
  }, [handleFile]);
  
  const handleUrlSubmit = useCallback(() => {
    if (!url.trim()) return;
    let normalizedUrl = url.trim();
    if (!normalizedUrl.startsWith('http://') && !normalizedUrl.startsWith('https://') && !normalizedUrl.startsWith('data:')) {
      normalizedUrl = 'https://' + normalizedUrl;
    }
    setPreview(normalizedUrl);
    // Try to load to get dimensions
    const img = new window.Image();
    img.onload = () => {
      const scale = Math.min(1, 32 / img.height);
      setWidth(Math.round(img.width * scale));
      setHeight(Math.round(img.height * scale));
    };
    img.onerror = () => {
      setError('Could not load image from URL');
    };
    img.src = normalizedUrl;
  }, [url]);
  
  const handleInsert = useCallback(() => {
    if (preview) {
      onInsert(preview, width, height);
      onClose();
    }
  }, [preview, width, height, onInsert, onClose]);
  
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
          borderRadius: 8,
          boxShadow: '0 4px 24px rgba(0, 0, 0, 0.2)',
          width: 400,
          maxWidth: '90vw',
        }}
      >
        <h3 style={{ margin: 0, padding: '16px 20px 0', fontSize: 16, fontWeight: 500, color: '#202124' }}>
          Insert Image
        </h3>
        
        {/* Tabs */}
        <div style={{ display: 'flex', borderBottom: '1px solid #e8eaed', marginTop: 8 }}>
          <button
            type="button"
            onClick={() => setActiveTab('upload')}
            style={{
              padding: '10px 16px',
              fontSize: 13,
              fontWeight: 500,
              color: activeTab === 'upload' ? '#1a73e8' : '#5f6368',
              background: 'none',
              border: 'none',
              borderBottom: activeTab === 'upload' ? '2px solid #1a73e8' : '2px solid transparent',
              cursor: 'pointer',
            }}
          >
            Upload
          </button>
          <button
            type="button"
            onClick={() => setActiveTab('url')}
            style={{
              padding: '10px 16px',
              fontSize: 13,
              fontWeight: 500,
              color: activeTab === 'url' ? '#1a73e8' : '#5f6368',
              background: 'none',
              border: 'none',
              borderBottom: activeTab === 'url' ? '2px solid #1a73e8' : '2px solid transparent',
              cursor: 'pointer',
            }}
          >
            By URL
          </button>
        </div>
        
        <div style={{ padding: '16px 20px 20px' }}>
          {activeTab === 'upload' && (
            <div
              onClick={() => fileInputRef.current?.click()}
              onDragOver={(e) => { e.preventDefault(); setIsDragOver(true); }}
              onDragLeave={() => setIsDragOver(false)}
              onDrop={handleDrop}
              style={{
                border: `2px dashed ${isDragOver ? '#1a73e8' : '#dadce0'}`,
                borderRadius: 8,
                padding: preview ? 12 : 32,
                textAlign: 'center',
                cursor: 'pointer',
                backgroundColor: isDragOver ? 'rgba(26, 115, 232, 0.05)' : '#fafafa',
              }}
            >
              {preview ? (
                <div>
                  <img src={preview} alt="Preview" style={{ maxWidth: '100%', maxHeight: 80, objectFit: 'contain' }} />
                  <p style={{ margin: '8px 0 0', fontSize: 12, color: '#5f6368' }}>Click or drag to replace</p>
                </div>
              ) : (
                <>
                  <Upload size={32} style={{ color: '#9aa0a6', marginBottom: 8 }} />
                  <p style={{ margin: 0, fontSize: 13, color: '#202124' }}>Drag and drop or click to browse</p>
                  <p style={{ margin: '4px 0 0', fontSize: 11, color: '#9aa0a6' }}>PNG, JPG, GIF up to 5MB</p>
                </>
              )}
            </div>
          )}
          
          {activeTab === 'url' && (
            <div>
              <div style={{ display: 'flex', gap: 8 }}>
                <input
                  type="text"
                  value={url}
                  onChange={(e) => setUrl(e.target.value)}
                  placeholder="https://example.com/logo.png"
                  style={{
                    flex: 1,
                    padding: '8px 12px',
                    fontSize: 13,
                    border: '1px solid #dadce0',
                    borderRadius: 4,
                    outline: 'none',
                  }}
                  onKeyDown={(e) => { if (e.key === 'Enter') handleUrlSubmit(); }}
                />
                <button
                  type="button"
                  onClick={handleUrlSubmit}
                  style={{
                    padding: '8px 16px',
                    fontSize: 13,
                    background: '#f1f3f4',
                    border: '1px solid #dadce0',
                    borderRadius: 4,
                    cursor: 'pointer',
                  }}
                >
                  Load
                </button>
              </div>
              {preview && (
                <div style={{ marginTop: 12, textAlign: 'center' }}>
                  <img src={preview} alt="Preview" style={{ maxWidth: '100%', maxHeight: 80, objectFit: 'contain' }} />
                </div>
              )}
            </div>
          )}
          
          <input
            ref={fileInputRef}
            type="file"
            accept="image/*"
            onChange={(e) => { const file = e.target.files?.[0]; if (file) handleFile(file); }}
            style={{ display: 'none' }}
          />
          
          {error && <p style={{ margin: '12px 0 0', fontSize: 12, color: '#d93025' }}>{error}</p>}
          
          {/* Size controls */}
          {preview && (
            <div style={{ display: 'flex', gap: 12, marginTop: 16 }}>
              <div style={{ flex: 1 }}>
                <label style={{ display: 'block', fontSize: 12, color: '#5f6368', marginBottom: 4 }}>Width</label>
                <input
                  type="number"
                  value={width}
                  onChange={(e) => setWidth(parseInt(e.target.value) || 32)}
                  min={8}
                  max={200}
                  style={{
                    width: '100%',
                    padding: '6px 10px',
                    fontSize: 13,
                    border: '1px solid #dadce0',
                    borderRadius: 4,
                  }}
                />
              </div>
              <div style={{ flex: 1 }}>
                <label style={{ display: 'block', fontSize: 12, color: '#5f6368', marginBottom: 4 }}>Height</label>
                <input
                  type="number"
                  value={height}
                  onChange={(e) => setHeight(parseInt(e.target.value) || 32)}
                  min={8}
                  max={64}
                  style={{
                    width: '100%',
                    padding: '6px 10px',
                    fontSize: 13,
                    border: '1px solid #dadce0',
                    borderRadius: 4,
                  }}
                />
              </div>
            </div>
          )}
          
          {/* Buttons */}
          <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 20 }}>
            <button
              type="button"
              onClick={onClose}
              style={{
                padding: '8px 16px',
                fontSize: 13,
                color: '#5f6368',
                background: 'transparent',
                border: '1px solid #dadce0',
                borderRadius: 4,
                cursor: 'pointer',
              }}
            >
              Cancel
            </button>
            <button
              type="button"
              onClick={handleInsert}
              disabled={!preview}
              style={{
                padding: '8px 20px',
                fontSize: 13,
                color: 'white',
                background: preview ? '#1a73e8' : '#9aa0a6',
                border: 'none',
                borderRadius: 4,
                cursor: preview ? 'pointer' : 'not-allowed',
              }}
            >
              Insert
            </button>
          </div>
        </div>
      </div>
    </div>
  );
});

// ============================================================================
// Fields Dropdown Component
// ============================================================================

interface FieldsDropdownProps {
  onInsertField: (fieldType: DynamicFieldType) => void;
}

const FIELD_OPTIONS: { type: DynamicFieldType; label: string; icon: React.ReactNode }[] = [
  { type: 'pageNumber', label: 'Page Number', icon: <Hash size={14} /> },
  { type: 'totalPages', label: 'Total Pages', icon: <Layers size={14} /> },
  { type: 'date', label: 'Current Date', icon: <CalendarDays size={14} /> },
  { type: 'time', label: 'Current Time', icon: <Clock size={14} /> },
  { type: 'title', label: 'Document Title', icon: <FileText size={14} /> },
];

const FieldsDropdown = memo(function FieldsDropdown({ onInsertField }: FieldsDropdownProps) {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef<HTMLDivElement>(null);
  
  useEffect(() => {
    if (!isOpen) return;
    const handleClickOutside = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [isOpen]);
  
  return (
    <div ref={dropdownRef} style={{ position: 'relative', flexShrink: 0 }}>
      <button
        type="button"
        title="Insert Field"
        onClick={(e) => {
          e.preventDefault();
          e.stopPropagation();
          setIsOpen(!isOpen);
        }}
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: 3,
          padding: '3px 6px',
          fontSize: 11,
          background: '#f8f9fa',
          border: '1px solid #dadce0',
          borderRadius: 4,
          cursor: 'pointer',
          color: '#5f6368',
          whiteSpace: 'nowrap',
        }}
        onMouseEnter={(e) => {
          e.currentTarget.style.background = '#e8eaed';
        }}
        onMouseLeave={(e) => {
          e.currentTarget.style.background = '#f8f9fa';
        }}
      >
        <Hash size={11} />
        <span>Field</span>
        <ChevronDown size={9} />
      </button>
      
      {isOpen && (
        <div
          style={{
            position: 'absolute',
            top: '100%',
            left: 0,
            marginTop: 4,
            background: 'white',
            borderRadius: 8,
            boxShadow: '0 4px 16px rgba(0, 0, 0, 0.15)',
            padding: 4,
            zIndex: 1003,
            minWidth: 160,
          }}
        >
          {FIELD_OPTIONS.map((option) => (
            <button
              key={option.type}
              type="button"
              onClick={(e) => {
                e.preventDefault();
                e.stopPropagation();
                onInsertField(option.type);
                setIsOpen(false);
              }}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 8,
                width: '100%',
                padding: '8px 12px',
                background: 'transparent',
                border: 'none',
                borderRadius: 4,
                cursor: 'pointer',
                color: '#202124',
                fontSize: 12,
                textAlign: 'left',
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.background = '#f1f3f4';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.background = 'transparent';
              }}
            >
              <span style={{ color: '#5f6368' }}>{option.icon}</span>
              {option.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
});

// ============================================================================
// Main Component
// ============================================================================

export const HeaderFooterEditor = memo(function HeaderFooterEditor({
  type,
  content,
  onChange,
  isEditing,
  onClose,
  pageIndex,
  totalPages,
  documentTitle = 'Untitled',
  contentWidth,
  scale,
  position,
}: HeaderFooterEditorProps) {
  const editorContainerRef = useRef<HTMLDivElement>(null);
  const editorViewRef = useRef<EditorView | null>(null);
  const [differentFirstPage, setDifferentFirstPage] = useState(content.differentFirstPage ?? false);
  const [showImageDialog, setShowImageDialog] = useState(false);
  const [currentAlignment, setCurrentAlignment] = useState<'left' | 'center' | 'right'>('center');
  
  // Get the current paragraph alignment from editor state
  const getAlignmentFromState = useCallback((state: EditorState): 'left' | 'center' | 'right' => {
    const { $from } = state.selection;
    // Walk up from the selection position to find the paragraph
    for (let depth = $from.depth; depth >= 0; depth--) {
      const node = $from.node(depth);
      if (node.type.name === 'paragraph') {
        return (node.attrs.alignment as 'left' | 'center' | 'right') || 'center';
      }
    }
    // Fallback: check if doc has a first paragraph
    if (state.doc.firstChild?.type.name === 'paragraph') {
      return (state.doc.firstChild.attrs.alignment as 'left' | 'center' | 'right') || 'center';
    }
    return 'center';
  }, []);
  
  // Use ref to avoid stale closure in dispatchTransaction
  const getAlignmentFromStateRef = useRef(getAlignmentFromState);
  getAlignmentFromStateRef.current = getAlignmentFromState;
  
  // Save changes and close
  const handleSave = useCallback(() => {
    if (editorViewRef.current) {
      const newContent = pmDocToContent(editorViewRef.current.state.doc, content, pageIndex);
      onChange(newContent);
    }
    onClose();
  }, [content, pageIndex, onChange, onClose]);
  
  // Insert a dynamic field at cursor
  const insertField = useCallback((fieldType: DynamicFieldType) => {
    const view = editorViewRef.current;
    if (!view) return;
    
    const { state } = view;
    const node = headerFooterSchema.node('dynamicField', { fieldType });
    const tr = state.tr.replaceSelectionWith(node);
    view.dispatch(tr);
    view.focus();
  }, []);
  
  // Handle different first page toggle
  const handleDifferentFirstPageChange = useCallback((checked: boolean) => {
    setDifferentFirstPage(checked);
    onChange({
      ...content,
      differentFirstPage: checked,
      firstPageBlocks: checked ? [] : undefined,
    });
  }, [content, onChange]);
  
  // Toggle bold
  const toggleBold = useCallback(() => {
    const view = editorViewRef.current;
    if (!view) return;
    toggleMark(headerFooterSchema.marks.bold)(view.state, view.dispatch);
    view.focus();
  }, []);
  
  // Toggle italic
  const toggleItalic = useCallback(() => {
    const view = editorViewRef.current;
    if (!view) return;
    toggleMark(headerFooterSchema.marks.italic)(view.state, view.dispatch);
    view.focus();
  }, []);
  
  // Toggle underline
  const toggleUnderline = useCallback(() => {
    const view = editorViewRef.current;
    if (!view) return;
    toggleMark(headerFooterSchema.marks.underline)(view.state, view.dispatch);
    view.focus();
  }, []);
  
  // Set text color
  const setTextColor = useCallback((color: string) => {
    const view = editorViewRef.current;
    if (!view) return;
    const { from, to } = view.state.selection;
    if (from === to) {
      // No selection, just store for next input
      view.focus();
      return;
    }
    const tr = view.state.tr.addMark(from, to, headerFooterSchema.marks.textColor.create({ color }));
    view.dispatch(tr);
    view.focus();
  }, []);
  
  // Set paragraph alignment
  const setAlignment = useCallback((alignment: 'left' | 'center' | 'right') => {
    const view = editorViewRef.current;
    if (!view) return;
    const { from, to } = view.state.selection;
    let tr = view.state.tr;
    view.state.doc.nodesBetween(from, to, (node, pos) => {
      if (node.type.name === 'paragraph') {
        tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, alignment });
      }
    });
    view.dispatch(tr);
    // Update state immediately for responsive UI
    setCurrentAlignment(alignment);
    view.focus();
  }, []);
  
  // Insert image from dialog
  const insertImage = useCallback((src: string, imgWidth: number, imgHeight: number) => {
    const view = editorViewRef.current;
    if (!view) return;
    
    const node = headerFooterSchema.node('image', { src, width: imgWidth, height: imgHeight });
    const tr = view.state.tr.replaceSelectionWith(node);
    view.dispatch(tr);
    view.focus();
  }, []);
  
  // Initialize ProseMirror editor
  useEffect(() => {
    if (!isEditing || !editorContainerRef.current) return;
    
    const doc = contentToPmDoc(content, pageIndex);
    
    const state = EditorState.create({
      doc,
      schema: headerFooterSchema,
      plugins: [
        history(),
        keymap({
          'Mod-z': undo,
          'Mod-y': redo,
          'Mod-Shift-z': redo,
          'Mod-b': toggleMark(headerFooterSchema.marks.bold),
          'Mod-i': toggleMark(headerFooterSchema.marks.italic),
          'Mod-u': toggleMark(headerFooterSchema.marks.underline),
        }),
        keymap(baseKeymap),
      ],
    });
    
    const view = new EditorView(editorContainerRef.current, {
      state,
      dispatchTransaction(tr: Transaction) {
        const newState = view.state.apply(tr);
        view.updateState(newState);
        // Update alignment state when selection or doc changes
        // Use ref to get fresh function reference
        const alignment = getAlignmentFromStateRef.current(newState);
        setCurrentAlignment(alignment);
      },
      attributes: {
        class: 'header-footer-pm-editor',
        style: `
          outline: none;
          min-height: 20px;
          font-size: ${12 * scale}px;
          line-height: 1.5;
          text-align: center;
        `,
      },
    });
    
    // Set initial alignment from content
    setCurrentAlignment(getAlignmentFromStateRef.current(state));
    
    editorViewRef.current = view;
    
    // Focus with a small delay to ensure it's mounted
    setTimeout(() => {
      view.focus();
      // Move cursor to end
      const endPos = view.state.doc.content.size - 1;
      if (endPos > 0) {
        try {
          const tr = view.state.tr.setSelection(
            TextSelection.near(view.state.doc.resolve(endPos))
          );
          view.dispatch(tr);
        } catch {
          // Selection might fail if doc is empty, that's fine
        }
      }
    }, 50);
    
    return () => {
      view.destroy();
      editorViewRef.current = null;
    };
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isEditing, content, pageIndex, scale]);
  
  // Handle escape key
  useEffect(() => {
    if (!isEditing) return;
    
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        handleSave();
      }
    };
    
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isEditing, handleSave]);
  
  if (!isEditing) return null;
  
  const isFirstPage = pageIndex === 0;
  
  return (
    <>
      {/* Overlay to dim main content */}
      <div
        className="hf-overlay"
        onClick={handleSave}
        style={{
          position: 'fixed',
          inset: 0,
          background: 'rgba(0, 0, 0, 0.15)',
          zIndex: 999,
        }}
      />
      
      {/* Floating toolbar for header/footer */}
      <div
        className="hf-toolbar"
        style={{
          position: 'fixed',
          top: 60,
          left: '50%',
          transform: 'translateX(-50%)',
          background: 'white',
          borderRadius: 8,
          boxShadow: '0 4px 16px rgba(0, 0, 0, 0.15)',
          padding: '6px 12px',
          display: 'flex',
          alignItems: 'center',
          gap: 6,
          zIndex: 1002,
          whiteSpace: 'nowrap',
        }}
      >
        {/* Mode indicator */}
        <div style={{
          display: 'flex',
          alignItems: 'center',
          gap: 4,
          padding: '3px 6px',
          background: '#e8f0fe',
          borderRadius: 4,
          color: '#1967d2',
          fontSize: 11,
          fontWeight: 500,
        }}>
          <FileText size={12} />
          {type === 'header' ? 'Header' : 'Footer'}
        </div>
        
        {/* Separator */}
        <div style={{ width: 1, height: 20, background: '#e8eaed', flexShrink: 0 }} />
        
        {/* Different first page checkbox */}
        <label style={{
          display: 'flex',
          alignItems: 'center',
          gap: 4,
          fontSize: 11,
          color: '#5f6368',
          cursor: 'pointer',
          userSelect: 'none',
        }}>
          <input
            type="checkbox"
            checked={differentFirstPage}
            onChange={(e) => handleDifferentFirstPageChange(e.target.checked)}
            style={{ margin: 0, width: 14, height: 14 }}
          />
          Different first page
        </label>
        
        {/* Info badge */}
        {differentFirstPage && (
          <div style={{
            fontSize: 10,
            color: '#1967d2',
            background: '#e8f0fe',
            padding: '2px 6px',
            borderRadius: 3,
          }}>
            {isFirstPage ? 'First page' : 'Other pages'}
          </div>
        )}
        
        {/* Separator */}
        <div style={{ width: 1, height: 20, background: '#e8eaed', flexShrink: 0 }} />
        
        {/* Text formatting */}
        <div style={{ display: 'flex', gap: 1, alignItems: 'center' }}>
          <ToolbarButton title="Bold (Ctrl+B)" onClick={toggleBold}>
            <Bold size={14} />
          </ToolbarButton>
          <ToolbarButton title="Italic (Ctrl+I)" onClick={toggleItalic}>
            <Italic size={14} />
          </ToolbarButton>
          <ToolbarButton title="Underline (Ctrl+U)" onClick={toggleUnderline}>
            <Underline size={14} />
          </ToolbarButton>
        </div>
        
        {/* Separator */}
        <div style={{ width: 1, height: 20, background: '#e8eaed', flexShrink: 0 }} />
        
        {/* Text color dropdown */}
        <ColorDropdown onColorSelect={setTextColor} />
        
        {/* Separator */}
        <div style={{ width: 1, height: 20, background: '#e8eaed', flexShrink: 0 }} />
        
        {/* Alignment */}
        <div style={{ display: 'flex', gap: 1, alignItems: 'center' }}>
          <ToolbarButton title="Align Left" onClick={() => setAlignment('left')} active={currentAlignment === 'left'}>
            <AlignLeft size={14} />
          </ToolbarButton>
          <ToolbarButton title="Align Center" onClick={() => setAlignment('center')} active={currentAlignment === 'center'}>
            <AlignCenter size={14} />
          </ToolbarButton>
          <ToolbarButton title="Align Right" onClick={() => setAlignment('right')} active={currentAlignment === 'right'}>
            <AlignRight size={14} />
          </ToolbarButton>
        </div>
        
        {/* Separator */}
        <div style={{ width: 1, height: 20, background: '#e8eaed', flexShrink: 0 }} />
        
        {/* Insert image */}
        <ToolbarButton title="Insert Image" onClick={() => setShowImageDialog(true)}>
          <Image size={14} />
        </ToolbarButton>
        
        {/* Separator */}
        <div style={{ width: 1, height: 20, background: '#e8eaed', flexShrink: 0 }} />
        
        {/* Insert dynamic fields dropdown */}
        <FieldsDropdown onInsertField={insertField} />
        
        {/* Separator */}
        <div style={{ width: 1, height: 20, background: '#e8eaed', flexShrink: 0 }} />
        
        {/* Close button */}
        <button
          type="button"
          onClick={handleSave}
          style={{
            display: 'flex',
            alignItems: 'center',
            padding: '4px 10px',
            background: '#1a73e8',
            border: 'none',
            borderRadius: 4,
            cursor: 'pointer',
            color: 'white',
            fontSize: 11,
            fontWeight: 500,
            flexShrink: 0,
          }}
        >
          Done
        </button>
      </div>
      
      {/* Editor positioned at the actual header/footer location */}
      <div
        className="hf-editor-container"
        style={{
          position: 'fixed',
          top: type === 'header' 
            ? (position.pageY + (position.top ?? 48) * scale)
            : undefined,
          bottom: type === 'footer'
            ? `calc(100vh - ${position.pageY + (position.bottom ?? 48) * scale}px)`
            : undefined,
          left: position.left,
          width: contentWidth,
          zIndex: 1001,
          background: 'white',
          boxShadow: '0 0 0 2px #1a73e8',
          borderRadius: 4,
          padding: '8px 12px',
          minHeight: 32,
        }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* ProseMirror editor container */}
        <div 
          ref={editorContainerRef}
          className="hf-pm-container"
        />
        
        {/* Preview info */}
        <div style={{
          marginTop: 8,
          paddingTop: 8,
          borderTop: '1px solid #e8eaed',
          fontSize: 10,
          color: '#9aa0a6',
          textAlign: 'center',
        }}>
          Preview: Page {pageIndex + 1} of {totalPages} • {documentTitle}
        </div>
      </div>
      
      {/* Styles */}
      <style>{`
        .hf-pm-container .ProseMirror {
          outline: none !important;
          min-height: 20px;
        }
        .hf-pm-container .ProseMirror p {
          margin: 0;
        }
        .hf-pm-container .ProseMirror:focus {
          outline: none !important;
        }
        .hf-pm-container .dynamic-field {
          user-select: all;
        }
      `}</style>
      
      {/* Image Upload Dialog */}
      <HFImageDialog
        isOpen={showImageDialog}
        onClose={() => setShowImageDialog(false)}
        onInsert={insertImage}
      />
    </>
  );
});

export default HeaderFooterEditor;
