import React, { memo, useState, useCallback, useMemo, useEffect, useRef } from 'react';
import { EditorView } from 'prosemirror-view';
import { TextSelection, type Command } from 'prosemirror-state';
import { useDocument, useHistory } from '../context/DocumentContext';
import { 
  docsSchema, 
  createCommands,
  undo as pmUndo,
  redo as pmRedo,
} from '@b32nio/spindle-docs-core';
import type { HeadingLevel } from '@b32nio/spindle-docs-core';
import type { ActiveMarks } from './ProseMirrorEditor';
import type { CellSelection } from '../core';
import { ResponsiveToolbar, useToolbarMenu, MenuRow } from '@b32nio/spindle-shared/react';
import { LinkDialog } from './LinkDialog';
import { ImageDialog } from './ImageDialog';
import { TableSizePicker } from './TableSizePicker';

// Lucide Icons
import {
  Undo2,
  Redo2,
  Bold,
  Italic,
  Underline,
  Strikethrough,
  Superscript,
  Subscript,
  Link,
  Image,
  Table,
  Minus,
  Plus,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  List,
  ListOrdered,
  IndentDecrease,
  IndentIncrease,
  ChevronDown,
  Palette,
  Highlighter,
  SeparatorHorizontal,
  FileDown,
  MessageSquarePlus,
  Type,
  ALargeSmall,
  Settings2,
} from 'lucide-react';

// ============================================================================
// Toolbar Styles
// ============================================================================

const styles = {
  // Main floating toolbar container
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: '2px',
    padding: '6px 12px',
    background: 'linear-gradient(135deg, rgba(255,255,255,0.95) 0%, rgba(248,250,252,0.95) 100%)',
    backdropFilter: 'blur(20px) saturate(180%)',
    WebkitBackdropFilter: 'blur(20px) saturate(180%)',
    borderRadius: '16px',
    boxShadow: `
      0 4px 6px -1px rgba(0, 0, 0, 0.05),
      0 10px 15px -3px rgba(0, 0, 0, 0.08),
      0 20px 25px -5px rgba(0, 0, 0, 0.05),
      inset 0 1px 0 rgba(255, 255, 255, 0.9),
      0 0 0 1px rgba(0, 0, 0, 0.04)
    `,
    fontFamily: '"Inter", "SF Pro Display", -apple-system, BlinkMacSystemFont, sans-serif',
    position: 'relative' as const,
    zIndex: 100,
    margin: '8px 0',
    width: '100%',
    boxSizing: 'border-box' as const,
  } as React.CSSProperties,
  
  // Button base style
  button: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 32,
    height: 32,
    border: 'none',
    borderRadius: '8px',
    backgroundColor: 'transparent',
    color: '#64748b',
    cursor: 'pointer',
    transition: 'all 0.15s cubic-bezier(0.4, 0, 0.2, 1)',
    position: 'relative' as const,
  } as React.CSSProperties,
  
  buttonHover: {
    backgroundColor: 'rgba(99, 102, 241, 0.08)',
    color: '#6366f1',
    transform: 'translateY(-1px)',
  },
  
  buttonActive: {
    backgroundColor: 'rgba(99, 102, 241, 0.15)',
    color: '#6366f1',
    boxShadow: 'inset 0 1px 2px rgba(99, 102, 241, 0.15)',
  },
  
  buttonDisabled: {
    opacity: 0.4,
    cursor: 'not-allowed',
    pointerEvents: 'none' as const,
  },
  
  // Divider
  divider: {
    width: 1,
    height: 24,
    background: 'linear-gradient(180deg, transparent 0%, rgba(0,0,0,0.08) 50%, transparent 100%)',
    margin: '0 6px',
    flexShrink: 0,
  } as React.CSSProperties,
  
  // Dropdown button
  dropdownButton: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '4px',
    height: 32,
    padding: '0 8px 0 12px',
    border: 'none',
    borderRadius: '8px',
    backgroundColor: 'transparent',
    color: '#475569',
    cursor: 'pointer',
    fontSize: '13px',
    fontWeight: 500,
    fontFamily: '"Inter", sans-serif',
    transition: 'all 0.15s cubic-bezier(0.4, 0, 0.2, 1)',
    whiteSpace: 'nowrap' as const,
  } as React.CSSProperties,
  
  // Dropdown menu
  dropdown: {
    position: 'absolute' as const,
    top: 'calc(100% + 8px)',
    left: 0,
    backgroundColor: 'rgba(255, 255, 255, 0.98)',
    backdropFilter: 'blur(20px)',
    borderRadius: '12px',
    boxShadow: `
      0 4px 6px -1px rgba(0, 0, 0, 0.07),
      0 10px 15px -3px rgba(0, 0, 0, 0.1),
      0 20px 25px -5px rgba(0, 0, 0, 0.08),
      0 0 0 1px rgba(0, 0, 0, 0.05)
    `,
    zIndex: 10000,
    padding: '6px',
    minWidth: '120px',
    animation: 'dropdownFadeIn 0.15s ease-out',
  } as React.CSSProperties,
  
  // Dropdown item
  dropdownItem: {
    padding: '8px 12px',
    borderRadius: '6px',
    cursor: 'pointer',
    fontSize: '13px',
    color: '#334155',
    transition: 'all 0.1s ease',
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  } as React.CSSProperties,
  
  dropdownItemHover: {
    backgroundColor: 'rgba(99, 102, 241, 0.08)',
    color: '#6366f1',
  },
  
  dropdownItemActive: {
    backgroundColor: 'rgba(99, 102, 241, 0.12)',
    color: '#6366f1',
    fontWeight: 500,
  },
  
  // Font size input
  fontSizeInput: {
    width: 40,
    height: 28,
    border: '1px solid rgba(0, 0, 0, 0.1)',
    borderRadius: '6px',
    textAlign: 'center' as const,
    fontSize: '13px',
    fontWeight: 500,
    color: '#334155',
    backgroundColor: 'rgba(255, 255, 255, 0.8)',
    fontFamily: '"Inter", sans-serif',
    cursor: 'pointer',
    transition: 'all 0.15s ease',
  } as React.CSSProperties,
  
  // Tooltip
  tooltip: {
    position: 'absolute' as const,
    top: 'calc(100% + 8px)',
    left: '50%',
    transform: 'translateX(-50%) translateY(4px)',
    padding: '6px 10px',
    backgroundColor: '#1e293b',
    color: '#fff',
    fontSize: '11px',
    fontWeight: 500,
    borderRadius: '6px',
    whiteSpace: 'nowrap' as const,
    pointerEvents: 'none' as const,
    opacity: 0,
    transition: 'opacity 0.2s ease, transform 0.2s ease',
    zIndex: 10001,
  } as React.CSSProperties,
};

// CSS keyframes injection
const injectStyles = () => {
  if (typeof document !== 'undefined' && !document.getElementById('toolbar-animations')) {
    const style = document.createElement('style');
    style.id = 'toolbar-animations';
    style.textContent = `
      @keyframes dropdownFadeIn {
        from { opacity: 0; transform: translateY(-4px); }
        to { opacity: 1; transform: translateY(0); }
      }
      .toolbar-btn .toolbar-tooltip {
        transition: opacity 0.2s ease 0.4s, transform 0.2s ease 0.4s;
      }
      .toolbar-btn:hover .toolbar-tooltip {
        opacity: 1 !important;
        transform: translateX(-50%) translateY(0) !important;
        transition: opacity 0.15s ease 0.4s, transform 0.15s ease 0.4s;
      }
      .docs-toolbar:hover .toolbar-btn:hover .toolbar-tooltip {
        transition-delay: 0s, 0s;
      }
      .docs-toolbar:hover .toolbar-btn .toolbar-tooltip {
        transition-delay: 0s, 0s;
      }
      .color-swatch:hover {
        transform: scale(1.15);
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
      }
    `;
    document.head.appendChild(style);
  }
};

// ============================================================================
// Toolbar Components
// ============================================================================

interface ToolbarProps {
  onPageSetup?: () => void;
  editorView?: EditorView | null;
  activeMarks?: ActiveMarks;
  selectedCell?: CellSelection | null;
  /** Start a comment on the current selection. */
  onAddComment?: () => void;
}

interface ToolbarButtonProps {
  onClick?: () => void;
  active?: boolean;
  disabled?: boolean;
  tooltip?: string;
  children: React.ReactNode;
}

const ToolbarButton = memo(function ToolbarButton({
  onClick,
  active,
  disabled,
  tooltip,
  children,
}: ToolbarButtonProps) {
  const [isHovered, setIsHovered] = useState(false);
  
  return (
    <button
      className="toolbar-btn"
      onClick={onClick}
      disabled={disabled}
      onMouseEnter={() => setIsHovered(true)}
      onMouseLeave={() => setIsHovered(false)}
      style={{
        ...styles.button,
        ...(isHovered && !active && !disabled ? styles.buttonHover : {}),
        ...(active ? styles.buttonActive : {}),
        ...(disabled ? styles.buttonDisabled : {}),
      }}
    >
      {children}
      {tooltip && (
        <span className="toolbar-tooltip" style={styles.tooltip}>
          {tooltip}
        </span>
      )}
    </button>
  );
});

const Divider = memo(function Divider() {
  return <div style={styles.divider} />;
});

const docHex = (c: string) => /^#[0-9a-fA-F]{6}$/.test(c);
const DOC_COLORS = [
  '#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#ffffff',
  '#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#9900ff',
  '#e6b8af', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#d9d2e9', '#ead1dc',
  '#cc4125', '#e69138', '#f1c232', '#6aa84f', '#45818e', '#3c78d8', '#674ea7', '#a64d79',
];

// A dropdown control that works inline (button + popover) and inside the
// overflow surface (menu row + modal). `triggerLabel` set → text trigger.
const OverflowDropdown = memo(function OverflowDropdown({ triggerIcon, triggerLabel, menuLabel, hint, width, align, minWidth, children }: {
  triggerIcon?: React.ReactNode; triggerLabel?: React.ReactNode; menuLabel: string; hint?: React.ReactNode;
  width?: number; align?: 'left' | 'right'; minWidth?: number; children: React.ReactNode;
}) {
  const menu = useToolbarMenu();
  const [open, setOpen] = useState(false);
  if (menu.inMenu) {
    return <MenuRow icon={triggerIcon} label={menuLabel} hint={hint} onClick={() => menu.openModal(menuLabel, <div onClick={() => menu.closeMenu()}>{children}</div>)} />;
  }
  return (
    <div style={{ position: 'relative' }}>
      {triggerLabel !== undefined ? (
        <button onClick={() => setOpen((o) => !o)} style={{ ...styles.dropdownButton, ...(minWidth ? { minWidth } : {}), ...(open ? styles.buttonActive : {}) }}>
          {triggerIcon}
          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: 70 }}>{triggerLabel}</span>
          <ChevronDown size={14} />
        </button>
      ) : (
        <ToolbarButton onClick={() => setOpen((o) => !o)} active={open} tooltip={menuLabel}>{triggerIcon}</ToolbarButton>
      )}
      {open && (
        <div style={{ ...styles.dropdown, ...(width ? { minWidth: width } : {}), ...(align === 'right' ? { left: 'auto', right: 0 } : {}) }} onMouseLeave={() => setOpen(false)}>
          <div onClick={() => setOpen(false)}>{children}</div>
        </div>
      )}
    </div>
  );
});

const DocColorControl = memo(function DocColorControl({ icon, label, value, onApply }: {
  icon: React.ReactNode; label: string; value?: string; onApply: (c: string) => void;
}) {
  const menu = useToolbarMenu();
  const [open, setOpen] = useState(false);
  const swatches = (done: () => void) => (
    <div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(8, 1fr)', gap: 8 }}>
        {DOC_COLORS.map((c) => (
          <button key={c} title={c} onClick={() => { onApply(c); done(); }} style={{ width: '100%', aspectRatio: '1', borderRadius: 8, border: c === value ? '2px solid #2d7ff9' : '1px solid #d5d9e0', background: c, cursor: 'pointer' }} />
        ))}
      </div>
      <label style={{ display: 'flex', alignItems: 'center', gap: 10, marginTop: 14, fontSize: 14, color: '#334155', cursor: 'pointer' }}>
        <input type="color" value={value && docHex(value) ? value : '#000000'} onChange={(e) => onApply(e.target.value)} style={{ width: 40, height: 32, border: '1px solid #d5d9e0', borderRadius: 6, background: 'none', padding: 2, cursor: 'pointer' }} />
        Custom color
      </label>
    </div>
  );
  if (menu.inMenu) {
    return <MenuRow icon={icon} label={label} hint={<span style={{ width: 16, height: 16, borderRadius: 4, border: '1px solid #d5d9e0', background: value || '#000', display: 'inline-block' }} />} onClick={() => menu.openModal(label, swatches(() => menu.closeMenu()))} />;
  }
  return (
    <div style={{ position: 'relative' }}>
      <ToolbarButton onClick={() => setOpen((o) => !o)} active={open} tooltip={label}>
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2 }}>
          {icon}
          <div style={{ width: 14, height: 3, borderRadius: 1, backgroundColor: value || '#000' }} />
        </div>
      </ToolbarButton>
      {open && (
        <div style={{ ...styles.dropdown, padding: 8, minWidth: 216 }} onMouseLeave={() => setOpen(false)}>
          {swatches(() => setOpen(false))}
        </div>
      )}
    </div>
  );
});

const DocFontSizeControl = memo(function DocFontSizeControl({ value, sizes, onPick }: {
  value: number; sizes: number[]; onPick: (s: number) => void;
}) {
  const menu = useToolbarMenu();
  const [open, setOpen] = useState(false);
  const step = (dir: 'down' | 'up') => {
    const idx = sizes.indexOf(value);
    if (dir === 'down') { if (idx > 0) onPick(sizes[idx - 1]); else if (idx === -1) { const s = [...sizes].reverse().find((x) => x < value); if (s) onPick(s); } }
    else { if (idx !== -1 && idx < sizes.length - 1) onPick(sizes[idx + 1]); else if (idx === -1) { const s = sizes.find((x) => x > value); if (s) onPick(s); } }
  };
  const grid = (done: () => void) => (
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, 48px)', gap: 6, justifyContent: 'center' }}>
      {sizes.map((s) => (
        <button key={s} onClick={() => { onPick(s); done(); }} style={{ height: 40, borderRadius: 8, border: s === value ? '2px solid #2d7ff9' : '1px solid #d5d9e0', background: '#fff', cursor: 'pointer', fontSize: 14, color: '#334155' }}>{s}</button>
      ))}
    </div>
  );
  if (menu.inMenu) {
    return <MenuRow icon={<ALargeSmall size={17} />} label="Font size" hint={value} onClick={() => menu.openModal('Font size', grid(() => menu.closeMenu()))} />;
  }
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 2 }}>
      <ToolbarButton onClick={() => step('down')} tooltip="Decrease size"><Minus size={14} strokeWidth={2.5} /></ToolbarButton>
      <div style={{ position: 'relative' }}>
        <input type="text" value={value} readOnly onClick={() => setOpen((o) => !o)} style={styles.fontSizeInput} />
        {open && (
          <div style={{ ...styles.dropdown, left: '50%', transform: 'translateX(-50%)', minWidth: 60, maxHeight: 240, overflowY: 'auto' }} onMouseLeave={() => setOpen(false)}>
            {sizes.map((s) => (
              <DropdownItemStandalone key={s} onClick={() => { onPick(s); setOpen(false); }} active={value === s}>{s}</DropdownItemStandalone>
            ))}
          </div>
        )}
      </div>
      <ToolbarButton onClick={() => step('up')} tooltip="Increase size"><Plus size={14} strokeWidth={2.5} /></ToolbarButton>
    </div>
  );
});

// Minimal standalone item (module-level; the in-component DropdownItem is a closure).
const DropdownItemStandalone = memo(function DropdownItemStandalone({ onClick, active, children, style }: {
  onClick: () => void; active?: boolean; children: React.ReactNode; style?: React.CSSProperties;
}) {
  const [hover, setHover] = useState(false);
  return (
    <div onClick={onClick} onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)}
      style={{ ...styles.dropdownItem, ...(hover ? styles.dropdownItemHover : {}), ...(active ? styles.dropdownItemActive : {}), justifyContent: 'center', ...style }}>
      {children}
    </div>
  );
});

/**
 * Document editor toolbar - Modern floating design
 */
export const Toolbar = memo(function Toolbar({ 
  onPageSetup,
  onAddComment,
  editorView,
  activeMarks,
  selectedCell,
}: ToolbarProps) {
  const { zoom, setZoom } = useDocument();
  const { canUndo, canRedo, undo, redo } = useHistory();
  
  const [showLinkDialog, setShowLinkDialog] = useState(false);
  const [showImageDialog, setShowImageDialog] = useState(false);
  const [showTablePicker, setShowTablePicker] = useState(false);

  const toolbarRef = useRef<HTMLDivElement>(null);

  // Inject animation styles
  useEffect(() => {
    injectStyles();
  }, []);

  const commands = createCommands(docsSchema);
  
  const getSelectedText = useCallback((): string => {
    if (!editorView) return '';
    const { from, to } = editorView.state.selection;
    if (from === to) return '';
    return editorView.state.doc.textBetween(from, to, ' ');
  }, [editorView]);
  
  const existingLink = useMemo(() => activeMarks?.link ?? null, [activeMarks]);
  
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
        e.preventDefault();
        setShowLinkDialog(true);
      }
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, []);
  
  const runCommand = useCallback((command: Command) => {
    if (!editorView) return false;
    
    if (selectedCell) {
      const { state } = editorView;
      const from = selectedCell.pmStart + 2;
      const to = selectedCell.pmEnd - 2;
      if (from < to && from > 0 && to <= state.doc.content.size) {
        const tr = state.tr.setSelection(TextSelection.create(state.doc, from, to));
        editorView.dispatch(tr);
      }
    }
    
    const result = command(editorView.state, editorView.dispatch, editorView);
    editorView.focus();
    return result;
  }, [editorView, selectedCell]);
  
  const handleUndo = useCallback(() => {
    if (editorView) {
      pmUndo(editorView.state, editorView.dispatch);
      editorView.focus();
    } else {
      undo();
    }
  }, [editorView, undo]);
  
  const handleRedo = useCallback(() => {
    if (editorView) {
      pmRedo(editorView.state, editorView.dispatch);
      editorView.focus();
    } else {
      redo();
    }
  }, [editorView, redo]);
  
  const handleHighlightColor = useCallback((color: string) => {
    if (!editorView) return;
    
    if (selectedCell) {
      const cmd = commands.setCellBackgroundColor(selectedCell.pmStart, selectedCell.pmEnd, color);
      cmd(editorView.state, editorView.dispatch, editorView);
    } else {
      runCommand(commands.setHighlight(color));
    }
    editorView.focus();
  }, [editorView, selectedCell, commands, runCommand]);
  
  const fonts = [
    'Inter', 'Arial', 'Georgia', 'Times New Roman', 'Verdana',
    'Roboto', 'Open Sans', 'Lato', 'Montserrat', 'Poppins',
  ];
  
  const fontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
  const zoomLevels = [50, 75, 90, 100, 125, 150, 200];
  
  const headingOptions: Array<{ level: HeadingLevel | 'normal'; label: string; size: number }> = [
    { level: 'normal', label: 'Paragraph', size: 14 },
    { level: 1, label: 'Heading 1', size: 24 },
    { level: 2, label: 'Heading 2', size: 20 },
    { level: 3, label: 'Heading 3', size: 16 },
  ];
  
  const currentHeadingLabel = (() => {
    if (!activeMarks) return 'Paragraph';
    if (activeMarks.blockType === 'heading' && activeMarks.headingLevel) {
      const option = headingOptions.find(o => o.level === activeMarks.headingLevel);
      return option?.label || 'Paragraph';
    }
    return 'Paragraph';
  })();
  
  // Heading level to font size mapping (matches dom-painter.ts)
  const headingFontSizes: Record<number, number> = {
    1: 24,
    2: 20,
    3: 16,
    4: 14,
    5: 12,
    6: 11,
  };
  
  // Calculate current font size: prefer explicit textStyle fontSize, then heading size, then default
  const currentFontSize = (() => {
    // If there's an explicit fontSize in textStyle, use that (user applied font size)
    if (activeMarks?.textStyle?.fontSize) {
      return activeMarks.textStyle.fontSize;
    }
    // Otherwise, if in a heading, use the heading's inherent font size
    if (activeMarks?.blockType === 'heading' && activeMarks.headingLevel) {
      return headingFontSizes[activeMarks.headingLevel] || 16;
    }
    // Default font size
    return 11;
  })();
  
  const currentFont = activeMarks?.textStyle?.fontFamily || 'Inter';
  
  const DropdownItem = ({
    onClick,
    active,
    children,
    style: customStyle,
  }: {
    onClick: () => void;
    active?: boolean;
    children: React.ReactNode;
    style?: React.CSSProperties;
  }) => {
    const [isHovered, setIsHovered] = useState(false);
    
    return (
      <div
        onClick={onClick}
        onMouseEnter={() => setIsHovered(true)}
        onMouseLeave={() => setIsHovered(false)}
        style={{
          ...styles.dropdownItem,
          ...(isHovered ? styles.dropdownItemHover : {}),
          ...(active ? styles.dropdownItemActive : {}),
          ...customStyle,
        }}
      >
        {children}
      </div>
    );
  };
  
  return (
    <div 
      ref={toolbarRef}
      className="docs-toolbar"
      style={styles.toolbar}
    >
      <ResponsiveToolbar gap={2}>
      {/* Undo/Redo */}
      <ToolbarButton onClick={handleUndo} disabled={!canUndo && !editorView} tooltip="Undo">
        <Undo2 size={18} strokeWidth={2} />
      </ToolbarButton>
      <ToolbarButton onClick={handleRedo} disabled={!canRedo && !editorView} tooltip="Redo">
        <Redo2 size={18} strokeWidth={2} />
      </ToolbarButton>
      
      <Divider />
      
      {/* Zoom */}
      <OverflowDropdown menuLabel="Zoom" triggerLabel={`${zoom}%`} hint={`${zoom}%`} width={80}>
        {zoomLevels.map((level) => (
          <DropdownItem key={level} onClick={() => setZoom(level)} active={zoom === level}>{level}%</DropdownItem>
        ))}
      </OverflowDropdown>

      <Divider />

      {/* Heading Style */}
      <OverflowDropdown menuLabel="Style" triggerIcon={<Type size={14} style={{ marginRight: 4 }} />} triggerLabel={currentHeadingLabel} hint={currentHeadingLabel} width={160} minWidth={100}>
        {headingOptions.map((option) => {
          const isActive = option.level === 'normal'
            ? activeMarks?.blockType === 'paragraph'
            : activeMarks?.blockType === 'heading' && activeMarks?.headingLevel === option.level;
          return (
            <DropdownItem
              key={option.level}
              onClick={() => { if (option.level === 'normal') runCommand(commands.setParagraph()); else runCommand(commands.setHeading(option.level)); }}
              active={isActive}
              style={{ fontSize: option.size, fontWeight: option.level !== 'normal' ? 600 : 400 }}
            >
              {option.label}
            </DropdownItem>
          );
        })}
      </OverflowDropdown>

      <Divider />

      {/* Font Family */}
      <OverflowDropdown menuLabel="Font" triggerLabel={currentFont} hint={currentFont} width={180} minWidth={90}>
        {fonts.map((font) => (
          <DropdownItem key={font} onClick={() => runCommand(commands.setFontFamily(font))} active={currentFont === font} style={{ fontFamily: font }}>{font}</DropdownItem>
        ))}
      </OverflowDropdown>
      
      {/* Font Size */}
      <DocFontSizeControl value={currentFontSize} sizes={fontSizes} onPick={(s) => runCommand(commands.setFontSize(s))} />
      
      <Divider />
      
      {/* Text Formatting */}
      <ToolbarButton onClick={() => runCommand(commands.toggleBold())} active={activeMarks?.bold} tooltip="Bold">
        <Bold size={18} strokeWidth={2.5} />
      </ToolbarButton>
      <ToolbarButton onClick={() => runCommand(commands.toggleItalic())} active={activeMarks?.italic} tooltip="Italic">
        <Italic size={18} strokeWidth={2} />
      </ToolbarButton>
      <ToolbarButton onClick={() => runCommand(commands.toggleUnderline())} active={activeMarks?.underline} tooltip="Underline">
        <Underline size={18} strokeWidth={2} />
      </ToolbarButton>
      <ToolbarButton onClick={() => runCommand(commands.toggleStrikethrough())} active={activeMarks?.strikethrough} tooltip="Strikethrough">
        <Strikethrough size={18} strokeWidth={2} />
      </ToolbarButton>
      
      {/* Text Color */}
      <DocColorControl icon={<Palette size={16} strokeWidth={2} />} label="Text color" value={activeMarks?.textStyle?.color} onApply={(color) => runCommand(commands.setTextColor(color))} />

      {/* Highlight Color */}
      <DocColorControl icon={<Highlighter size={18} strokeWidth={2} />} label={selectedCell ? 'Cell background' : 'Highlight'} value={activeMarks?.textStyle?.backgroundColor} onApply={(color) => handleHighlightColor(color)} />
      
      <Divider />
      
      {/* Link & Image */}
      <ToolbarButton onClick={() => setShowLinkDialog(true)} active={!!existingLink} tooltip="Insert link">
        <Link size={18} strokeWidth={2} />
      </ToolbarButton>
      <ToolbarButton onClick={() => setShowImageDialog(true)} tooltip="Insert image">
        <Image size={18} strokeWidth={2} />
      </ToolbarButton>
      
      <Divider />
      
      {/* Alignment */}
      <ToolbarButton 
        onClick={() => runCommand(commands.alignLeft())} 
        active={!activeMarks?.alignment || activeMarks?.alignment === 'left'}
        tooltip="Align left"
      >
        <AlignLeft size={18} strokeWidth={2} />
      </ToolbarButton>
      <ToolbarButton 
        onClick={() => runCommand(commands.alignCenter())} 
        active={activeMarks?.alignment === 'center'}
        tooltip="Align center"
      >
        <AlignCenter size={18} strokeWidth={2} />
      </ToolbarButton>
      <ToolbarButton 
        onClick={() => runCommand(commands.alignRight())} 
        active={activeMarks?.alignment === 'right'}
        tooltip="Align right"
      >
        <AlignRight size={18} strokeWidth={2} />
      </ToolbarButton>
      <ToolbarButton 
        onClick={() => runCommand(commands.alignJustify())} 
        active={activeMarks?.alignment === 'justify'}
        tooltip="Justify"
      >
        <AlignJustify size={18} strokeWidth={2} />
      </ToolbarButton>
      
      <Divider />
      
      {/* Lists */}
      <ToolbarButton 
        onClick={() => runCommand(commands.toggleBulletList())} 
        active={activeMarks?.listType === 'bullet_list'}
        tooltip="Bullet list"
      >
        <List size={18} strokeWidth={2} />
      </ToolbarButton>
      <ToolbarButton 
        onClick={() => runCommand(commands.toggleOrderedList())} 
        active={activeMarks?.listType === 'ordered_list'}
        tooltip="Numbered list"
      >
        <ListOrdered size={18} strokeWidth={2} />
      </ToolbarButton>
      <ToolbarButton onClick={() => runCommand(commands.decreaseIndent())} tooltip="Decrease indent">
        <IndentDecrease size={18} strokeWidth={2} />
      </ToolbarButton>
      <ToolbarButton onClick={() => runCommand(commands.increaseIndent())} tooltip="Increase indent">
        <IndentIncrease size={18} strokeWidth={2} />
      </ToolbarButton>
      
      <Divider />
      
      {/* Table */}
      <div style={{ position: 'relative' }}>
        <ToolbarButton onClick={() => setShowTablePicker(!showTablePicker)} tooltip="Insert table">
          <Table size={18} strokeWidth={2} />
        </ToolbarButton>
        <TableSizePicker
          editorView={editorView ?? null}
          isOpen={showTablePicker}
          onClose={() => setShowTablePicker(false)}
        />
      </div>
      
      {/* More formatting */}
      <ToolbarButton onClick={() => runCommand(commands.insertHorizontalRule())} tooltip="Horizontal line">
        <SeparatorHorizontal size={18} strokeWidth={2} />
      </ToolbarButton>
      <ToolbarButton onClick={() => runCommand(commands.insertPageBreak())} tooltip="Page break">
        <FileDown size={18} strokeWidth={2} />
      </ToolbarButton>
      
      <Divider />
      
      {/* Super/Subscript */}
      <ToolbarButton onClick={() => runCommand(commands.toggleSuperscript())} active={activeMarks?.superscript} tooltip="Superscript">
        <Superscript size={18} strokeWidth={2} />
      </ToolbarButton>
      <ToolbarButton onClick={() => runCommand(commands.toggleSubscript())} active={activeMarks?.subscript} tooltip="Subscript">
        <Subscript size={18} strokeWidth={2} />
      </ToolbarButton>
      
      <Divider />

      {/* Comment */}
      <ToolbarButton onClick={onAddComment} tooltip="Comment">
        <MessageSquarePlus size={18} strokeWidth={2} />
      </ToolbarButton>

      {/* Page Setup */}
      <ToolbarButton onClick={onPageSetup} tooltip="Page setup">
        <Settings2 size={18} strokeWidth={2} />
      </ToolbarButton>
      </ResponsiveToolbar>

      {/* Dialogs */}
      <LinkDialog
        editorView={editorView ?? null}
        isOpen={showLinkDialog}
        onClose={() => setShowLinkDialog(false)}
        initialUrl={existingLink?.href || ''}
        initialText={getSelectedText()}
        isEditing={!!existingLink}
      />
      <ImageDialog
        editorView={editorView ?? null}
        isOpen={showImageDialog}
        onClose={() => setShowImageDialog(false)}
      />
    </div>
  );
});
