import React, { memo, useCallback, useState, useEffect } from 'react';
import { useWorkbook } from '../context/WorkbookContext';

interface SheetTabsProps {
  onSheetSelect?: (sheetId: string) => void;
  onSheetRename?: (sheetId: string, newName: string) => void;
  onSheetAdd?: () => void;
  onSheetDelete?: (sheetId: string) => void;
  /** Disables add/rename/delete; switching the active sheet still works. */
  readOnly?: boolean;
}

export const SheetTabs = memo(function SheetTabs({
  onSheetSelect,
  onSheetRename,
  onSheetAdd,
  onSheetDelete,
  readOnly = false,
}: SheetTabsProps) {
  const { workbook } = useWorkbook();
  const [editingSheetId, setEditingSheetId] = useState<string | null>(null);
  const [editingName, setEditingName] = useState('');
  const [, forceUpdate] = useState({});

  const sheets = Array.from(workbook.sheets.values());
  const activeSheetId = workbook.activeSheetId;

  // Force re-render when sheets change
  useEffect(() => {
    const handleSheetChange = () => {
      forceUpdate({});
    };
    
    workbook.on('sheetAdd', handleSheetChange);
    workbook.on('sheetDelete', handleSheetChange);
    
    return () => {
      workbook.off('sheetAdd', handleSheetChange);
      workbook.off('sheetDelete', handleSheetChange);
    };
  }, [workbook]);

  const handleSheetClick = useCallback(
    (sheetId: string) => {
      if (sheetId !== activeSheetId) {
        workbook.setActiveSheet(sheetId);
        onSheetSelect?.(sheetId);
      }
    },
    [activeSheetId, workbook, onSheetSelect]
  );

  const handleSheetDoubleClick = useCallback(
    (sheetId: string, currentName: string) => {
      if (readOnly) return;
      setEditingSheetId(sheetId);
      setEditingName(currentName);
    },
    [readOnly]
  );

  const handleRenameSubmit = useCallback(
    (sheetId: string) => {
      if (!readOnly && editingName.trim() && editingName !== workbook.getSheet(sheetId).name) {
        workbook.renameSheet(sheetId, editingName.trim());
        onSheetRename?.(sheetId, editingName.trim());
      }
      setEditingSheetId(null);
      setEditingName('');
    },
    [editingName, workbook, onSheetRename, readOnly]
  );

  const handleRenameKeyDown = useCallback(
    (e: React.KeyboardEvent, sheetId: string) => {
      if (e.key === 'Enter') {
        handleRenameSubmit(sheetId);
      } else if (e.key === 'Escape') {
        setEditingSheetId(null);
        setEditingName('');
      }
    },
    [handleRenameSubmit]
  );

  const handleAddSheet = useCallback(() => {
    if (readOnly) return;
    const newSheet = workbook.addSheet(`Sheet${sheets.length + 1}`);
    workbook.setActiveSheet(newSheet.id);
    onSheetAdd?.();
  }, [workbook, sheets.length, onSheetAdd, readOnly]);

  const handleDeleteSheet = useCallback(
    (sheetId: string, e: React.MouseEvent) => {
      e.stopPropagation();
      if (readOnly) return;
      if (sheets.length > 1) {
        workbook.deleteSheet(sheetId);
        onSheetDelete?.(sheetId);
      }
    },
    [workbook, sheets.length, onSheetDelete, readOnly]
  );

  return (
    <div
      className="sheet-tabs"
      style={{
        display: 'flex',
        alignItems: 'center',
        gap: '4px',
        borderTop: '1px solid rgba(15, 23, 42, 0.06)',
        background: 'rgba(255, 255, 255, 0.55)',
        backdropFilter: 'blur(12px) saturate(180%)',
        WebkitBackdropFilter: 'blur(12px) saturate(180%)',
        overflowX: 'auto',
        overflowY: 'hidden',
        minHeight: '38px',
        height: '38px',
        flexShrink: 0,
        padding: '0 8px',
        position: 'relative',
        zIndex: 10,
        fontFamily: '"Inter", "SF Pro Display", -apple-system, BlinkMacSystemFont, sans-serif',
      }}
    >
      {sheets.map((sheet) => {
        const isActive = sheet.id === activeSheetId;
        const isEditing = editingSheetId === sheet.id;

        return (
          <div
            key={sheet.id}
            className="sheet-tab"
            style={{
              position: 'relative',
              padding: '6px 30px 6px 14px',
              backgroundColor: isActive ? '#ffffff' : 'transparent',
              cursor: 'pointer',
              userSelect: 'none',
              minWidth: '92px',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              fontSize: '13px',
              fontWeight: isActive ? 600 : 500,
              color: isActive ? '#6366f1' : '#64748b',
              borderRadius: '9px',
              border: isActive ? '1px solid rgba(99, 102, 241, 0.25)' : '1px solid transparent',
              transition: 'all 0.15s ease',
              boxShadow: isActive ? '0 2px 6px rgba(99, 102, 241, 0.16)' : 'none',
            }}
            onClick={() => handleSheetClick(sheet.id)}
            onDoubleClick={() => handleSheetDoubleClick(sheet.id, sheet.name)}
            onMouseEnter={(e) => {
              if (!isActive) {
                e.currentTarget.style.backgroundColor = 'rgba(99, 102, 241, 0.08)';
              }
            }}
            onMouseLeave={(e) => {
              if (!isActive) {
                e.currentTarget.style.backgroundColor = 'transparent';
              }
            }}
          >
            {isEditing ? (
              <input
                type="text"
                value={editingName}
                onChange={(e) => setEditingName(e.target.value)}
                onBlur={() => handleRenameSubmit(sheet.id)}
                onKeyDown={(e) => handleRenameKeyDown(e, sheet.id)}
                onClick={(e) => e.stopPropagation()}
                style={{
                  border: '2px solid #6366f1',
                  outline: 'none',
                  padding: '4px 8px',
                  fontSize: '13px',
                  width: '100%',
                  maxWidth: '150px',
                  borderRadius: '7px',
                  backgroundColor: '#ffffff',
                  color: '#1e293b',
                  fontFamily: 'inherit',
                }}
                autoFocus
              />
            ) : (
              <>
                <span style={{ flex: 1, textAlign: 'center', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                  {sheet.name}
                </span>
                {sheets.length > 1 && !readOnly && (
                  <button
                    onClick={(e) => handleDeleteSheet(sheet.id, e)}
                    style={{
                      position: 'absolute',
                      right: '8px',
                      top: '50%',
                      transform: 'translateY(-50%)',
                      border: 'none',
                      background: 'transparent',
                      cursor: 'pointer',
                      fontSize: '15px',
                      color: '#94a3b8',
                      padding: '2px 6px',
                      lineHeight: 1,
                      borderRadius: '50%',
                      width: '18px',
                      height: '18px',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      transition: 'all 0.15s ease',
                    }}
                    title="Delete sheet"
                    onMouseEnter={(e) => {
                      e.currentTarget.style.backgroundColor = '#ef4444';
                      e.currentTarget.style.color = '#ffffff';
                    }}
                    onMouseLeave={(e) => {
                      e.currentTarget.style.backgroundColor = 'transparent';
                      e.currentTarget.style.color = '#94a3b8';
                    }}
                  >
                    ×
                  </button>
                )}
              </>
            )}
          </div>
        );
      })}
      {!readOnly && <button
        onClick={handleAddSheet}
        style={{
          padding: '0',
          border: 'none',
          backgroundColor: 'transparent',
          cursor: 'pointer',
          fontSize: '20px',
          color: '#64748b',
          lineHeight: 1,
          borderRadius: '8px',
          transition: 'all 0.15s ease',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          minWidth: '30px',
          height: '30px',
          flexShrink: 0,
        }}
        title="Add new sheet"
        onMouseEnter={(e) => {
          e.currentTarget.style.backgroundColor = 'rgba(99, 102, 241, 0.08)';
          e.currentTarget.style.color = '#6366f1';
        }}
        onMouseLeave={(e) => {
          e.currentTarget.style.backgroundColor = 'transparent';
          e.currentTarget.style.color = '#64748b';
        }}
      >
        +
      </button>}
    </div>
  );
});

