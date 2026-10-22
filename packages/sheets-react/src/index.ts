// Main workbook components (Canvas-based)
export * from './components/WorkbookCanvas';
export * from './components/CanvasGrid';
export * from './components/EditOverlay';

// Re-export WorkbookCanvas as Workbook for backwards compatibility
export { WorkbookCanvas as Workbook } from './components/WorkbookCanvas';

// Shared UI components
export * from './components/FormulaBar';
export * from './components/Toolbar';
export * from './components/SheetTabs';
export * from './components/ContextMenu';
export * from './components/CommentsPanel';

// Context
export * from './context/WorkbookContext';

// Hooks
export * from './hooks/useComments';


// File open / download helpers (XLSX, CSV)
export * from './file-io';

// Import loss report (what an Open would not bring across) — see inspectWorkbookFile / inspectDocumentFile.
export { hasImportLoss, describeImportLoss } from '@b32nio/spindle-shared';
export type { ImportLossReport, ImportLossItem, ImportLossFormat } from '@b32nio/spindle-shared';
