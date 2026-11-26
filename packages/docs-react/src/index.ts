// @b32nio/spindle-docs-react
// React components for Spindle Docs

// Context
export { DocumentProvider, useDocument, useSelection, useSections, useHistory } from './context/DocumentContext';

// Components
export { DocumentEditor } from './components/DocumentEditor';
export { Toolbar } from './components/Toolbar';
export { CommentsSidebar } from './components/CommentsSidebar';
export { RevisionLegend } from './components/RevisionLegend';
export { ModeSwitch } from './components/ModeSwitch';
export type { EditorMode, ModeSwitchProps } from './components/ModeSwitch';
export { ReviewPanel } from './components/ReviewPanel';
export { RevisionViewBanner } from './components/RevisionViewBanner';
export type { RevisionViewBannerProps } from './components/RevisionViewBanner';
export type { ReviewPanelProps } from './components/ReviewPanel';
export type { RevisionLegendProps } from './components/RevisionLegend';
export { Ruler } from './components/Ruler';
export { PageView } from './components/PageView';
export { PageSetupModal } from './components/PageSetupModal';
export { HeaderFooterEditor } from './components/HeaderFooterEditor';

// ============================================================================
// True Layout Engine (Recommended)
// ============================================================================

export {
  // Main component
  TrueLayoutEditor,
  // FlowBlocks
  proseMirrorToFlowBlocks,
  createBlockPositionMap,
  // Measurer
  DomMeasurer,
  hasLineData,
  getLineCount,
  getLineRangeHeight,
  // Layout Engine
  computeTrueLayout,
  findBlockPage,
  findBlockFragments,
  getPageY,
  findPageAtY,
  getPageConfig,
  computeSectionedLayout,
  canContinueOnPage,
  formatPageNumber,
  pickHeaderFooterVariant,
  resolveInheritedContent,
  proseMirrorToSectionedFlowBlocks,
  blockPositionToPageCoords,
  // DOM Painter
  DomPainter,
  createPageElement,
  // Input Bridge
  InputBridge,
  createInputBridge,
  getInputBridge,
  // Selection Overlay
  SelectionOverlayManager,
  getSelectionOverlayStyles,
} from './core';

export type {
  TrueLayoutEditorProps,
  TrueLayoutEditorHandle,
  TrueLayoutSection,
  // FlowBlocks
  FlowBlock,
  Run,
  TextRun,
  ParagraphBlock,
  HeadingBlock,
  ListItemBlock,
  TableBlock,
  ImageBlock,
  // Measurer
  MeasurerConfig,
  LineMeasure,
  Measure,
  // Layout
  PageConfig,
  PageFragment,
  PageLayout,
  DocumentLayout,
  LayoutOptions,
  HeaderFooterVariant,
  LayoutSection,
  LayoutSectionBreak,
  SectionedLayoutOptions,
  SectionedFlowBlocks,
  // Painter
  PainterConfig,
  RenderedPage,
  // Selection
  CaretPosition,
  SelectionRect,
  SelectionState,
  // Header/Footer
  HeaderFooterContent,
  HeaderFooterParagraph,
  HeaderFooterInlineContent,
  HeaderFooterTextRun,
  DynamicFieldRun,
  DynamicFieldType,
  DynamicFieldContext,
  // Which view of the tracked changes the editor shows
  RevisionView,
} from './core';

// ProseMirror components (for direct access if needed)
export { 
  ProseMirrorEditor, 
  useProseMirrorCommands,
} from './components/ProseMirrorEditor';

export type {
  ProseMirrorEditorProps,
  ProseMirrorEditorRef,
  ActiveMarks,
} from './components/ProseMirrorEditor';

// File open / download helpers (DOCX)
export * from './file-io';

// Import loss report (what an Open would not bring across) — see inspectWorkbookFile / inspectDocumentFile.
export { hasImportLoss, describeImportLoss } from '@b32nio/spindle-shared';
export type { ImportLossReport, ImportLossItem, ImportLossFormat } from '@b32nio/spindle-shared';
export { ImportConfirmDialog } from '@b32nio/spindle-shared/react';
export type { ImportConfirmation, ImportConfirmDialogProps } from '@b32nio/spindle-shared/react';
