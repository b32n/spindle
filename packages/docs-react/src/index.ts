// @b32nio/spindle-docs-react
// React components for Spindle Docs

// Context
export { DocumentProvider, useDocument, useSelection, useSections, useHistory } from './context/DocumentContext';

// Components
export { DocumentEditor } from './components/DocumentEditor';
export { Toolbar } from './components/Toolbar';
export { CommentsSidebar } from './components/CommentsSidebar';
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
