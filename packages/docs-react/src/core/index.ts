/**
 * Core components for the Spindle Docs layout engine
 * 
 * The True Layout Engine (TrueLayoutEditor) drives a hidden ProseMirror
 * instance and paints pages independently, giving line-level pagination
 * with no clipping.
 */

// ============================================================================
// True Layout Engine (Recommended)
// ============================================================================

// Main component
export { TrueLayoutEditor } from './TrueLayoutEditor';
export type { TrueLayoutEditorProps, TrueLayoutEditorHandle, ActivePageInfo } from './TrueLayoutEditor';

// FlowBlocks - Abstract content representation
export * from './flow-blocks';

// ProseMirror to FlowBlocks converter
export { proseMirrorToFlowBlocks, proseMirrorToSectionedFlowBlocks, createBlockPositionMap } from './pm-to-blocks';
export type { ConversionOptions, SectionedFlowBlocks } from './pm-to-blocks';

// DOM Measurer
export { DomMeasurer, hasLineData, getLineCount, getLineRangeHeight } from './measurer';
export type {
  MeasurerConfig,
  LineRunSegment,
  LineMeasure,
  ParagraphMeasure,
  HeadingMeasure,
  ListItemMeasure,
  TableMeasure,
  ImageMeasure,
  HorizontalRuleMeasure,
  PageBreakMeasure,
  Measure,
} from './measurer';

// True Layout Engine
export {
  computeTrueLayout,
  findBlockPage,
  findBlockFragments,
  getPageY,
  findPageAtY,
  getPageConfig,
  blockPositionToPageCoords,
} from './true-layout-engine';
export {
  computeSectionedLayout,
  canContinueOnPage,
  formatPageNumber,
  pickHeaderFooterVariant,
  resolveInheritedContent,
} from './section-layout';
export type { LayoutSection, LayoutSectionBreak, SectionedLayoutOptions } from './section-layout';
export type {
  PageConfig,
  PageFragment,
  PageLayout,
  DocumentLayout,
  LayoutOptions,
  HeaderFooterVariant,
} from './true-layout-engine';

// DOM Painter
export { DomPainter, createPageElement } from './dom-painter';
export type { 
  PainterConfig, 
  RenderedPage, 
  PageViewProps,
  // Header/Footer types
  HeaderFooterContent,
  HeaderFooterParagraph,
  HeaderFooterInlineContent,
  HeaderFooterTextRun,
  HeaderFooterImageRun,
  DynamicFieldRun,
  DynamicFieldType,
  DynamicFieldContext,
  HeaderFooterClickHandler,
} from './dom-painter';

// Input Bridge
export { InputBridge, createInputBridge, getInputBridge } from './input-bridge';
export type { 
  VisiblePosition, 
  SelectionChangeCallback, 
  FocusChangeCallback,
  CellSelection,
  CellSelectionChangeCallback,
} from './input-bridge';

// Selection Overlay
export { SelectionOverlayManager, getSelectionOverlayStyles } from './selection-overlay';
export type { CaretPosition, SelectionRect, SelectionState } from './selection-overlay';

// Table Interactions
export { 
  TableInteractionManager, 
  createTableInteractionManager, 
  getTableInteractionManager 
} from './table-interactions';
export type { TablePosition, ResizeState, ContextMenuAction } from './table-interactions';

// Image Interactions
export { 
  ImageInteractionManager, 
  createImageInteractionManager, 
  getImageInteractionManager 
} from './image-interactions';
export type { ImagePosition } from './image-interactions';
