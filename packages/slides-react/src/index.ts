// @b32nio/spindle-slides-react
// React components + hooks for Spindle Slides.

// Context
export { DeckProvider, useDeckContext } from './context/DeckContext';
export type { DeckProviderProps, DeckContextValue } from './context/DeckContext';
export { ElementStore } from './context/element-store';

// Interactions
export { NodeRegistry } from './interactions/node-registry';
export { TransientStore } from './interactions/transient-store';
export type { TransientState } from './interactions/transient-store';
export { EditingStore } from './interactions/editing-store';
export { applyFormat } from './interactions/formatting';
export type { FormatTarget } from './interactions/formatting';
export { screenToSlide, screenDistanceToSlide } from './interactions/coords';
export type { StageMetrics } from './interactions/coords';
export {
  createMoveGesture,
  createResizeGesture,
  createRotateGesture,
  createMarqueeGesture,
  expandGroups,
} from './interactions/gesture';
export type { Gesture, GestureContext } from './interactions/gesture';

// Hooks
export {
  useDeck,
  useSlideIds,
  useSlide,
  useSlideElementIds,
  useElement,
  useSelection,
  useActiveSlideId,
  useTheme,
  useClipboard,
  useKeyboardShortcuts,
  useEditingId,
  useCommentsOpen,
  useFilmstripOpen,
} from './hooks';
export { useComments } from './hooks/useComments';
export type { UseComments } from './hooks/useComments';
export { usePasteImport } from './hooks/usePasteImport';
export { clipboardToGrid, parseHtmlTable, parseTsv, normalizeGrid } from './interactions/table-paste';

// Components
export { SlidesEditor } from './components/SlidesEditor';
export { Toolbar } from './components/Toolbar';
export { TextFormatBar } from './components/TextFormatBar';
export { LineFormatBar } from './components/LineFormatBar';
export { ShapeFormatBar } from './components/ShapeFormatBar';
export { ImageFormatBar } from './components/ImageFormatBar';
export { TableFormatBar } from './components/TableFormatBar';
export { TableCellMenu } from './components/TableCellMenu';
export { TableGuttersOverlay } from './components/TableGuttersOverlay';
export { TableSelectionStore } from './interactions/table-selection-store';
export type { TableSelection } from './interactions/table-selection-store';
export { RichTextEditor } from './components/RichTextEditor';
export { NotesPanel } from './components/NotesPanel';
export { DeckControls } from './components/DeckControls';
export { PresentMode } from './components/PresentMode';
export { SlideStage } from './components/SlideStage';
export { InteractiveSlide } from './components/InteractiveSlide';
export { Filmstrip } from './components/Filmstrip';
export { SlideView, ScaledSlide } from './components/SlideView';
export { SelectionOverlay } from './components/SelectionOverlay';
export { GuidesOverlay } from './components/GuidesOverlay';
export { RemotePresenceOverlay } from './components/RemotePresenceOverlay';
export { CommentsPanel } from './components/CommentsPanel';
export { CommentBadgesOverlay } from './components/CommentBadgesOverlay';
export { ConnectionPointsOverlay } from './components/ConnectionPointsOverlay';
export { ConnectorStore } from './interactions/connector-store';
export type { ConnectorState, ConnectorDraft } from './interactions/connector-store';
export { Menu } from './components/Menu';
export type { MenuItem } from './components/Menu';
export { ContextMenu } from './components/ContextMenu';
export { SlideContextMenu } from './components/SlideContextMenu';
export { LayoutThumb } from './components/LayoutThumb';
export { ElementView } from './components/elements/ElementView';
export { StaticRichText } from './components/elements/StaticRichText';
export { shapeGeom } from './components/elements/shapes';
export type { ShapeGeom } from './components/elements/shapes';
