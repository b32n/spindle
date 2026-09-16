# Docs Component Reference

Complete guide to all React components in the `@b32nio/spindle-docs-react` package.

## High-Level Components

### DocumentEditor

The main document editor component providing a complete editing experience with toolbar, rulers, and paginated content.

```tsx
import { DocumentEditor } from '@b32nio/spindle-docs-react';

<DocumentEditor
  width={800}
  height={600}
  showToolbar={true}
  showRuler={true}
/>
```

#### Props

| Prop | Type | Default | Description |
|------|------|---------|-------------|
| `width` | `number \| string` | `'100%'` | Container width |
| `height` | `number \| string` | `'100%'` | Container height |
| `showToolbar` | `boolean` | `true` | Show the formatting toolbar |
| `showRuler` | `boolean` | `true` | Show horizontal and vertical rulers |

#### Features

- Full formatting toolbar with all text and paragraph options
- Horizontal ruler for margins and indentation
- Vertical ruler for page position
- Zoom controls
- Page setup modal
- Keyboard shortcuts (Cmd/Ctrl+Z for undo, etc.)

---

### TrueLayoutEditor

The core editor component with true line-level pagination. This is the recommended component for custom implementations.

```tsx
import { TrueLayoutEditor } from '@b32nio/spindle-docs-react';
import type { TrueLayoutEditorHandle } from '@b32nio/spindle-docs-react';

const editorRef = useRef<TrueLayoutEditorHandle>(null);

<TrueLayoutEditor
  ref={editorRef}
  initialBlocks={blocks}
  pageConfig={{
    width: 816,
    height: 1056,
    margins: { top: 96, bottom: 96, left: 96, right: 96 }
  }}
  zoom={1}
  editable={true}
  onDocChange={(blocks) => console.log('Changed:', blocks)}
  onSelectionChange={(state) => console.log('Selection:', state)}
  onReady={(view) => console.log('Editor ready')}
  pageGap={24}
  minLinesAtBreak={2}
  header={headerContent}
  footer={footerContent}
  documentTitle="My Document"
/>
```

#### Props

| Prop | Type | Default | Description |
|------|------|---------|-------------|
| `initialBlocks` | `Block[]` | Required | Initial document blocks |
| `pageConfig` | `PageConfig` | Required | Page dimensions and margins |
| `zoom` | `number` | `1` | Zoom level (1 = 100%) |
| `editable` | `boolean` | `true` | Whether content is editable |
| `onDocChange` | `(blocks: Block[]) => void` | - | Called when content changes |
| `onSelectionChange` | `(state: EditorState) => void` | - | Called when selection changes |
| `onCellSelectionChange` | `(selection: CellSelection \| null) => void` | - | Called when table cell selection changes |
| `onActivePageChange` | `(info: ActivePageInfo) => void` | - | Called when active page changes |
| `onReady` | `(view: EditorView) => void` | - | Called when editor is initialized |
| `width` | `number \| string` | `'100%'` | Container width |
| `height` | `number \| string` | `'100%'` | Container height |
| `pageGap` | `number` | `24` | Gap between pages in pixels |
| `minLinesAtBreak` | `number` | `2` | Minimum lines at page break (widow/orphan control) |
| `textStylePool` | `TextStylePool` | - | Shared style pool for efficiency |
| `header` | `HeaderFooterContent` | - | Header configuration |
| `footer` | `HeaderFooterContent` | - | Footer configuration |
| `documentTitle` | `string` | - | Title for dynamic fields |
| `onHeaderChange` | `(content: HeaderFooterContent) => void` | - | Called when header is edited |
| `onFooterChange` | `(content: HeaderFooterContent) => void` | - | Called when footer is edited |
| `headerFooterEditable` | `boolean` | `true` | Whether headers/footers are editable |

#### Handle Methods

Access via ref:

```typescript
interface TrueLayoutEditorHandle {
  getView(): EditorView | null;        // Get ProseMirror view
  getState(): EditorState | null;      // Get current state
  focus(): void;                       // Focus the editor
  hasFocus(): boolean;                 // Check focus state
  getLayout(): DocumentLayout | null;  // Get current layout
  reflow(): void;                      // Force re-layout
  getScrollContainer(): HTMLElement | null; // Get scroll container
}
```

---

### DocumentProvider

Context provider for document state management.

```tsx
import { DocumentImpl } from '@b32nio/spindle-docs-core';
import { DocumentProvider } from '@b32nio/spindle-docs-react';

const doc = new DocumentImpl('doc_1', 'My Document');

<DocumentProvider document={doc} zoom={100}>
  <DocumentEditor />
</DocumentProvider>
```

#### Props

| Prop | Type | Default | Description |
|------|------|---------|-------------|
| `document` | `DocumentImpl` | Required | The document model instance |
| `zoom` | `number` | `100` | Initial zoom level (percentage) |
| `children` | `ReactNode` | Required | Child components |

#### Context Hooks

```typescript
// Get document model and zoom controls
const { document, zoom, setZoom } = useDocument();

// Get current selection
const selection = useSelection();

// Get sections
const sections = useSections();

// Get history controls
const { canUndo, canRedo, undo, redo } = useHistory();
```

---

## Toolbar Components

### Toolbar

The formatting toolbar with all editing controls.

```tsx
import { Toolbar } from '@b32nio/spindle-docs-react';

<Toolbar
  editorView={editorView}
  activeMarks={activeMarks}
  selectedCell={selectedCell}
  onPageSetup={() => setShowPageSetup(true)}
/>
```

#### Props

| Prop | Type | Description |
|------|------|-------------|
| `editorView` | `EditorView \| null` | ProseMirror editor view |
| `activeMarks` | `ActiveMarks` | Currently active formatting |
| `selectedCell` | `CellSelection \| null` | Selected table cell |
| `onPageSetup` | `() => void` | Callback to open page setup |

#### Toolbar Sections

1. **Undo/Redo** - History navigation
2. **Zoom** - Document zoom level
3. **Heading Style** - Paragraph/Heading selection
4. **Font Family** - Font selection dropdown
5. **Font Size** - Size with +/- buttons
6. **Text Formatting** - Bold, italic, underline, strikethrough
7. **Colors** - Text color and highlight
8. **Links & Images** - Insert links and images
9. **Alignment** - Left, center, right, justify
10. **Lists** - Bullet and numbered lists, indentation
11. **Tables** - Table insertion
12. **Page Elements** - Horizontal rule, page break
13. **Super/Subscript** - Text positioning
14. **Page Setup** - Open page configuration

---

### ColorPicker

Color selection dropdown for text and background colors.

```tsx
import { ColorPicker } from '@b32nio/spindle-docs-react';

<ColorPicker
  currentColor={activeMarks?.textStyle?.color}
  onColorSelect={(color) => applyColor(color)}
  onClose={() => setShowPicker(false)}
  showNoColor={true}
  noColorLabel="Remove color"
/>
```

---

### TableSizePicker

Grid-based table size selection.

```tsx
import { TableSizePicker } from '@b32nio/spindle-docs-react';

<TableSizePicker
  editorView={editorView}
  isOpen={showTablePicker}
  onClose={() => setShowTablePicker(false)}
/>
```

---

## Dialog Components

### LinkDialog

Modal for inserting and editing hyperlinks.

```tsx
import { LinkDialog } from '@b32nio/spindle-docs-react';

<LinkDialog
  editorView={editorView}
  isOpen={showLinkDialog}
  onClose={() => setShowLinkDialog(false)}
  initialUrl="https://example.com"
  initialText="Link text"
  isEditing={false}
/>
```

---

### ImageDialog

Modal for inserting images by URL.

```tsx
import { ImageDialog } from '@b32nio/spindle-docs-react';

<ImageDialog
  editorView={editorView}
  isOpen={showImageDialog}
  onClose={() => setShowImageDialog(false)}
/>
```

---

### PageSetupModal

Modal for configuring page size, orientation, and margins.

```tsx
import { PageSetupModal } from '@b32nio/spindle-docs-react';

<PageSetupModal
  isOpen={showPageSetup}
  pageConfig={currentPageConfig}
  onClose={() => setShowPageSetup(false)}
  onConfirm={(config) => applyPageConfig(config)}
/>
```

#### Props

| Prop | Type | Description |
|------|------|-------------|
| `isOpen` | `boolean` | Whether modal is visible |
| `pageConfig` | `PageConfig` | Current page configuration |
| `onClose` | `() => void` | Close callback |
| `onConfirm` | `(config: PageConfig) => void` | Apply callback |

---

## Ruler Components

### Ruler (Horizontal)

Displays page width with margin handles.

```tsx
import { Ruler } from '@b32nio/spindle-docs-react';

<Ruler
  pageConfig={currentPageConfig}
  onMarginsChange={(margins) => updateMargins(margins)}
/>
```

---

### VerticalRuler

Displays page height with margin handles, synchronized with scroll position.

```tsx
import { VerticalRuler } from '@b32nio/spindle-docs-react';

<VerticalRuler
  pageConfig={currentPageConfig}
  onMarginsChange={(margins) => updateMargins(margins)}
  pageIndex={activePageInfo.pageIndex}
  pageHeight={activePageInfo.pageHeight}
  pageGap={24}
  scrollContainerRef={scrollContainerRef}
/>
```

---

## Header/Footer Components

### HeaderFooterEditor

In-place editor for header and footer content.

```tsx
import { HeaderFooterEditor } from '@b32nio/spindle-docs-react';

<HeaderFooterEditor
  type="header"
  content={headerContent}
  pageConfig={pageConfig}
  pageIndex={0}
  totalPages={5}
  documentTitle="My Document"
  onContentChange={(content) => updateHeader(content)}
  onClose={() => setEditing(false)}
/>
```

---

## Low-Level Components

### ProseMirrorEditor

Direct access to ProseMirror for advanced use cases.

```tsx
import { ProseMirrorEditor } from '@b32nio/spindle-docs-react';

<ProseMirrorEditor
  initialDoc={pmDoc}
  editable={true}
  onChange={(state) => handleChange(state)}
  onTransaction={(tr) => handleTransaction(tr)}
/>
```

---

### PageView

Renders a single page with content.

```tsx
import { PageView } from '@b32nio/spindle-docs-react';

<PageView
  pageConfig={pageConfig}
  blocks={pageBlocks}
  pageNumber={1}
/>
```

---

## Exported Types

```typescript
// Component Props
export type { TrueLayoutEditorProps, TrueLayoutEditorHandle };
export type { ProseMirrorEditorProps, ProseMirrorEditorRef, ActiveMarks };

// Layout Types
export type { PageConfig, PageFragment, PageLayout, DocumentLayout, LayoutOptions };

// Flow Blocks
export type { FlowBlock, Run, TextRun, ParagraphBlock, HeadingBlock, ListItemBlock, TableBlock, ImageBlock };

// Selection
export type { CaretPosition, SelectionRect, SelectionState, CellSelection };

// Header/Footer
export type { HeaderFooterContent, HeaderFooterParagraph, HeaderFooterInlineContent };
export type { HeaderFooterTextRun, DynamicFieldRun, DynamicFieldType, DynamicFieldContext };
```

---

## Styling

Components use inline styles with CSS variables for theming. The toolbar uses a modern glass-morphism design with:

- Backdrop blur effects
- Smooth hover transitions
- Tooltip support
- Dropdown menus with animations

To customize, you can:

1. Override CSS variables
2. Wrap components with styled containers
3. Use the low-level components with custom styling

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Cmd/Ctrl + Z` | Undo |
| `Cmd/Ctrl + Shift + Z` | Redo |
| `Cmd/Ctrl + Y` | Redo |
| `Cmd/Ctrl + B` | Bold |
| `Cmd/Ctrl + I` | Italic |
| `Cmd/Ctrl + U` | Underline |
| `Cmd/Ctrl + K` | Insert link |
| `Tab` | Increase indent |
| `Shift + Tab` | Decrease indent |
| `Enter` | New paragraph |
| `Shift + Enter` | Soft line break |
