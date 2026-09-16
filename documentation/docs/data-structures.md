# Docs Data Structures

Complete type definitions for the `@b32nio/spindle-docs-core` package.

## Document Structure

### Document

The root document structure containing all content.

```typescript
interface Document {
  id: string;
  title: string;
  sections: Section[];
  defaultPageConfig: PageConfig;
  createdAt?: string;
  updatedAt?: string;
}
```

### Section

A section defines a page configuration boundary with its own content and optional header/footer.

```typescript
interface Section {
  id: string;
  pageConfig: PageConfig;
  blocks: Block[];
  header?: HeaderFooterContent;
  footer?: HeaderFooterContent;
}
```

### Serialized Data

For persistence, use these data types:

```typescript
interface DocumentData {
  id: string;
  title: string;
  sections: SectionData[];
  defaultPageConfig: PageConfig;
  textStylePool: Record<string, TextStyle>;
  paragraphStylePool: Record<string, ParagraphStyle>;
  createdAt?: string;
  updatedAt?: string;
}

interface SectionData {
  id: string;
  pageConfig: PageConfig;
  blocks: Block[];
  header?: HeaderFooterContent;
  footer?: HeaderFooterContent;
}
```

---

## Page Configuration

### PageConfig

```typescript
interface PageConfig {
  size: PageSize;
  margins: PageMargins;
  orientation: 'portrait' | 'landscape';
}

interface PageSize {
  w: number;  // Width in pixels (96 DPI)
  h: number;  // Height in pixels
}

interface PageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header?: number;  // Distance from top edge to header area
  footer?: number;  // Distance from bottom edge to footer area
}
```

### Standard Page Sizes

All measurements in pixels at 96 DPI:

```typescript
const PAGE_SIZES = {
  LETTER: { w: 816, h: 1056 },      // 8.5" x 11"
  LEGAL: { w: 816, h: 1344 },       // 8.5" x 14"
  A4: { w: 794, h: 1123 },          // 210mm x 297mm
  A5: { w: 559, h: 794 },           // 148mm x 210mm
  TABLOID: { w: 1056, h: 1632 },    // 11" x 17"
  EXECUTIVE: { w: 696, h: 1008 },   // 7.25" x 10.5"
  B5: { w: 708, h: 1001 },          // 182mm x 257mm
} as const;
```

### Default Configuration

```typescript
const DEFAULT_PAGE_CONFIG: PageConfig = {
  size: PAGE_SIZES.LETTER,
  margins: {
    top: 96,      // 1 inch
    right: 96,
    bottom: 96,
    left: 96,
    header: 48,   // 0.5 inch
    footer: 48,
  },
  orientation: 'portrait',
};
```

---

## Block Types

### Base Block

All blocks extend this base:

```typescript
interface BaseBlock {
  id: string;
  type: BlockType;
}

type BlockType = 
  | 'paragraph' 
  | 'heading' 
  | 'list-item' 
  | 'table' 
  | 'image' 
  | 'horizontal-rule' 
  | 'page-break';
```

### ParagraphBlock

```typescript
interface ParagraphBlock extends BaseBlock {
  type: 'paragraph';
  content: InlineContent[];
  styleId?: string;
  // Inline styles (when not using styleId)
  alignment?: TextAlignment;
  indent?: number;
  lineSpacing?: LineSpacing;
  spaceBefore?: number;
  spaceAfter?: number;
}
```

### HeadingBlock

```typescript
interface HeadingBlock extends BaseBlock {
  type: 'heading';
  level: HeadingLevel;  // 1 | 2 | 3 | 4 | 5 | 6
  content: InlineContent[];
  styleId?: string;
  alignment?: TextAlignment;
}
```

### ListItemBlock

```typescript
interface ListItemBlock extends BaseBlock {
  type: 'list-item';
  listType: ListType;   // 'bullet' | 'numbered'
  level: number;        // Nesting level (0-based)
  content: InlineContent[];
  styleId?: string;
}
```

### TableBlock

```typescript
interface TableBlock extends BaseBlock {
  type: 'table';
  rows: TableRow[];
  colWidths?: number[];
  styleId?: string;
}

interface TableRow {
  id: string;
  cells: TableCell[];
  height?: number;
}

interface TableCell {
  id: string;
  content: InlineContent[];
  colspan?: number;
  rowspan?: number;
  styleId?: string;
}
```

### ImageBlock

```typescript
interface ImageBlock extends BaseBlock {
  type: 'image';
  src: string;
  width: number;
  height: number;
  alt?: string;
  alignment?: 'left' | 'center' | 'right';
  caption?: InlineContent[];
}
```

### Special Blocks

```typescript
interface HorizontalRuleBlock extends BaseBlock {
  type: 'horizontal-rule';
}

interface PageBreakBlock extends BaseBlock {
  type: 'page-break';
}
```

### Block Union Type

```typescript
type Block =
  | ParagraphBlock
  | HeadingBlock
  | ListItemBlock
  | TableBlock
  | ImageBlock
  | HorizontalRuleBlock
  | PageBreakBlock;
```

---

## Inline Content

Content that appears within blocks.

### TextRun

```typescript
interface TextRun {
  type: 'text';
  text: string;
  styleId?: string;
  // Inline styles
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  fontFamily?: string;
  fontSize?: number;
  color?: string;
  backgroundColor?: string;
  smallCaps?: boolean;
}
```

### InlineImage

```typescript
interface InlineImage {
  type: 'image';
  src: string;
  width: number;
  height: number;
  alt?: string;
}
```

### InlineLink

```typescript
interface InlineLink {
  type: 'link';
  text: string;
  href: string;
  styleId?: string;
}
```

### DynamicFieldRun

Used in headers/footers for dynamic content:

```typescript
interface DynamicFieldRun {
  type: 'dynamicField';
  fieldType: DynamicFieldType;
  format?: string;  // Optional format string
}

type DynamicFieldType = 
  | 'pageNumber' 
  | 'totalPages' 
  | 'date' 
  | 'time' 
  | 'title';
```

### InlineContent Union

```typescript
type InlineContent = TextRun | InlineImage | InlineLink | DynamicFieldRun;
```

---

## Text Styling

### TextStyle

```typescript
interface TextStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontFamily?: string;
  fontSize?: number;
  color?: string;
  backgroundColor?: string;
  superscript?: boolean;
  subscript?: boolean;
  smallCaps?: boolean;
}

const DEFAULT_TEXT_STYLE: TextStyle = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  fontFamily: 'Arial',
  fontSize: 11,
  color: '#000000',
};
```

### ParagraphStyle

```typescript
interface ParagraphStyle {
  alignment?: TextAlignment;
  lineSpacing?: LineSpacing;
  spaceBefore?: number;
  spaceAfter?: number;
  firstLineIndent?: number;
  leftIndent?: number;
  rightIndent?: number;
  keepWithNext?: boolean;
  keepTogether?: boolean;
  pageBreakBefore?: boolean;
}

type TextAlignment = 'left' | 'center' | 'right' | 'justify';

interface LineSpacing {
  type: LineSpacingType;
  value?: number;
}

type LineSpacingType = 
  | 'single' 
  | 'onePointFive' 
  | 'double' 
  | 'atLeast' 
  | 'exactly' 
  | 'multiple';

const DEFAULT_PARAGRAPH_STYLE: ParagraphStyle = {
  alignment: 'left',
  lineSpacing: { type: 'single' },
  spaceBefore: 0,
  spaceAfter: 8,
  firstLineIndent: 0,
  leftIndent: 0,
  rightIndent: 0,
};
```

---

## Headers and Footers

### HeaderFooterContent

```typescript
interface HeaderFooterContent {
  blocks: HeaderFooterParagraph[];
  differentFirstPage?: boolean;
  firstPageBlocks?: HeaderFooterParagraph[];
}

interface HeaderFooterParagraph {
  type: 'paragraph';
  content: HeaderFooterInlineContent[];
  alignment?: 'left' | 'center' | 'right';
}

type HeaderFooterInlineContent = 
  | HeaderFooterTextRun 
  | DynamicFieldRun 
  | HeaderFooterImageRun;

interface HeaderFooterTextRun {
  type: 'text';
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontFamily?: string;
  color?: string;
}

interface HeaderFooterImageRun {
  type: 'image';
  src: string;
  alt?: string;
  width?: number;
  height?: number;
}
```

### Example Header

```typescript
const header: HeaderFooterContent = {
  blocks: [
    {
      type: 'paragraph',
      content: [
        { type: 'text', text: 'Document: ' },
        { type: 'dynamicField', fieldType: 'title' },
      ],
      alignment: 'left',
    },
    {
      type: 'paragraph',
      content: [
        { type: 'text', text: 'Page ' },
        { type: 'dynamicField', fieldType: 'pageNumber' },
        { type: 'text', text: ' of ' },
        { type: 'dynamicField', fieldType: 'totalPages' },
      ],
      alignment: 'right',
    },
  ],
  differentFirstPage: true,
  firstPageBlocks: [], // No header on first page
};
```

---

## Selection

### TextSelection

```typescript
interface TextSelection {
  anchor: CursorPosition;
  focus: CursorPosition;
  isCollapsed: boolean;
}

interface CursorPosition {
  blockId: string;
  offset: number;
  runIndex?: number;
}
```

---

## Style Pools

For efficient style management, use style pools:

```typescript
interface TextStylePool {
  styles: Map<string, TextStyle>;
  getOrCreate(style: TextStyle): string;      // Returns styleId
  get(styleId: string): TextStyle | undefined;
  getAllStyles(): Map<string, TextStyle>;
}

interface ParagraphStylePool {
  styles: Map<string, ParagraphStyle>;
  getOrCreate(style: ParagraphStyle): string;
  get(styleId: string): ParagraphStyle | undefined;
  getAllStyles(): Map<string, ParagraphStyle>;
}
```

### Usage

```typescript
import { TextStylePoolImpl } from '@b32nio/spindle-docs-core';

const pool = new TextStylePoolImpl();

// Creating styles
const boldId = pool.getOrCreate({ bold: true });
const sameId = pool.getOrCreate({ bold: true }); // Returns same ID

// Retrieving styles
const style = pool.get(boldId); // { bold: true }
```

---

## Events

### Event Types

```typescript
type DocumentEventType =
  | 'documentChange'
  | 'blockChange'
  | 'blockAdd'
  | 'blockDelete'
  | 'selectionChange'
  | 'pageConfigChange'
  | 'historyChange';

interface DocumentEventData {
  type: DocumentEventType;
  payload: unknown;
}

type DocumentEventHandler = (data: DocumentEventData) => void;
```

### Subscribing to Events

```typescript
// Subscribe
const unsubscribe = doc.on('documentChange', (event) => {
  console.log('Changed:', event.payload);
});

// Unsubscribe
unsubscribe();
// or
doc.off('documentChange', handler);
```

---

## History

### History Types

```typescript
interface HistoryEntry {
  timestamp: number;
  snapshot: DocumentSnapshot;
  description?: string;
}

interface DocumentSnapshot {
  sections: Section[];
  selection?: TextSelection;
}
```

### Using History

```typescript
// Record checkpoint
doc.recordHistory('Added paragraph');

// Undo/Redo
if (doc.canUndo()) {
  doc.undo();
}

if (doc.canRedo()) {
  doc.redo();
}
```

---

## Layout Types (React Package)

Types from `@b32nio/spindle-docs-react` for the layout engine:

### FlowBlock

Internal representation for layout:

```typescript
interface FlowBlock {
  id: string;
  kind: 'paragraph' | 'heading' | 'listItem' | 'table' | 'image' | 'hr' | 'pageBreak';
  runs?: Run[];
  level?: number;      // Heading level or list nesting
  listType?: 'bullet' | 'ordered';
  alignment?: 'left' | 'center' | 'right' | 'justify';
  // ... additional properties per kind
}

interface Run {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  // ... other text properties
}
```

### PageFragment

A portion of a block on a page:

```typescript
interface PageFragment {
  blockId: string;
  blockIndex: number;
  fromLine: number;    // Starting line (0-based)
  toLine: number;      // Ending line (exclusive)
  x: number;           // X position in content area
  y: number;           // Y position in content area
  height: number;
  isFirstFragment: boolean;
  isLastFragment: boolean;
}
```

### PageLayout

```typescript
interface PageLayout {
  pageIndex: number;
  fragments: PageFragment[];
  contentHeight: number;
  remainingHeight: number;
}
```

### DocumentLayout

```typescript
interface DocumentLayout {
  pages: PageLayout[];
  totalHeight: number;
  pageConfig: PageConfig;
  scale: number;
  pageGap: number;
}
```

### Measurements

```typescript
interface Measure {
  kind: 'block' | 'table' | 'image' | 'pageBreak';
  height: number;
  lines?: LineMeasure[];
}

interface LineMeasure {
  y: number;       // Y offset from block top
  height: number;  // Line height
}
```

---

## Creating Content

### Helper Functions

```typescript
import { 
  createParagraphFromText, 
  createHeading,
  createListItem,
  createTable,
} from '@b32nio/spindle-docs-core';

// Create paragraph
const para = createParagraphFromText('Hello, world!');

// Create heading
const h1 = createHeading(1, 'Document Title');

// Create list item
const item = createListItem('bullet', 0, 'First item');

// Create table
const table = createTable(3, 3); // 3 rows, 3 columns
```

### Creating Documents

```typescript
import { DocumentImpl, createDocument, createSection } from '@b32nio/spindle-docs-core';

// Using class
const doc = new DocumentImpl('doc_1', 'My Document');

// Using function
const document = createDocument('My Document', customPageConfig);

// Add section
const section = doc.addSection(pageConfig);
```
