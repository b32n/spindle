/**
 * TrueLayoutDemo - Demo of the True Layout Engine
 * 
 * This demo showcases the new SuperDoc-style architecture with:
 * - Hidden ProseMirror editor
 * - Line-level pagination
 * - Independent page rendering
 */

import { useState, useRef, useEffect } from 'react';
import { TrueLayoutEditor, TrueLayoutEditorHandle, PageConfig } from '@b32nio/spindle-docs-react';
import { 
  createParagraphFromText, 
  createHeadingFromText,
  createListItemFromText,
  Block,
} from '@b32nio/spindle-docs-core';

// Sample content that will span multiple pages
function createSampleContent(): Block[] {
  const blocks: Block[] = [
    createHeadingFromText(1, 'True Layout Engine Demo'),
    createParagraphFromText('This demonstrates the new SuperDoc-style layout engine with line-level pagination. Unlike the previous approach that clipped content at arbitrary pixel positions, this engine:'),
    createListItemFromText('bullet', 'Converts content to abstract FlowBlocks', 0),
    createListItemFromText('bullet', 'Measures each block with line-level precision', 0),
    createListItemFromText('bullet', 'Computes layout with proper page breaks at line boundaries', 0),
    createListItemFromText('bullet', 'Renders each page independently with only its content', 0),
    createHeadingFromText(2, 'How It Works'),
    createParagraphFromText('The architecture follows SuperDoc\'s approach:'),
    createParagraphFromText('1. A hidden ProseMirror editor handles all text editing, keyboard input, and selection management. This editor is positioned off-screen but remains focusable.'),
    createParagraphFromText('2. When the document changes, we convert the ProseMirror DOM to FlowBlocks - an abstract representation of paragraphs, headings, lists, tables, and other content.'),
    createParagraphFromText('3. A DOM Measurer renders each block to a hidden container and uses getClientRects() to measure line heights with pixel precision.'),
    createParagraphFromText('4. The Layout Engine takes these measurements and assigns content to pages. When a block doesn\'t fit entirely, it\'s split at line boundaries - no more mid-line clipping!'),
    createParagraphFromText('5. The DOM Painter renders each page independently. Each page only contains the fragments assigned to it, so there\'s no need for overflow:hidden clipping.'),
    createParagraphFromText('6. An Input Bridge forwards keyboard and mouse events from the visible pages to the hidden editor, maintaining the editing experience.'),
    createParagraphFromText('7. A Selection Overlay renders the cursor and selection highlights on top of the pages, mapped from the ProseMirror selection state.'),
    createHeadingFromText(2, 'Benefits'),
    createParagraphFromText('This architecture solves several problems with the previous approach:'),
    createListItemFromText('bullet', 'No more lines getting cut in half at page boundaries', 0),
    createListItemFromText('bullet', 'Proper widow/orphan control - we can ensure minimum lines at breaks', 0),
    createListItemFromText('bullet', 'Selection and cursor work correctly across page boundaries', 0),
    createListItemFromText('bullet', 'Each page can be rendered and printed independently', 0),
    createListItemFromText('bullet', 'Better performance for large documents - only visible pages need to be in the DOM', 0),
    createHeadingFromText(2, 'Try It Out'),
    createParagraphFromText('Type in this editor to see the layout update in real-time. Add enough content to create multiple pages and observe how paragraphs are split cleanly at line boundaries.'),
    createParagraphFromText('You can also try:'),
    createListItemFromText('bullet', 'Changing the zoom level to see the layout adapt', 0),
    createListItemFromText('bullet', 'Adding long paragraphs that span multiple pages', 0),
    createListItemFromText('bullet', 'Creating headings and lists', 0),
    createListItemFromText('bullet', 'Selecting text across page boundaries', 0),
  ];
  
  // Add some filler paragraphs to demonstrate multi-page layout
  for (let i = 1; i <= 5; i++) {
    blocks.push(createHeadingFromText(3, `Section ${i}`));
    blocks.push(createParagraphFromText(
      `This is paragraph ${i} of the filler content. It's designed to push the document to multiple pages so you can see how the layout engine handles page breaks. The text wraps naturally and when it reaches the bottom of a page, it continues on the next page without any mid-line clipping. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.`
    ));
  }
  
  return blocks;
}

// Default page config (Letter size, 1-inch margins)
const defaultPageConfig: PageConfig = {
  width: 816,  // 8.5" at 96 DPI
  height: 1056, // 11" at 96 DPI
  margins: {
    top: 96,    // 1"
    bottom: 96, // 1"
    left: 96,   // 1"
    right: 96,  // 1"
  },
};

export function TrueLayoutDemo() {
  const editorRef = useRef<TrueLayoutEditorHandle>(null);
  const [zoom, setZoom] = useState(1);
  const [blocks] = useState(createSampleContent);
  const [pageCount, setPageCount] = useState(1);
  
  // Update page count when layout changes
  useEffect(() => {
    const interval = setInterval(() => {
      const layout = editorRef.current?.getLayout();
      if (layout) {
        setPageCount(layout.pages.length);
      }
    }, 500);
    
    return () => clearInterval(interval);
  }, []);
  
  const handleZoomIn = () => {
    setZoom(z => Math.min(z + 0.1, 2));
  };
  
  const handleZoomOut = () => {
    setZoom(z => Math.max(z - 0.1, 0.5));
  };
  
  const handleResetZoom = () => {
    setZoom(1);
  };
  
  const handleFocus = () => {
    editorRef.current?.focus();
  };
  
  return (
    <div style={{
      display: 'flex',
      flexDirection: 'column',
      height: '100vh',
      backgroundColor: '#f8f9fa',
    }}>
      {/* Toolbar */}
      <div style={{
        display: 'flex',
        alignItems: 'center',
        gap: 16,
        padding: '12px 24px',
        backgroundColor: '#fff',
        borderBottom: '1px solid #e0e0e0',
        flexShrink: 0,
      }}>
        <h1 style={{ margin: 0, fontSize: 18, fontWeight: 500 }}>
          True Layout Engine Demo
        </h1>
        
        <div style={{ flex: 1 }} />
        
        <span style={{ fontSize: 14, color: '#5f6368' }}>
          Pages: {pageCount}
        </span>
        
        <div style={{
          display: 'flex',
          alignItems: 'center',
          gap: 8,
          padding: '4px 8px',
          backgroundColor: '#f1f3f4',
          borderRadius: 4,
        }}>
          <button
            onClick={handleZoomOut}
            style={{
              border: 'none',
              background: 'none',
              cursor: 'pointer',
              padding: '4px 8px',
              fontSize: 16,
            }}
          >
            −
          </button>
          <span
            onClick={handleResetZoom}
            style={{
              fontSize: 14,
              minWidth: 50,
              textAlign: 'center',
              cursor: 'pointer',
            }}
          >
            {Math.round(zoom * 100)}%
          </span>
          <button
            onClick={handleZoomIn}
            style={{
              border: 'none',
              background: 'none',
              cursor: 'pointer',
              padding: '4px 8px',
              fontSize: 16,
            }}
          >
            +
          </button>
        </div>
        
        <button
          onClick={handleFocus}
          style={{
            padding: '8px 16px',
            backgroundColor: '#1a73e8',
            color: '#fff',
            border: 'none',
            borderRadius: 4,
            cursor: 'pointer',
            fontSize: 14,
          }}
        >
          Focus Editor
        </button>
      </div>
      
      {/* Editor */}
      <div style={{ flex: 1, overflow: 'hidden' }}>
        <TrueLayoutEditor
          ref={editorRef}
          initialBlocks={blocks}
          pageConfig={defaultPageConfig}
          zoom={zoom}
          editable={true}
          pageGap={24}
          minLinesAtBreak={2}
          onDocChange={(newBlocks) => {
            console.log('Document changed:', newBlocks.length, 'blocks');
          }}
          onReady={(view) => {
            console.log('Editor ready');
            // Focus after a short delay
            setTimeout(() => view.focus(), 100);
          }}
        />
      </div>
    </div>
  );
}

export default TrueLayoutDemo;

