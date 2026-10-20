import { toggleMark, setBlockType } from 'prosemirror-commands';
import { liftListItem, sinkListItem, wrapInList } from 'prosemirror-schema-list';
import type { Schema, MarkType, NodeType, Attrs, Mark } from 'prosemirror-model';
import { TextSelection } from 'prosemirror-state';
import type { Command, EditorState, Transaction } from 'prosemirror-state';

/**
 * Command to toggle a mark with optional attributes
 */
export function toggleMarkCommand(markType: MarkType, attrs?: Attrs): Command {
  return toggleMark(markType, attrs);
}

/**
 * Command to set or update mark attributes (for textStyle marks like color, fontSize)
 * Merges new attributes with existing mark attributes to preserve fontFamily, color, etc.
 */
export function setMarkAttrs(markType: MarkType, attrs: Attrs): Command {
  return (state: EditorState, dispatch?: (tr: Transaction) => void): boolean => {
    const { from, to, empty } = state.selection;
    
    if (dispatch) {
      if (empty) {
        // At cursor - merge with stored marks or marks at position
        const storedMarks = state.storedMarks || state.doc.resolve(from).marks();
        const existingMark = storedMarks.find((m: Mark) => m.type === markType);
        const existingAttrs = existingMark ? existingMark.attrs : {};
        const mergedAttrs = { ...existingAttrs, ...attrs };
        const mark = markType.create(mergedAttrs);
        const tr = state.tr.addStoredMark(mark);
        dispatch(tr);
      } else {
        // Selection - for each position, merge with existing mark attributes
        // First, collect existing marks in the selection to get a base
        let baseAttrs: Attrs = {};
        state.doc.nodesBetween(from, to, (node) => {
          if (node.isText && node.marks) {
            const existingMark = node.marks.find((m: Mark) => m.type === markType);
            if (existingMark) {
              baseAttrs = { ...baseAttrs, ...existingMark.attrs };
            }
          }
        });
        const mergedAttrs = { ...baseAttrs, ...attrs };
        const tr = state.tr.addMark(from, to, markType.create(mergedAttrs));
        dispatch(tr);
      }
    }
    return true;
  };
}

/**
 * Command to remove a mark
 */
export function removeMarkCommand(markType: MarkType): Command {
  return (state: EditorState, dispatch?: (tr: Transaction) => void): boolean => {
    const { from, to, empty } = state.selection;
    
    if (empty) {
      if (dispatch) {
        dispatch(state.tr.removeStoredMark(markType));
      }
      return true;
    }
    
    if (!state.doc.rangeHasMark(from, to, markType)) {
      return false;
    }
    
    if (dispatch) {
      dispatch(state.tr.removeMark(from, to, markType));
    }
    return true;
  };
}

/**
 * Command to set block type (paragraph, heading, etc.)
 */
export function setBlockTypeCommand(nodeType: NodeType, attrs?: Attrs): Command {
  return setBlockType(nodeType, attrs);
}

/**
 * Toggle a list type over the selection:
 *  - already in a list of this type: lift the selected items out of it;
 *  - in a list of the other type: convert that list to this type;
 *  - otherwise: wrap the selected blocks in a new list.
 */
function toggleListType(schema: Schema, listType: NodeType): Command {
  const listTypes = [schema.nodes.bullet_list, schema.nodes.ordered_list].filter(Boolean);
  return (state, dispatch) => {
    const { $from, $to } = state.selection;
    if (!$from.blockRange($to)) return false;

    // Lists the selection touches: the enclosing one (caret inside an item)
    // plus any it spans (e.g. select-all, where $from sits at the doc level).
    const touched: Array<{ pos: number; type: NodeType }> = [];
    for (let d = $from.depth; d > 0; d--) {
      if (listTypes.includes($from.node(d).type)) {
        touched.push({ pos: $from.before(d), type: $from.node(d).type });
        break;
      }
    }
    state.doc.nodesBetween($from.pos, $to.pos, (node, pos) => {
      if (listTypes.includes(node.type)) {
        if (!touched.some((t) => t.pos === pos)) touched.push({ pos, type: node.type });
        return false;
      }
      return true;
    });

    if (touched.length > 0) {
      if (touched.every((t) => t.type === listType)) {
        return liftListItem(schema.nodes.list_item)(state, dispatch);
      }
      if (dispatch) {
        const tr = state.tr;
        for (const t of touched) {
          if (t.type !== listType) tr.setNodeMarkup(t.pos, listType);
        }
        dispatch(tr.scrollIntoView());
      }
      return true;
    }

    return wrapInList(listType)(state, dispatch);
  };
}

/**
 * Command to toggle bullet list
 */
export function toggleBulletList(schema: Schema): Command {
  return toggleListType(schema, schema.nodes.bullet_list);
}

/**
 * Command to toggle ordered list
 */
export function toggleOrderedList(schema: Schema): Command {
  return toggleListType(schema, schema.nodes.ordered_list);
}

/**
 * Command to increase indent (sink list item or add indent to paragraph)
 */
export function increaseIndent(schema: Schema): Command {
  return (state, dispatch) => {
    // Try to sink list item first
    if (schema.nodes.list_item && sinkListItem(schema.nodes.list_item)(state, dispatch)) {
      return true;
    }
    
    // Otherwise, increase paragraph indent
    const { $from, $to } = state.selection;
    let changed = false;
    
    if (dispatch) {
      const tr = state.tr;
      
      // Get the start position that includes the parent block
      // This ensures we find the paragraph even when selection is inside it
      const startPos = $from.start($from.depth);
      const endPos = $to.end($to.depth);
      
      state.doc.nodesBetween(startPos, endPos, (node, pos) => {
        if (node.type === schema.nodes.paragraph || node.type === schema.nodes.heading) {
          const currentIndent = node.attrs.indent || 0;
          tr.setNodeMarkup(pos, null, { ...node.attrs, indent: currentIndent + 1 });
          changed = true;
        }
      });
      if (changed) {
        dispatch(tr);
      }
    }
    
    return true; // Always return true to indicate command can be executed
  };
}

/**
 * Command to decrease indent (lift list item or reduce indent from paragraph)
 */
export function decreaseIndent(schema: Schema): Command {
  return (state, dispatch) => {
    // Try to lift list item first
    if (schema.nodes.list_item && liftListItem(schema.nodes.list_item)(state, dispatch)) {
      return true;
    }
    
    // Otherwise, decrease paragraph indent
    const { $from, $to } = state.selection;
    let changed = false;
    
    if (dispatch) {
      const tr = state.tr;
      
      // Get the start position that includes the parent block
      // This ensures we find the paragraph even when selection is inside it
      const startPos = $from.start($from.depth);
      const endPos = $to.end($to.depth);
      
      state.doc.nodesBetween(startPos, endPos, (node, pos) => {
        if (node.type === schema.nodes.paragraph || node.type === schema.nodes.heading) {
          const currentIndent = node.attrs.indent || 0;
          if (currentIndent > 0) {
            tr.setNodeMarkup(pos, null, { ...node.attrs, indent: currentIndent - 1 });
            changed = true;
          }
        }
      });
      if (changed) {
        dispatch(tr);
      }
    }
    
    return changed;
  };
}

/**
 * Command to set text alignment
 */
export function setAlignment(schema: Schema, alignment: 'left' | 'center' | 'right' | 'justify'): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;
    let changed = false;
    
    if (dispatch) {
      const tr = state.tr;
      
      // Get the start position that includes the parent block
      // This ensures we find the paragraph even when selection is inside it
      const startPos = $from.start($from.depth);
      const endPos = $to.end($to.depth);
      
      state.doc.nodesBetween(startPos, endPos, (node, pos) => {
        if (node.type === schema.nodes.paragraph || node.type === schema.nodes.heading) {
          tr.setNodeMarkup(pos, null, { ...node.attrs, alignment });
          changed = true;
        }
      });
      
      if (changed) {
        dispatch(tr);
      }
    }
    
    return changed;
  };
}

/**
 * Command to insert a horizontal rule
 */
export function insertHorizontalRule(schema: Schema): Command {
  return (state, dispatch) => {
    const hrType = schema.nodes.horizontal_rule;
    
    if (!hrType) return false;
    
    if (dispatch) {
      dispatch(state.tr.replaceSelectionWith(hrType.create()));
    }
    return true;
  };
}

/**
 * Command to insert a page break
 */
export function insertPageBreak(schema: Schema): Command {
  return (state, dispatch) => {
    const pageBreakType = schema.nodes.page_break;
    
    if (!pageBreakType) return false;
    
    if (dispatch) {
      dispatch(state.tr.replaceSelectionWith(pageBreakType.create()));
    }
    return true;
  };
}

/**
 * Command to insert an image
 */
export function insertImage(schema: Schema, attrs: { src: string; alt?: string; width?: number; height?: number }): Command {
  return (state, dispatch) => {
    const imageType = schema.nodes.image;
    
    if (!imageType) return false;
    
    if (dispatch) {
      dispatch(state.tr.replaceSelectionWith(imageType.create(attrs)));
    }
    return true;
  };
}

/**
 * Command to insert a link
 */
export function insertLink(schema: Schema, attrs: { href: string; title?: string }): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    const linkMark = schema.marks.link;
    
    if (!linkMark) return false;
    
    if (empty) return false;
    
    if (dispatch) {
      dispatch(state.tr.addMark(from, to, linkMark.create(attrs)));
    }
    return true;
  };
}

/**
 * Command to remove a link
 */
export function removeLink(schema: Schema): Command {
  return (state, dispatch) => {
    const { from, to } = state.selection;
    const linkMark = schema.marks.link;

    if (!linkMark) return false;

    if (!state.doc.rangeHasMark(from, to, linkMark)) {
      return false;
    }

    if (dispatch) {
      dispatch(state.tr.removeMark(from, to, linkMark));
    }
    return true;
  };
}

/**
 * Apply a comment mark over the current selection, anchoring it to a thread.
 */
export function addComment(schema: Schema, threadId: string): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    const commentMark = schema.marks.comment;
    if (!commentMark) return false;
    if (empty) return false;
    if (dispatch) {
      dispatch(state.tr.addMark(from, to, commentMark.create({ threadId })));
    }
    return true;
  };
}

/**
 * Strip a comment thread's mark from wherever it appears in the document —
 * used when a thread is deleted. Removes the exact mark instances so an
 * overlapping comment on the same text is left intact.
 */
export function removeComment(schema: Schema, threadId: string): Command {
  return (state, dispatch) => {
    const commentMark = schema.marks.comment;
    if (!commentMark) return false;

    const hits: Array<{ from: number; to: number; mark: Mark }> = [];
    state.doc.descendants((node, pos) => {
      if (!node.isText) return;
      for (const mark of node.marks) {
        if (mark.type === commentMark && mark.attrs.threadId === threadId) {
          hits.push({ from: pos, to: pos + node.nodeSize, mark });
        }
      }
    });
    if (hits.length === 0) return false;

    if (dispatch) {
      let tr = state.tr;
      for (const hit of hits) {
        tr = tr.removeMark(hit.from, hit.to, hit.mark);
      }
      dispatch(tr);
    }
    return true;
  };
}

/**
 * Command to insert a table with specified rows and columns
 * 
 * If the table is inserted at the end of the document, an empty paragraph
 * is added after it to ensure there's always a valid cursor position.
 */
export function insertTable(schema: Schema, rows: number, cols: number): Command {
  return (state, dispatch) => {
    const tableType = schema.nodes.table;
    const rowType = schema.nodes.table_row;
    const cellType = schema.nodes.table_cell;
    const paragraphType = schema.nodes.paragraph;
    
    if (!tableType || !rowType || !cellType || !paragraphType) return false;
    
    if (dispatch) {
      // Create table structure
      const tableRows = [];
      for (let r = 0; r < rows; r++) {
        const cells = [];
        for (let c = 0; c < cols; c++) {
          // Each cell contains an empty paragraph
          cells.push(cellType.create(null, paragraphType.create()));
        }
        tableRows.push(rowType.create(null, cells));
      }
      const table = tableType.create(null, tableRows);
      
      // Check if we're inserting at the end of the document
      const { $to } = state.selection;
      const isAtEnd = $to.pos >= state.doc.content.size - 1;
      
      let tr = state.tr.replaceSelectionWith(table);
      
      // If at the end, add an empty paragraph after the table
      // This ensures there's always a valid cursor position after the table
      if (isAtEnd) {
        const insertPos = tr.doc.content.size;
        tr = tr.insert(insertPos, paragraphType.create());
        
        // Set cursor to the new paragraph
        const newPos = insertPos + 1; // Inside the new paragraph
        tr = tr.setSelection(TextSelection.create(tr.doc, newPos));
      }
      
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/**
 * Command to set cell background color for a table cell at given position range
 * 
 * @param pmStart - Position at the start of the cell (before cell node)
 * @param pmEnd - Position at the end of the cell (after cell node)
 */
export function setCellBackgroundColor(
  schema: Schema, 
  pmStart: number, 
  _pmEnd: number, 
  color: string | null
): Command {
  return (state, dispatch) => {
    const tableCellType = schema.nodes.table_cell;
    const tableHeaderType = schema.nodes.table_header;
    
    if (!tableCellType && !tableHeaderType) return false;
    
    // Resolve position INSIDE the cell (pmStart + 1 to enter the cell node)
    // pmStart is at the cell boundary, pmStart + 1 is inside the cell
    const insidePos = Math.min(pmStart + 1, state.doc.content.size);
    const $pos = state.doc.resolve(insidePos);
    
    // Walk up to find the table_cell or table_header node
    for (let d = $pos.depth; d >= 0; d--) {
      const node = $pos.node(d);
      if ((tableCellType && node.type === tableCellType) ||
          (tableHeaderType && node.type === tableHeaderType)) {
        const cellPos = $pos.before(d);
        
        if (dispatch) {
          const tr = state.tr.setNodeMarkup(cellPos, null, {
            ...node.attrs,
            backgroundColor: color,
          });
          dispatch(tr);
        }
        return true;
      }
    }
    
    return false;
  };
}

/**
 * Create a commands object for the given schema
 */
export function createCommands(schema: Schema) {
  return {
    // Inline formatting
    toggleBold: () => toggleMarkCommand(schema.marks.bold),
    toggleItalic: () => toggleMarkCommand(schema.marks.italic),
    toggleUnderline: () => toggleMarkCommand(schema.marks.underline),
    toggleStrikethrough: () => toggleMarkCommand(schema.marks.strikethrough),
    toggleSuperscript: () => toggleMarkCommand(schema.marks.superscript),
    toggleSubscript: () => toggleMarkCommand(schema.marks.subscript),
    
    // Text style
    setTextColor: (color: string) => setMarkAttrs(schema.marks.textStyle, { color }),
    setHighlight: (color: string) => setMarkAttrs(schema.marks.textStyle, { backgroundColor: color }),
    setFontSize: (size: number) => setMarkAttrs(schema.marks.textStyle, { fontSize: size }),
    setFontFamily: (family: string) => setMarkAttrs(schema.marks.textStyle, { fontFamily: family }),
    clearTextStyle: () => removeMarkCommand(schema.marks.textStyle),
    
    // Block formatting
    setParagraph: () => setBlockTypeCommand(schema.nodes.paragraph),
    setHeading: (level: number) => setBlockTypeCommand(schema.nodes.heading, { level }),
    
    // Lists
    toggleBulletList: () => toggleBulletList(schema),
    toggleOrderedList: () => toggleOrderedList(schema),
    increaseIndent: () => increaseIndent(schema),
    decreaseIndent: () => decreaseIndent(schema),
    
    // Alignment
    alignLeft: () => setAlignment(schema, 'left'),
    alignCenter: () => setAlignment(schema, 'center'),
    alignRight: () => setAlignment(schema, 'right'),
    alignJustify: () => setAlignment(schema, 'justify'),
    
    // Insert
    insertHorizontalRule: () => insertHorizontalRule(schema),
    insertPageBreak: () => insertPageBreak(schema),
    insertImage: (attrs: { src: string; alt?: string; width?: number; height?: number }) => insertImage(schema, attrs),
    insertLink: (attrs: { href: string; title?: string }) => insertLink(schema, attrs),
    removeLink: () => removeLink(schema),
    addComment: (threadId: string) => addComment(schema, threadId),
    removeComment: (threadId: string) => removeComment(schema, threadId),
    insertTable: (rows: number, cols: number) => insertTable(schema, rows, cols),
    
    // Cell formatting
    setCellBackgroundColor: (pmStart: number, pmEnd: number, color: string | null) => 
      setCellBackgroundColor(schema, pmStart, pmEnd, color),
  };
}

export type DocsCommands = ReturnType<typeof createCommands>;

