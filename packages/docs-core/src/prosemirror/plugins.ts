import { keymap } from 'prosemirror-keymap';
import { history, undo, redo } from 'prosemirror-history';
import { baseKeymap, toggleMark, chainCommands, exitCode, joinUp, joinDown, lift, selectParentNode } from 'prosemirror-commands';
import { 
  inputRules, 
  wrappingInputRule, 
  textblockTypeInputRule,
  smartQuotes,
  emDash,
  ellipsis,
  InputRule,
} from 'prosemirror-inputrules';
import { Plugin, PluginKey, type Command } from 'prosemirror-state';
import { goToNextCell } from './commands';
import type { Schema } from 'prosemirror-model';
import { liftListItem, sinkListItem, splitListItem } from 'prosemirror-schema-list';

/**
 * Create heading input rules: # Heading 1, ## Heading 2, etc.
 */
function headingRule(schema: Schema, level: number): InputRule {
  return textblockTypeInputRule(
    new RegExp(`^(#{1,${level}})\\s$`),
    schema.nodes.heading,
    (match) => ({ level: match[1].length })
  );
}

/**
 * Create bullet list input rule: - item, * item, + item
 */
function bulletListRule(schema: Schema): InputRule {
  return wrappingInputRule(/^\s*([-+*])\s$/, schema.nodes.bullet_list);
}

/**
 * Create ordered list input rule: 1. item, 2. item, etc.
 */
function orderedListRule(schema: Schema): InputRule {
  return wrappingInputRule(
    /^\s*(\d+)\.\s$/,
    schema.nodes.ordered_list,
    (match) => ({ start: +match[1] }),
    (match, node) => node.childCount + node.attrs.start === +match[1]
  );
}

/**
 * Create horizontal rule input rule: --- or ***
 */
function horizontalRuleRule(schema: Schema): InputRule {
  return new InputRule(/^(?:---|\*\*\*)$/, (state, _match, start, end) => {
    const { tr } = state;
    tr.replaceWith(start - 1, end, schema.nodes.horizontal_rule.create());
    return tr;
  });
}

/**
 * Build input rules for the schema
 */
export function buildInputRules(schema: Schema): Plugin {
  const rules: InputRule[] = [
    ...smartQuotes,
    ellipsis,
    emDash,
  ];

  // Heading rules (h1-h6)
  if (schema.nodes.heading) {
    for (let i = 1; i <= 6; i++) {
      rules.push(headingRule(schema, i));
    }
  }

  // List rules
  if (schema.nodes.bullet_list) {
    rules.push(bulletListRule(schema));
  }
  if (schema.nodes.ordered_list) {
    rules.push(orderedListRule(schema));
  }

  // Horizontal rule
  if (schema.nodes.horizontal_rule) {
    rules.push(horizontalRuleRule(schema));
  }

  return inputRules({ rules });
}

/**
 * Build the keymap for the schema
 */
export function buildKeymap(schema: Schema): Plugin {
  const keys: Record<string, Command> = {};

  // Undo/Redo
  keys['Mod-z'] = undo;
  keys['Mod-y'] = redo;
  keys['Mod-Shift-z'] = redo;

  // Text formatting
  if (schema.marks.bold) {
    keys['Mod-b'] = toggleMark(schema.marks.bold, null, { removeWhenPresent: false });
    keys['Mod-B'] = toggleMark(schema.marks.bold, null, { removeWhenPresent: false });
  }
  if (schema.marks.italic) {
    keys['Mod-i'] = toggleMark(schema.marks.italic, null, { removeWhenPresent: false });
    keys['Mod-I'] = toggleMark(schema.marks.italic, null, { removeWhenPresent: false });
  }
  if (schema.marks.underline) {
    keys['Mod-u'] = toggleMark(schema.marks.underline, null, { removeWhenPresent: false });
    keys['Mod-U'] = toggleMark(schema.marks.underline, null, { removeWhenPresent: false });
  }
  if (schema.marks.strikethrough) {
    keys['Mod-Shift-s'] = toggleMark(schema.marks.strikethrough, null, { removeWhenPresent: false });
    keys['Mod-Shift-S'] = toggleMark(schema.marks.strikethrough, null, { removeWhenPresent: false });
  }
  
  // Link shortcut - Mod-K
  // Note: This just prevents default browser behavior (Mod-K opens search in some browsers)
  // The actual dialog opening is handled by the toolbar via InputBridge forwarding
  if (schema.marks.link) {
    keys['Mod-k'] = () => true;  // Return true to indicate handled, let Toolbar handle dialog
    keys['Mod-K'] = () => true;
  }

  // Hard break
  if (schema.nodes.hard_break) {
    const br = schema.nodes.hard_break;
    const cmd = chainCommands(exitCode, (state, dispatch) => {
      if (dispatch) {
        dispatch(state.tr.replaceSelectionWith(br.create()).scrollIntoView());
      }
      return true;
    });
    keys['Mod-Enter'] = cmd;
    keys['Shift-Enter'] = cmd;
  }

  // List handling
  if (schema.nodes.list_item) {
    keys['Enter'] = splitListItem(schema.nodes.list_item);
    keys['Tab'] = sinkListItem(schema.nodes.list_item);
    keys['Shift-Tab'] = liftListItem(schema.nodes.list_item);
  }

  // Tab moves between table cells (Tab in the last cell adds a row). Inside a
  // list item the list commands above win.
  if (schema.nodes.table) {
    keys['Tab'] = keys['Tab'] ? chainCommands(keys['Tab'], goToNextCell(1)) : goToNextCell(1);
    keys['Shift-Tab'] = keys['Shift-Tab'] ? chainCommands(keys['Shift-Tab'], goToNextCell(-1)) : goToNextCell(-1);
  }

  // Structure commands
  keys['Alt-ArrowUp'] = joinUp;
  keys['Alt-ArrowDown'] = joinDown;
  keys['Mod-BracketLeft'] = lift;
  keys['Escape'] = selectParentNode;

  return keymap(keys);
}

/**
 * Plugin key for tracking active marks
 */
export const activeMarksPluginKey = new PluginKey('activeMarks');

/**
 * Block-level state tracked by the plugin
 */
export interface BlockState {
  /** Current block type: 'paragraph' | 'heading' */
  blockType: 'paragraph' | 'heading';
  /** Heading level (1-6) if blockType is 'heading', null otherwise */
  headingLevel: number | null;
  /** Current list type: 'bullet_list' | 'ordered_list' | null */
  listType: 'bullet_list' | 'ordered_list' | null;
  /** Current text alignment */
  alignment: 'left' | 'center' | 'right' | 'justify';
}

/**
 * Full active state (marks + block-level)
 */
export interface ActiveState {
  // Mark-level state
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  superscript: boolean;
  subscript: boolean;
  link: { href: string; title?: string } | null;
  textStyle: { color?: string; backgroundColor?: string; fontSize?: number; fontFamily?: string } | null;
  // Block-level state
  blockType: 'paragraph' | 'heading';
  headingLevel: number | null;
  listType: 'bullet_list' | 'ordered_list' | null;
  alignment: 'left' | 'center' | 'right' | 'justify';
}

/**
 * Plugin to track which marks are active at the current selection
 * Also tracks block-level state (heading, list type, alignment)
 */
export function createActiveMarksPlugin(): Plugin {
  return new Plugin({
    key: activeMarksPluginKey,
    state: {
      init(): ActiveState {
        return {
          // Mark-level
          bold: false,
          italic: false,
          underline: false,
          strikethrough: false,
          superscript: false,
          subscript: false,
          link: null,
          textStyle: null,
          // Block-level
          blockType: 'paragraph',
          headingLevel: null,
          listType: null,
          alignment: 'left',
        };
      },
      apply(_tr, _value, _oldState, newState): ActiveState {
        const { from, to, empty } = newState.selection;
        const state: ActiveState = {
          // Mark-level
          bold: false,
          italic: false,
          underline: false,
          strikethrough: false,
          superscript: false,
          subscript: false,
          link: null,
          textStyle: null,
          // Block-level defaults
          blockType: 'paragraph',
          headingLevel: null,
          listType: null,
          alignment: 'left',
        };

        // === Mark-level tracking ===
        if (empty) {
          // At cursor, check stored marks or marks at position
          const storedMarks = newState.storedMarks || newState.doc.resolve(from).marks();
          for (const mark of storedMarks) {
            if (mark.type.name === 'bold') state.bold = true;
            if (mark.type.name === 'italic') state.italic = true;
            if (mark.type.name === 'underline') state.underline = true;
            if (mark.type.name === 'strikethrough') state.strikethrough = true;
            if (mark.type.name === 'superscript') state.superscript = true;
            if (mark.type.name === 'subscript') state.subscript = true;
            if (mark.type.name === 'link') state.link = mark.attrs as { href: string; title?: string };
            if (mark.type.name === 'textStyle') state.textStyle = mark.attrs as { color?: string; backgroundColor?: string; fontSize?: number; fontFamily?: string };
          }
        } else {
          // Selection - check if marks are active across the selection
          newState.doc.nodesBetween(from, to, (node) => {
            if (node.isText && node.marks) {
              for (const mark of node.marks) {
                if (mark.type.name === 'bold') state.bold = true;
                if (mark.type.name === 'italic') state.italic = true;
                if (mark.type.name === 'underline') state.underline = true;
                if (mark.type.name === 'strikethrough') state.strikethrough = true;
                if (mark.type.name === 'superscript') state.superscript = true;
                if (mark.type.name === 'subscript') state.subscript = true;
                if (mark.type.name === 'link') state.link = mark.attrs as { href: string; title?: string };
                if (mark.type.name === 'textStyle') state.textStyle = mark.attrs as { color?: string; backgroundColor?: string; fontSize?: number; fontFamily?: string };
              }
            }
          });
        }

        // === Block-level tracking ===
        // Find the block node at the cursor position
        const $pos = newState.doc.resolve(from);
        
        // Walk up from cursor to find relevant block nodes
        for (let depth = $pos.depth; depth >= 0; depth--) {
          const node = $pos.node(depth);
          const nodeName = node.type.name;
          
          // Check for heading
          if (nodeName === 'heading') {
            state.blockType = 'heading';
            state.headingLevel = node.attrs.level || 1;
            // Get alignment from heading (check both 'alignment' and 'align' for compatibility)
            if (node.attrs.alignment) {
              state.alignment = node.attrs.alignment as 'left' | 'center' | 'right' | 'justify';
            } else if (node.attrs.align) {
              state.alignment = node.attrs.align as 'left' | 'center' | 'right' | 'justify';
            }
          }
          
          // Check for paragraph
          if (nodeName === 'paragraph') {
            state.blockType = 'paragraph';
            // Get alignment from paragraph (check both 'alignment' and 'align' for compatibility)
            if (node.attrs.alignment) {
              state.alignment = node.attrs.alignment as 'left' | 'center' | 'right' | 'justify';
            } else if (node.attrs.align) {
              state.alignment = node.attrs.align as 'left' | 'center' | 'right' | 'justify';
            }
          }
          
          // Check for list
          if (nodeName === 'bullet_list') {
            state.listType = 'bullet_list';
          }
          if (nodeName === 'ordered_list') {
            state.listType = 'ordered_list';
          }
        }

        return state;
      },
    },
  });
}

/**
 * Create all plugins for the editor
 */
export function createPlugins(schema: Schema): Plugin[] {
  return [
    buildInputRules(schema),
    buildKeymap(schema),
    keymap(baseKeymap),
    history(),
    createActiveMarksPlugin(),
  ];
}

export { undo, redo };

