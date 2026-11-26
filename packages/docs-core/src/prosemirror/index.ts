// ProseMirror integration for docs-core

// Schema
export { docsSchema, nodeSpecs, markSpecs } from './schema';

// Plugins
export { 
  createPlugins, 
  buildInputRules, 
  buildKeymap,
  createActiveMarksPlugin,
  activeMarksPluginKey,
  undo,
  redo,
} from './plugins';
export type { ActiveState, BlockState } from './plugins';
export { newRevisionId, freshRevisionIds, typeOutsideDeletion, pasteOutsideDeletion, insertTrackedLineBreak, createRevisionEditPlugin } from './revision-edit';

// Commands
export { 
  createCommands,
  toggleMarkCommand,
  setMarkAttrs,
  removeMarkCommand,
  setBlockTypeCommand,
  toggleBulletList,
  toggleOrderedList,
  increaseIndent,
  decreaseIndent,
  setAlignment,
  insertHorizontalRule,
  insertPageBreak,
  insertImage,
  insertLink,
  removeLink,
  insertTable,
  goToNextCell,
  toggleCheckboxControl,
  findCheckboxControl,
  CHECKBOX_CHECKED_GLYPH,
  CHECKBOX_UNCHECKED_GLYPH,
  type DocsCommands,
} from './commands';

// State synchronization
export {
  documentToProseMirror,
  sectionToProseMirror,
  proseMirrorToDocument,
  createPmDoc,
  blocksToPmDoc,
  sectionsToPmDoc,
  getPmPlainText,
  isDocEmpty,
} from './sync';


// Footnote and endnote references
export {
  findNoteRef,
  noteOrdinalAt,
  nextNoteId,
  insertNoteReference,
  noteTextFragment,
  type NoteKind,
} from './notes';

// Table of contents
export {
  DEFAULT_TOC_INSTRUCTION,
  EMPTY_TOC_TEXT,
  parseTocInstruction,
  insertTableOfContents,
  updateTableOfContents,
  removeTableOfContents,
  type TocSettings,
} from './toc';
