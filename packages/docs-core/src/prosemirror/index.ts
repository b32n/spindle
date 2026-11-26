// ProseMirror integration for docs-core

// Schema
export { docsSchema, nodeSpecs, markSpecs, sanitizeRunFormat } from './schema';

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
export { BLOCK_REVISION_ATTRS, BLOCK_REVISION_ATTR_OF, blockRevisionAttrs, readBlockRevision, withBlockRevision, withBlockRevisionKind } from './block-revision';
export type { BlockRevisionAttrs, BlockRevisionKind } from './block-revision';
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

// Reviewing tracked changes: what can be accepted or rejected, and the commands that do it
export { changeUnits, listChanges, settle, resolveChange, resolveAll, resolveSelection } from './review';
export type { ChangeUnit, ChangeUnitKind, ReviewChange } from './review';
