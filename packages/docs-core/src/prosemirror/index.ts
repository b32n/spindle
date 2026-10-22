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
  type DocsCommands,
} from './commands';

// State synchronization
export {
  documentToProseMirror,
  sectionToProseMirror,
  proseMirrorToDocument,
  createPmDoc,
  blocksToPmDoc,
  getPmPlainText,
  isDocEmpty,
} from './sync';

