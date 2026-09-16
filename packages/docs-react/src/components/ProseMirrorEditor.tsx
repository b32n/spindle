import React, { useEffect, useRef, useCallback, forwardRef, useImperativeHandle } from 'react';
import { EditorState, Transaction } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { ySyncPlugin, yCursorPlugin } from 'y-prosemirror';
import { ensureCollabCursorStyles } from '../core/collab-cursor-styles';
import {
  docsSchema,
  createPlugins,
  createCommands,
  activeMarksPluginKey,
  type DocsCommands,
  type ActiveState,
} from '@b32nio/spindle-docs-core';
import type { CollabHandle } from '@b32nio/spindle-docs-core/collab';

// ProseMirror CSS (basic styles)
const proseMirrorStyles = `
.ProseMirror {
  outline: none;
  white-space: pre-wrap;
  word-wrap: break-word;
  font-family: Arial, sans-serif;
  font-size: 11pt;
  line-height: 1.5;
  color: #202124;
}

.ProseMirror p {
  margin: 0;
  padding: 2px 0;
  min-height: 1.4em;
}

.ProseMirror h1 { font-size: 24pt; font-weight: bold; margin: 0.5em 0 0.3em; }
.ProseMirror h2 { font-size: 20pt; font-weight: bold; margin: 0.5em 0 0.3em; }
.ProseMirror h3 { font-size: 16pt; font-weight: bold; margin: 0.5em 0 0.3em; }
.ProseMirror h4 { font-size: 14pt; font-weight: bold; margin: 0.5em 0 0.3em; }
.ProseMirror h5 { font-size: 12pt; font-weight: bold; margin: 0.5em 0 0.3em; }
.ProseMirror h6 { font-size: 11pt; font-weight: bold; margin: 0.5em 0 0.3em; }

.ProseMirror ul, .ProseMirror ol {
  margin: 0.5em 0;
  padding-left: 24px;
}

.ProseMirror li {
  margin-bottom: 0.25em;
}

.ProseMirror li > p {
  margin: 0;
}

.ProseMirror table {
  width: 100%;
  border-collapse: collapse;
  margin: 1em 0;
}

.ProseMirror td, .ProseMirror th {
  border: 1px solid #e8eaed;
  padding: 8px;
  vertical-align: top;
  min-width: 50px;
}

.ProseMirror th {
  background-color: #f8f9fa;
  font-weight: bold;
}

.ProseMirror img {
  max-width: 100%;
  height: auto;
}

.ProseMirror hr {
  border: none;
  border-top: 1px solid #e8eaed;
  margin: 1em 0;
}

.ProseMirror .page-break {
  border-top: 1px dashed #5f6368;
  margin: 1em 0;
  position: relative;
}

.ProseMirror .page-break::after {
  content: 'Page Break';
  position: absolute;
  top: -10px;
  left: 50%;
  transform: translateX(-50%);
  background-color: #fff;
  padding: 0 8px;
  font-size: 10px;
  color: #5f6368;
}

.ProseMirror a {
  color: #1a73e8;
  text-decoration: underline;
}

.ProseMirror-selectednode {
  outline: 2px solid #1a73e8;
}

/* Placeholder for empty paragraphs */
.ProseMirror p.is-empty::before {
  content: attr(data-placeholder);
  color: #9aa0a6;
  pointer-events: none;
  position: absolute;
}

/* Selection styling */
.ProseMirror ::selection {
  background-color: #b4d5fe;
}
`;

/**
 * ActiveMarks extends ActiveState from docs-core
 * Includes both mark-level and block-level state
 */
export interface ActiveMarks extends ActiveState {}

export interface ProseMirrorEditorProps {
  initialContent?: unknown;
  onChange?: (state: EditorState) => void;
  onSelectionChange?: (state: EditorState, activeMarks: ActiveMarks) => void;
  onFocus?: () => void;
  onBlur?: () => void;
  className?: string;
  style?: React.CSSProperties;
  placeholder?: string;
  editable?: boolean;
  autoFocus?: boolean;
  /**
   * If present, the editor binds to the collab handle's Y.XmlFragment via
   * ySyncPlugin instead of initializing from `initialContent`. The handle
   * becomes the source of truth for body content; local edits propagate to
   * peers and remote edits land in the editor view.
   */
  collabHandle?: CollabHandle | null;
}

export interface ProseMirrorEditorRef {
  view: EditorView | null;
  commands: DocsCommands;
  focus: () => void;
  blur: () => void;
  getState: () => EditorState | null;
  getContent: () => unknown;
  runCommand: (command: ReturnType<DocsCommands[keyof DocsCommands]>) => boolean;
}

/**
 * React wrapper component for ProseMirror editor
 */
export const ProseMirrorEditor = forwardRef<ProseMirrorEditorRef, ProseMirrorEditorProps>(
  function ProseMirrorEditor(
    {
      initialContent,
      onChange,
      onSelectionChange,
      onFocus,
      onBlur,
      className,
      style,
      placeholder = 'Type something...',
      editable = true,
      autoFocus = false,
      collabHandle = null,
    },
    ref
  ) {
    const editorRef = useRef<HTMLDivElement>(null);
    const viewRef = useRef<EditorView | null>(null);
    const commandsRef = useRef<DocsCommands>(createCommands(docsSchema));
    
    // Inject styles once
    useEffect(() => {
      const styleId = 'prosemirror-docs-styles';
      if (!document.getElementById(styleId)) {
        const styleEl = document.createElement('style');
        styleEl.id = styleId;
        styleEl.textContent = proseMirrorStyles;
        document.head.appendChild(styleEl);
      }
    }, []);
    
    // Initialize editor
    useEffect(() => {
      if (!editorRef.current) return;

      // Two initialization paths:
      //   - Collab: ySyncPlugin owns the doc; it's seeded from the Y.XmlFragment
      //     on the handle. initialContent is ignored — the Y.Doc is the truth.
      //   - Local: build a PM doc from initialContent (or an empty paragraph).
      const plugins = createPlugins(docsSchema);
      let stateConfig: Parameters<typeof EditorState.create>[0];
      if (collabHandle) {
        ensureCollabCursorStyles();
        plugins.unshift(
          ySyncPlugin(collabHandle.xmlFragment),
          yCursorPlugin(collabHandle.awareness),
        );
        stateConfig = { schema: docsSchema, plugins };
      } else {
        let doc;
        try {
          doc = initialContent
            ? docsSchema.nodeFromJSON(initialContent)
            : docsSchema.node('doc', null, [docsSchema.node('paragraph')]);
        } catch (e) {
          console.error('Failed to parse initial content:', e);
          doc = docsSchema.node('doc', null, [docsSchema.node('paragraph')]);
        }
        stateConfig = { doc, plugins };
      }

      const state = EditorState.create(stateConfig);
      
      // `let view` + `view ?? this` fallback avoids a temporal-dead-zone
      // crash when ySyncPlugin dispatches its first sync transaction inside
      // the EditorView constructor — at that moment the outer `view` binding
      // hasn't been assigned yet.
      let view: EditorView | undefined;
      // eslint-disable-next-line prefer-const -- `const` would TDZ-crash; see above
      view = new EditorView(editorRef.current, {
        state,
        editable: () => editable,
        dispatchTransaction(transaction: Transaction) {
          // eslint-disable-next-line @typescript-eslint/no-this-alias
          const v = view ?? (this as unknown as EditorView);
          const newState = v.state.apply(transaction);
          v.updateState(newState);

          if (transaction.docChanged && onChange) {
            onChange(newState);
          }

          if ((transaction.selectionSet || transaction.docChanged) && onSelectionChange) {
            const activeMarks = activeMarksPluginKey.getState(newState) as ActiveMarks;
            onSelectionChange(newState, activeMarks);
          }
        },
        handleDOMEvents: {
          focus: () => {
            onFocus?.();
            return false;
          },
          blur: () => {
            onBlur?.();
            return false;
          },
        },
        attributes: {
          'data-placeholder': placeholder,
        },
      });
      
      viewRef.current = view;
      
      if (autoFocus) {
        setTimeout(() => view.focus(), 0);
      }
      
      return () => {
        view.destroy();
        viewRef.current = null;
      };
      // collabHandle in the deps so swapping in/out of collab mode rebuilds
      // the editor with the right plugin stack. eslint-disable-next-line:
      // initialContent intentionally excluded — it's seed data, not reactive.
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [collabHandle]);
    
    // Update editable state
    useEffect(() => {
      if (viewRef.current) {
        viewRef.current.setProps({ editable: () => editable });
      }
    }, [editable]);
    
    // Expose methods via ref
    useImperativeHandle(ref, () => ({
      view: viewRef.current,
      commands: commandsRef.current,
      focus: () => viewRef.current?.focus(),
      blur: () => viewRef.current?.dom.blur(),
      getState: () => viewRef.current?.state ?? null,
      getContent: () => viewRef.current?.state.doc.toJSON() ?? null,
      runCommand: (command) => {
        if (!viewRef.current) return false;
        return command(viewRef.current.state, viewRef.current.dispatch, viewRef.current);
      },
    }), []);
    
    return (
      <div
        ref={editorRef}
        className={className}
        style={{
          ...style,
          outline: 'none',
          minHeight: '1em',
        }}
      />
    );
  }
);

/**
 * Hook to use ProseMirror commands
 */
export function useProseMirrorCommands(editorRef: React.RefObject<ProseMirrorEditorRef>) {
  const runCommand = useCallback(
    (command: ReturnType<DocsCommands[keyof DocsCommands]>) => {
      return editorRef.current?.runCommand(command) ?? false;
    },
    [editorRef]
  );
  
  return {
    runCommand,
    toggleBold: () => runCommand(createCommands(docsSchema).toggleBold()),
    toggleItalic: () => runCommand(createCommands(docsSchema).toggleItalic()),
    toggleUnderline: () => runCommand(createCommands(docsSchema).toggleUnderline()),
    toggleStrikethrough: () => runCommand(createCommands(docsSchema).toggleStrikethrough()),
    setHeading: (level: number) => runCommand(createCommands(docsSchema).setHeading(level)),
    setParagraph: () => runCommand(createCommands(docsSchema).setParagraph()),
    toggleBulletList: () => runCommand(createCommands(docsSchema).toggleBulletList()),
    toggleOrderedList: () => runCommand(createCommands(docsSchema).toggleOrderedList()),
    alignLeft: () => runCommand(createCommands(docsSchema).alignLeft()),
    alignCenter: () => runCommand(createCommands(docsSchema).alignCenter()),
    alignRight: () => runCommand(createCommands(docsSchema).alignRight()),
    alignJustify: () => runCommand(createCommands(docsSchema).alignJustify()),
    setTextColor: (color: string) => runCommand(createCommands(docsSchema).setTextColor(color)),
    setHighlight: (color: string) => runCommand(createCommands(docsSchema).setHighlight(color)),
    setFontSize: (size: number) => runCommand(createCommands(docsSchema).setFontSize(size)),
    setFontFamily: (family: string) => runCommand(createCommands(docsSchema).setFontFamily(family)),
  };
}

