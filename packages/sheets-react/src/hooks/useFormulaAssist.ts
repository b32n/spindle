import { useCallback, useEffect, useMemo, useState } from 'react';
import type { KeyboardEvent } from 'react';
import { analyzeFormula, searchFunctions, getFunctionDoc } from '@b32nio/spindle-sheets-core';
import type { FunctionDoc } from '@b32nio/spindle-sheets-core';

export type FormulaAssistMode = 'autocomplete' | 'signature' | 'hidden';

export interface FormulaSignature {
  doc: FunctionDoc;
  activeArg: number;
}

export interface UseFormulaAssistOptions {
  /** The formula text being edited. */
  value: string;
  /** Caret offset within the text. */
  caret: number;
  /** Whether a formula is being edited at all (the result is `hidden` if not). */
  enabled: boolean;
  /** Applies an accepted suggestion — the consumer updates its editor state. */
  onAccept: (value: string, caret: number) => void;
}

export interface FormulaAssist {
  mode: FormulaAssistMode;
  suggestions: FunctionDoc[];
  highlightedIndex: number;
  setHighlightedIndex: (index: number) => void;
  signature: FormulaSignature | null;
  /** Returns true when it consumed the key (the editor should then ignore it). */
  onKeyDown: (e: KeyboardEvent) => boolean;
  /** Accept a suggestion by index — used by mouse clicks. */
  accept: (index: number) => void;
}

/**
 * Drives formula autocomplete and parameter help. Headless — it computes what
 * to show and how keys behave; the consumer renders the popups and applies
 * accepted suggestions. Reused by both the in-cell editor and the formula bar.
 */
export function useFormulaAssist(options: UseFormulaAssistOptions): FormulaAssist {
  const { value, caret, enabled, onAccept } = options;
  const [highlightedIndex, setHighlightedIndex] = useState(0);
  const [dismissed, setDismissed] = useState(false);

  const context = useMemo(
    () => (enabled ? analyzeFormula(value, caret) : {}),
    [enabled, value, caret]
  );

  const suggestions = useMemo<FunctionDoc[]>(
    () => (context.token ? searchFunctions(context.token.text) : []),
    [context]
  );

  const signature = useMemo<FormulaSignature | null>(() => {
    if (!context.call) return null;
    const doc = getFunctionDoc(context.call.name);
    return doc ? { doc, activeArg: context.call.argIndex } : null;
  }, [context]);

  // Typing clears a previous Escape-dismiss; a new suggestion list resets the cursor.
  useEffect(() => setDismissed(false), [value]);
  useEffect(() => setHighlightedIndex(0), [suggestions]);

  let mode: FormulaAssistMode = 'hidden';
  if (enabled && !dismissed) {
    if (suggestions.length > 0) mode = 'autocomplete';
    else if (signature) mode = 'signature';
  }

  const accept = useCallback(
    (index: number) => {
      const token = context.token;
      const chosen = suggestions[index];
      if (!token || !chosen) return;
      // Replace the partial identifier with `NAME(`, caret just after the `(`.
      const next = value.slice(0, token.start) + chosen.name + '(' + value.slice(token.end);
      onAccept(next, token.start + chosen.name.length + 1);
    },
    [context, suggestions, value, onAccept]
  );

  const onKeyDown = useCallback(
    (e: KeyboardEvent): boolean => {
      if (mode === 'autocomplete') {
        if (e.key === 'ArrowDown') {
          setHighlightedIndex((i) => Math.min(i + 1, suggestions.length - 1));
          e.preventDefault();
          return true;
        }
        if (e.key === 'ArrowUp') {
          setHighlightedIndex((i) => Math.max(i - 1, 0));
          e.preventDefault();
          return true;
        }
        if (e.key === 'Enter' || e.key === 'Tab') {
          accept(highlightedIndex);
          e.preventDefault();
          return true;
        }
        if (e.key === 'Escape') {
          setDismissed(true);
          e.preventDefault();
          return true;
        }
      } else if (mode === 'signature' && e.key === 'Escape') {
        setDismissed(true);
        e.preventDefault();
        return true;
      }
      return false;
    },
    [mode, suggestions.length, accept, highlightedIndex]
  );

  return { mode, suggestions, highlightedIndex, setHighlightedIndex, signature, onKeyDown, accept };
}
