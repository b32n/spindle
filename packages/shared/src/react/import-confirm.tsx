import React, { useCallback, useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { describeImportLoss, type ImportLossReport } from '../export/loss';

/**
 * How an editor's "Open" reacts to a file whose import would lose something:
 * `'dialog'` (default) asks with the ImportConfirmDialog, `'never'` imports without asking,
 * or pass a function (host UI) that resolves to `true` to import or `false` to cancel.
 * Nothing is asked when the report is empty.
 */
export type ImportConfirmation = 'dialog' | 'never' | ((report: ImportLossReport, file: File) => boolean | Promise<boolean>);

export interface ImportConfirmDialogProps {
  fileName: string;
  report: ImportLossReport;
  onConfirm: () => void;
  onCancel: () => void;
}

/**
 * "This file contains things Spindle can't import" — lists the lost features and offers
 * Import anyway / Cancel. Cancel is focused first and Escape cancels, so a stray Enter keeps the open document.
 */
export function ImportConfirmDialog({ fileName, report, onConfirm, onCancel }: ImportConfirmDialogProps): React.ReactElement {
  const cancelRef = useRef<HTMLButtonElement>(null);
  const lines = describeImportLoss(report);

  useEffect(() => {
    cancelRef.current?.focus();
  }, []);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        onCancel();
      }
    };
    document.addEventListener('keydown', onKey, true);
    return () => document.removeEventListener('keydown', onKey, true);
  }, [onCancel]);

  const button: React.CSSProperties = {
    padding: '9px 16px', fontSize: 14, fontWeight: 500, borderRadius: 4, cursor: 'pointer',
  };
  return createPortal(
    <div
      style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(0,0,0,0.3)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 10001 }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby="spindle-import-confirm-title"
        aria-describedby="spindle-import-confirm-desc"
        style={{ background: '#fff', borderRadius: 8, boxShadow: '0 4px 24px rgba(0,0,0,0.2)', width: 520, maxWidth: '90vw', maxHeight: '90vh', overflow: 'auto', padding: '20px 24px 24px', fontFamily: 'inherit', color: '#202124' }}
      >
        <h2 id="spindle-import-confirm-title" style={{ margin: 0, fontSize: 18, fontWeight: 500 }}>
          Import “{fileName}”?
        </h2>
        <p id="spindle-import-confirm-desc" style={{ margin: '12px 0 8px', fontSize: 14, color: '#3c4043' }}>
          This file contains things that can’t be imported. The current content will be replaced, and the following will be lost:
        </p>
        <ul data-testid="import-loss-list" style={{ margin: '0 0 16px', paddingLeft: 20, fontSize: 14, color: '#3c4043' }}>
          {lines.map((line, i) => (
            <li key={report.items[i]?.id ?? i} style={{ margin: '4px 0' }}>{line}</li>
          ))}
        </ul>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
          <button ref={cancelRef} type="button" onClick={onCancel} style={{ ...button, color: '#5f6368', background: 'transparent', border: '1px solid #dadce0' }}>
            Cancel
          </button>
          <button type="button" onClick={onConfirm} style={{ ...button, color: '#fff', background: '#1a73e8', border: 'none' }}>
            Import anyway
          </button>
        </div>
      </div>
    </div>,
    document.body
  );
}

/**
 * Wires an editor's Open flow to the confirmation policy. `confirm(file, report)` resolves
 * `true` immediately when nothing would be lost (or the mode is `'never'`), otherwise asks via
 * the dialog (render `dialog` somewhere in the tree) or the host callback.
 */
export function useImportConfirm(mode: ImportConfirmation = 'dialog'): {
  confirm: (file: File, report: ImportLossReport) => Promise<boolean>;
  dialog: React.ReactElement | null;
} {
  const [pending, setPending] = useState<{ file: File; report: ImportLossReport; resolve: (ok: boolean) => void } | null>(null);
  const pendingRef = useRef(pending);
  pendingRef.current = pending;

  // A dialog still open when the editor unmounts must not leave its Open promise hanging.
  useEffect(() => () => pendingRef.current?.resolve(false), []);

  const confirm = useCallback(
    async (file: File, report: ImportLossReport): Promise<boolean> => {
      if (report.items.length === 0 || mode === 'never') return true;
      if (typeof mode === 'function') return Boolean(await mode(report, file));
      return new Promise<boolean>((resolve) => setPending({ file, report, resolve }));
    },
    [mode]
  );

  const settle = (ok: boolean) => {
    pending?.resolve(ok);
    setPending(null);
  };
  const dialog = pending ? (
    <ImportConfirmDialog fileName={pending.file.name} report={pending.report} onConfirm={() => settle(true)} onCancel={() => settle(false)} />
  ) : null;
  return { confirm, dialog };
}
