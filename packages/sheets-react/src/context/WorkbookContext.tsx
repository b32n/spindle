import React, { createContext, useContext, useMemo, useState, useCallback, useEffect, useRef } from 'react';
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';
import type { CommentAuthor, SheetCommentEvent } from '@b32nio/spindle-sheets-core';

interface WorkbookContextValue {
  workbook: WorkbookImpl;
  updateWorkbook: (updater: (wb: WorkbookImpl) => void) => void;
  /** Identity attributed to comments created in this session. */
  currentUser: CommentAuthor;
  /** Users that can be @-mentioned in comments. */
  mentionableUsers: CommentAuthor[];
}

const WorkbookContext = createContext<WorkbookContextValue | undefined>(undefined);

const DEFAULT_USER: CommentAuthor = { id: 'local-user', name: 'You' };
const NO_USERS: CommentAuthor[] = [];

export function WorkbookProvider({
  workbook: initialWorkbook,
  currentUser,
  mentionableUsers,
  onCommentEvent,
  children,
}: {
  workbook: WorkbookImpl;
  /** Author for comments. Falls back to a generic local user when omitted. */
  currentUser?: CommentAuthor;
  /** Users that can be @-mentioned. Pass a stable reference. */
  mentionableUsers?: CommentAuthor[];
  /**
   * Called for the local user's comment actions (create / reply / resolve /
   * etc.) — the hook for sending notifications. Not fired for threads that
   * arrive from collaborators.
   */
  onCommentEvent?: (event: SheetCommentEvent) => void;
  children: React.ReactNode;
}) {
  const [workbook] = useState<WorkbookImpl>(initialWorkbook);
  const [updateTrigger, setUpdateTrigger] = useState(0);

  // Stable even if the app passes an inline currentUser object.
  const resolvedUser = useMemo<CommentAuthor>(
    () => currentUser ?? DEFAULT_USER,
    [currentUser?.id, currentUser?.name],
  );
  const resolvedMentionables = mentionableUsers ?? NO_USERS;

  // Forward semantic comment events to the host. A ref keeps the listener
  // stable even if the app passes an inline onCommentEvent callback.
  const onCommentEventRef = useRef(onCommentEvent);
  onCommentEventRef.current = onCommentEvent;
  useEffect(() => {
    return workbook.on('commentEvent', (data) => {
      onCommentEventRef.current?.(data.payload as SheetCommentEvent);
    });
  }, [workbook]);

  // Subscribe to workbook events to trigger re-renders
  useEffect(() => {
    const handleCellChange = () => {
      setUpdateTrigger((prev) => prev + 1);
    };

    const handleSheetChange = () => {
      setUpdateTrigger((prev) => prev + 1);
    };

    const handleSheetAdd = () => {
      setUpdateTrigger((prev) => prev + 1);
    };

    const handleSheetDelete = () => {
      setUpdateTrigger((prev) => prev + 1);
    };

    workbook.on('cellChange', handleCellChange);
    workbook.on('sheetChange', handleSheetChange);
    workbook.on('sheetAdd', handleSheetAdd);
    workbook.on('sheetDelete', handleSheetDelete);

    return () => {
      workbook.off('cellChange', handleCellChange);
      workbook.off('sheetChange', handleSheetChange);
      workbook.off('sheetAdd', handleSheetAdd);
      workbook.off('sheetDelete', handleSheetDelete);
    };
  }, [workbook]);

  const updateWorkbook = useCallback((updater: (wb: WorkbookImpl) => void) => {
    updater(workbook);
    setUpdateTrigger((prev) => prev + 1);
  }, [workbook]);

  const value = useMemo(
    () => ({
      workbook,
      updateWorkbook,
      currentUser: resolvedUser,
      mentionableUsers: resolvedMentionables,
    }),
    [workbook, updateWorkbook, updateTrigger, resolvedUser, resolvedMentionables]
  );

  return <WorkbookContext.Provider value={value}>{children}</WorkbookContext.Provider>;
}

export function useWorkbook(): WorkbookContextValue {
  const context = useContext(WorkbookContext);
  if (!context) {
    throw new Error('useWorkbook must be used within WorkbookProvider');
  }
  return context;
}

