// Open / Download controls for the toolbar (opt-in via the editor's `fileActions`). Inline they are plain toolbar buttons; inside
// the ResponsiveToolbar overflow menu (useToolbarMenu) they become rows, like the same controls of the docs and sheets toolbars.

import React from 'react';
import { FolderOpen, Download } from 'lucide-react';
import { useToolbarMenu, MenuRow } from '@b32nio/spindle-shared/react';
import { SLIDES_EXPORT_FORMATS, type SlidesExportFormat } from '../file-io';
import { ToolbarButton as IconButton, ToolbarDivider } from './toolbarUI';

export interface FileActionsProps {
  onOpen: () => void;
  onDownload: (format: SlidesExportFormat) => void;
}

export function FileActions({ onOpen, onDownload }: FileActionsProps): React.ReactElement {
  const menu = useToolbarMenu();
  const format = SLIDES_EXPORT_FORMATS[0];
  if (menu.inMenu) {
    return (
      <>
        <MenuRow icon={<FolderOpen size={17} />} label="Open file…" onClick={() => { onOpen(); menu.closeMenu(); }} />
        <MenuRow icon={<Download size={17} />} label={`Download as ${format.label}`} onClick={() => { onDownload(format.id); menu.closeMenu(); }} />
      </>
    );
  }
  return (
    <>
      <IconButton title="Open file (.pptx)" onClick={onOpen}>
        <FolderOpen size={16} />
      </IconButton>
      <IconButton title={`Download as ${format.label}`} onClick={() => onDownload(format.id)}>
        <Download size={16} />
      </IconButton>
      <ToolbarDivider />
    </>
  );
}
