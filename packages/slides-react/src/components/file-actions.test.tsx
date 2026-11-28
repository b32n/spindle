import React from 'react';
import { SLIDES_EXPORT_FORMATS } from '../file-io';

let menu = { inMenu: false, closeMenu: jest.fn() };
jest.mock('@b32nio/spindle-shared/react', () => ({
  ...jest.requireActual('@b32nio/spindle-shared/react'),
  useToolbarMenu: () => menu,
}));

import { MenuRow } from '@b32nio/spindle-shared/react';
import { FileActions } from './FileActions';
import { ToolbarButton, ToolbarDivider } from './toolbarUI';

type P = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
const kids = (node: React.ReactElement) => React.Children.toArray((node.props as P).children) as React.ReactElement<P>[];

beforeEach(() => {
  menu = { inMenu: false, closeMenu: jest.fn() };
});

describe('FileActions', () => {
  it('inline: an Open and a Download button, then a divider', () => {
    const onOpen = jest.fn();
    const onDownload = jest.fn();
    const [open, download, divider] = kids(FileActions({ onOpen, onDownload }));
    expect(open.type).toBe(ToolbarButton);
    expect(open.props.title).toBe('Open file (.pptx)');
    expect(download.props.title).toBe('Download as Microsoft PowerPoint (.pptx)');
    expect(divider.type).toBe(ToolbarDivider);
    open.props.onClick();
    download.props.onClick();
    expect(onOpen).toHaveBeenCalledTimes(1);
    expect(onDownload).toHaveBeenCalledWith(SLIDES_EXPORT_FORMATS[0].id);
  });

  it('in the overflow menu: two rows that act and close the menu', () => {
    menu.inMenu = true;
    const onOpen = jest.fn();
    const onDownload = jest.fn();
    const [open, download] = kids(FileActions({ onOpen, onDownload }));
    expect(open.type).toBe(MenuRow);
    expect(open.props.label).toBe('Open file…');
    expect(download.props.label).toBe('Download as Microsoft PowerPoint (.pptx)');
    open.props.onClick();
    expect(onOpen).toHaveBeenCalledTimes(1);
    expect(menu.closeMenu).toHaveBeenCalledTimes(1);
    download.props.onClick();
    expect(onDownload).toHaveBeenCalledWith('pptx');
    expect(menu.closeMenu).toHaveBeenCalledTimes(2);
  });
});
