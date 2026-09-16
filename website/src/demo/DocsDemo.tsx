import { useMemo } from 'react';
import { DocumentProvider, DocumentEditor } from '@b32nio/spindle-docs-react';
import {
  DocumentImpl,
  type DocumentData,
  type HeaderFooterContent,
  type CommentAuthor,
} from '@b32nio/spindle-docs-core';
import DemoChrome from './DemoChrome';

const USERS: CommentAuthor[] = [
  { id: 'you', name: 'You' },
  { id: 'alice', name: 'Alice' },
  { id: 'bob', name: 'Bob' },
];

const PAGE = {
  size: { w: 816, h: 1056 },
  margins: { top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 },
  orientation: 'portrait' as const,
};

const header: HeaderFooterContent = {
  blocks: [
    {
      type: 'paragraph',
      alignment: 'left',
      content: [{ type: 'dynamicField', fieldType: 'title' }],
    },
  ],
  differentFirstPage: true,
  firstPageBlocks: [],
};

const footer: HeaderFooterContent = {
  blocks: [
    {
      type: 'paragraph',
      alignment: 'center',
      content: [
        { type: 'text', text: 'Page ' },
        { type: 'dynamicField', fieldType: 'pageNumber' },
        { type: 'text', text: ' of ' },
        { type: 'dynamicField', fieldType: 'totalPages' },
      ],
    },
  ],
};

const lorem =
  'Spindle paginates at the line level, so what you see on screen is what prints: real page breaks, real margins, real headers and footers. Keep typing and the pages reflow beneath your cursor. ';

const sampleDocument: DocumentData = {
  id: 'spindle-docs-demo',
  title: 'Spindle Docs: Live Demo',
  defaultPageConfig: PAGE,
  textStylePool: {
    bold: { bold: true },
    italic: { italic: true },
    code: { fontFamily: 'monospace', backgroundColor: '#f2efe8' },
    highlight: { backgroundColor: '#fbeecb' },
  },
  paragraphStylePool: {},
  sections: [
    {
      id: 'section_main',
      pageConfig: PAGE,
      header,
      footer,
      blocks: [
        {
          id: 'b1',
          type: 'heading',
          level: 1,
          alignment: 'center',
          content: [{ type: 'text', text: 'Woven, not bolted on' }],
        },
        {
          id: 'b2',
          type: 'paragraph',
          content: [
            { type: 'text', text: 'This is the real ' },
            { type: 'text', text: 'Spindle Docs', styleId: 'code' },
            { type: 'text', text: ' editor, a ' },
            { type: 'text', text: 'print-true', styleId: 'bold' },
            {
              type: 'text',
              text: ' document surface on a ProseMirror engine. Everything here is editable: type, format, add a comment, or drag the ruler margins.',
            },
          ],
        },
        {
          id: 'b3',
          type: 'heading',
          level: 2,
          content: [{ type: 'text', text: 'What you can try' }],
        },
        {
          id: 'b4',
          type: 'list-item',
          listType: 'bullet',
          level: 0,
          content: [
            { type: 'text', text: 'Rich text: ' },
            { type: 'text', text: 'bold', styleId: 'bold' },
            { type: 'text', text: ', ' },
            { type: 'text', text: 'italic', styleId: 'italic' },
            { type: 'text', text: ', and ' },
            { type: 'text', text: 'highlight', styleId: 'highlight' },
          ],
        },
        {
          id: 'b5',
          type: 'list-item',
          listType: 'bullet',
          level: 0,
          content: [{ type: 'text', text: 'Headings, bulleted and numbered lists' }],
        },
        {
          id: 'b6',
          type: 'list-item',
          listType: 'bullet',
          level: 0,
          content: [{ type: 'text', text: 'Tables, headers, footers, and live page numbers' }],
        },
        {
          id: 'b7',
          type: 'list-item',
          listType: 'bullet',
          level: 0,
          content: [{ type: 'text', text: 'Page setup (📄 in the toolbar) and margin dragging on the ruler' }],
        },
        {
          id: 'b8',
          type: 'heading',
          level: 2,
          content: [{ type: 'text', text: 'A sample table' }],
        },
        {
          id: 'b9',
          type: 'table',
          rows: [
            {
              id: 'r1',
              cells: [
                { id: 'c11', content: [{ type: 'text', text: 'Surface', styleId: 'bold' }] },
                { id: 'c12', content: [{ type: 'text', text: 'Package', styleId: 'bold' }] },
                { id: 'c13', content: [{ type: 'text', text: 'Status', styleId: 'bold' }] },
              ],
            },
            {
              id: 'r2',
              cells: [
                { id: 'c21', content: [{ type: 'text', text: 'Sheets' }] },
                { id: 'c22', content: [{ type: 'text', text: 'spindle-sheets-react', styleId: 'code' }] },
                { id: 'c23', content: [{ type: 'text', text: '✅ Shipping' }] },
              ],
            },
            {
              id: 'r3',
              cells: [
                { id: 'c31', content: [{ type: 'text', text: 'Docs' }] },
                { id: 'c32', content: [{ type: 'text', text: 'spindle-docs-react', styleId: 'code' }] },
                { id: 'c33', content: [{ type: 'text', text: '✅ Shipping' }] },
              ],
            },
            {
              id: 'r4',
              cells: [
                { id: 'c41', content: [{ type: 'text', text: 'Slides' }] },
                { id: 'c42', content: [{ type: 'text', text: 'spindle-slides-react', styleId: 'code' }] },
                { id: 'c43', content: [{ type: 'text', text: '✅ Shipping' }] },
              ],
            },
          ],
        },
        {
          id: 'b10',
          type: 'heading',
          level: 2,
          content: [{ type: 'text', text: 'True Layout pagination' }],
        },
        {
          id: 'b11',
          type: 'paragraph',
          content: [{ type: 'text', text: lorem.repeat(10) }],
        },
        {
          id: 'b12',
          type: 'paragraph',
          content: [{ type: 'text', text: lorem.repeat(12) }],
        },
      ],
    },
  ],
  createdAt: '2026-01-01T00:00:00.000Z',
  updatedAt: '2026-01-01T00:00:00.000Z',
};

function makeDocument(): DocumentImpl {
  const doc = new DocumentImpl();
  doc.setData(sampleDocument);
  return doc;
}

export default function DocsDemo() {
  const doc = useMemo(makeDocument, []);
  return (
    <DemoChrome active="docs" hint="Real engine · type, format, paginate">
      {({ width, height }) => (
        <DocumentProvider document={doc} currentUser={USERS[0]} mentionableUsers={USERS}>
          <DocumentEditor width={width} height={height} showToolbar showRuler />
        </DocumentProvider>
      )}
    </DemoChrome>
  );
}
