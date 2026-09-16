import { useEffect, useMemo, useRef, useState } from 'react';
import { WorkbookProvider, WorkbookCanvas } from '@b32nio/spindle-sheets-react';
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';

// A real, on-brand workbook: a yarn inventory whose Total column and grand
// total are live formulas evaluated by the actual Spindle engine.
const yarnData = {
  id: 'spindle-hero',
  name: 'Yarn inventory',
  activeSheetId: 'sheet_1',
  defaultRowHeight: 22,
  defaultColWidth: 116,
  stylePool: {},
  formatPool: {},
  sheets: [
    {
      id: 'sheet_1',
      name: 'Inventory',
      cells: [
        { key: '0:0', cell: { value: 'Fiber' } },
        { key: '0:1', cell: { value: 'Skeins' } },
        { key: '0:2', cell: { value: '$/skein' } },
        { key: '0:3', cell: { value: 'Total' } },
        { key: '1:0', cell: { value: 'Indigo wool' } },
        { key: '1:1', cell: { value: 24 } },
        { key: '1:2', cell: { value: 8.5 } },
        { key: '1:3', cell: { value: 204, formula: '=B2*C2' } },
        { key: '2:0', cell: { value: 'Flax linen' } },
        { key: '2:1', cell: { value: 40 } },
        { key: '2:2', cell: { value: 6.25 } },
        { key: '2:3', cell: { value: 250, formula: '=B3*C3' } },
        { key: '3:0', cell: { value: 'Madder cotton' } },
        { key: '3:1', cell: { value: 18 } },
        { key: '3:2', cell: { value: 9 } },
        { key: '3:3', cell: { value: 162, formula: '=B4*C4' } },
        { key: '4:0', cell: { value: 'Brass silk' } },
        { key: '4:1', cell: { value: 12 } },
        { key: '4:2', cell: { value: 14 } },
        { key: '4:3', cell: { value: 168, formula: '=B5*C5' } },
        { key: '5:0', cell: { value: 'Total' } },
        { key: '5:3', cell: { value: 784, formula: '=SUM(D2:D5)' } },
      ],
      config: { defaultRowHeight: 22, defaultColWidth: 116, showGridLines: true },
      rowCount: 1000,
      colCount: 100,
    },
  ],
  selection: { ranges: [], activeCell: { row: 5, col: 3 } },
};

function makeWorkbook(): WorkbookImpl {
  const wb = new WorkbookImpl('spindle-hero', 'Yarn inventory');
  // setData accepts the serialized workbook shape; typed loosely here to keep
  // the marketing site decoupled from the core's internal data types.
  wb.setData(yarnData as unknown as Parameters<WorkbookImpl['setData']>[0]);
  return wb;
}

export default function LiveSheet() {
  const wb = useMemo(makeWorkbook, []);
  const bodyRef = useRef<HTMLDivElement>(null);
  const [size, setSize] = useState({ w: 0, h: 0 });

  useEffect(() => {
    const el = bodyRef.current;
    if (!el) return;
    const measure = () => setSize({ w: el.clientWidth, h: el.clientHeight });
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    measure();
    return () => ro.disconnect();
  }, []);

  return (
    <div className="livesheet">
      <div className="livesheet-bar">
        <span className="dot" /><span className="dot" /><span className="dot" />
        <span className="tt">yarn-inventory.sheet</span>
        <span className="live"><b />LIVE</span>
      </div>
      <div className="livesheet-body" ref={bodyRef}>
        {size.w > 0 && size.h > 0 && (
          <WorkbookProvider workbook={wb}>
            <WorkbookCanvas width={size.w} height={size.h} />
          </WorkbookProvider>
        )}
      </div>
    </div>
  );
}
