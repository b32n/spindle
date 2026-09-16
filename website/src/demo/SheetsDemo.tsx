import { useMemo } from 'react';
import { WorkbookProvider, WorkbookCanvas } from '@b32nio/spindle-sheets-react';
import { WorkbookImpl } from '@b32nio/spindle-sheets-core';
import DemoChrome from './DemoChrome';

// A realistic workbook: a studio's Q3 plan. Every total is a live formula
// evaluated by the real Spindle engine; edit a monthly figure and the row
// totals, column totals, net, and average all recalculate.
const data = {
  id: 'spindle-sheets-demo',
  name: 'Q3 plan',
  activeSheetId: 'sheet_1',
  defaultRowHeight: 24,
  defaultColWidth: 128,
  stylePool: {},
  formatPool: {},
  sheets: [
    {
      id: 'sheet_1',
      name: 'Q3 plan',
      cells: [
        { key: '0:0', cell: { value: 'Line item' } },
        { key: '0:1', cell: { value: 'Jul' } },
        { key: '0:2', cell: { value: 'Aug' } },
        { key: '0:3', cell: { value: 'Sep' } },
        { key: '0:4', cell: { value: 'Q3 total' } },

        { key: '1:0', cell: { value: 'Subscriptions' } },
        { key: '1:1', cell: { value: 42000 } },
        { key: '1:2', cell: { value: 45800 } },
        { key: '1:3', cell: { value: 51200 } },
        { key: '1:4', cell: { value: 139000, formula: '=SUM(B2:D2)' } },

        { key: '2:0', cell: { value: 'Services' } },
        { key: '2:1', cell: { value: 12000 } },
        { key: '2:2', cell: { value: 9500 } },
        { key: '2:3', cell: { value: 14300 } },
        { key: '2:4', cell: { value: 35800, formula: '=SUM(B3:D3)' } },

        { key: '3:0', cell: { value: 'Marketplace' } },
        { key: '3:1', cell: { value: 6800 } },
        { key: '3:2', cell: { value: 7200 } },
        { key: '3:3', cell: { value: 8100 } },
        { key: '3:4', cell: { value: 22100, formula: '=SUM(B4:D4)' } },

        { key: '4:0', cell: { value: 'Total revenue' } },
        { key: '4:1', cell: { value: 60800, formula: '=SUM(B2:B4)' } },
        { key: '4:2', cell: { value: 62500, formula: '=SUM(C2:C4)' } },
        { key: '4:3', cell: { value: 73600, formula: '=SUM(D2:D4)' } },
        { key: '4:4', cell: { value: 196900, formula: '=SUM(E2:E4)' } },

        { key: '6:0', cell: { value: 'Infrastructure' } },
        { key: '6:1', cell: { value: 8200 } },
        { key: '6:2', cell: { value: 8600 } },
        { key: '6:3', cell: { value: 9100 } },
        { key: '6:4', cell: { value: 25900, formula: '=SUM(B7:D7)' } },

        { key: '7:0', cell: { value: 'Salaries' } },
        { key: '7:1', cell: { value: 38000 } },
        { key: '7:2', cell: { value: 38000 } },
        { key: '7:3', cell: { value: 41000 } },
        { key: '7:4', cell: { value: 117000, formula: '=SUM(B8:D8)' } },

        { key: '8:0', cell: { value: 'Marketing' } },
        { key: '8:1', cell: { value: 5400 } },
        { key: '8:2', cell: { value: 7800 } },
        { key: '8:3', cell: { value: 6200 } },
        { key: '8:4', cell: { value: 19400, formula: '=SUM(B9:D9)' } },

        { key: '9:0', cell: { value: 'Total costs' } },
        { key: '9:1', cell: { value: 51600, formula: '=SUM(B7:B9)' } },
        { key: '9:2', cell: { value: 54400, formula: '=SUM(C7:C9)' } },
        { key: '9:3', cell: { value: 56300, formula: '=SUM(D7:D9)' } },
        { key: '9:4', cell: { value: 162300, formula: '=SUM(E7:E9)' } },

        { key: '11:0', cell: { value: 'Net (rev − costs)' } },
        { key: '11:1', cell: { value: 9200, formula: '=B5-B10' } },
        { key: '11:2', cell: { value: 8100, formula: '=C5-C10' } },
        { key: '11:3', cell: { value: 17300, formula: '=D5-D10' } },
        { key: '11:4', cell: { value: 34600, formula: '=E5-E10' } },

        { key: '13:0', cell: { value: 'Avg monthly revenue' } },
        { key: '13:1', cell: { value: 65633.33, formula: '=AVERAGE(B5:D5)' } },
      ],
      config: { defaultRowHeight: 24, defaultColWidth: 128, showGridLines: true },
      rowCount: 1000,
      colCount: 100,
    },
  ],
  selection: { ranges: [], activeCell: { row: 11, col: 4 } },
};

function makeWorkbook(): WorkbookImpl {
  const wb = new WorkbookImpl('spindle-sheets-demo', 'Q3 plan');
  wb.setData(data as unknown as Parameters<WorkbookImpl['setData']>[0]);
  return wb;
}

export default function SheetsDemo() {
  const wb = useMemo(makeWorkbook, []);
  return (
    <DemoChrome active="sheets" hint="Real engine · edit a cell, watch it recalc">
      {({ width, height }) => (
        <WorkbookProvider workbook={wb}>
          <WorkbookCanvas width={width} height={height} />
        </WorkbookProvider>
      )}
    </DemoChrome>
  );
}
