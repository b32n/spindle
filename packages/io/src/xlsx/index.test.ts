import * as xlsx from './index';
import * as io from '../index';

describe('@b32nio/spindle-io/xlsx', () => {
  it('re-exports the future function table helpers, the same ones the package root has', () => {
    expect(xlsx.fromFileFormula('_xlfn.IFS(A1,1)')).toBe('IFS(A1,1)');
    expect(xlsx.toFileFormula('IFS(A1,1)')).toBe('_xlfn.IFS(A1,1)');
    expect(xlsx.FUTURE_FUNCTIONS.get('IFS')).toBe('_xlfn.');
    expect(xlsx.exportXlsx).toBe(io.exportXlsx);
    expect(xlsx.importXlsx).toBe(io.importXlsx);
  });
});
