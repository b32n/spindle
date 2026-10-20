import { fileExtension, stripExtension, sanitizeFileName } from './download';

describe('download helpers', () => {
  it('reads the lower-cased extension', () => {
    expect(fileExtension('Report.XLSX')).toBe('xlsx');
    expect(fileExtension('a.b.csv')).toBe('csv');
    expect(fileExtension('noext')).toBe('');
  });

  it('strips only the final extension', () => {
    expect(stripExtension('a.b.csv')).toBe('a.b');
    expect(stripExtension('noext')).toBe('noext');
    expect(stripExtension('.hidden')).toBe('.hidden');
  });

  it('sanitizes file names', () => {
    expect(sanitizeFileName('a/b:c?.docx')).toBe('a_b_c_.docx');
    expect(sanitizeFileName('   ')).toBe('Untitled');
    expect(sanitizeFileName('', 'Sheet')).toBe('Sheet');
  });
});
