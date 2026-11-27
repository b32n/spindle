import { ImportError, toImportError } from './import-error';

describe('ImportError', () => {
  it('carries a code and a name, and is an Error', () => {
    const e = new ImportError('size-limit', 'too big');
    expect(e).toBeInstanceOf(Error);
    expect(e.name).toBe('ImportError');
    expect(e.code).toBe('size-limit');
    expect(e.message).toBe('too big');
  });
});

describe('toImportError', () => {
  it('passes an ImportError through unchanged', () => {
    const e = new ImportError('encrypted', 'locked');
    expect(toImportError(e)).toBe(e);
  });

  it.each([
    [new Error('boom'), 'The file could not be read (boom).'],
    [new Error(''), 'The file could not be read (Error).'],
    ['plain string', 'The file could not be read (plain string).'],
    [undefined, 'The file could not be read (undefined).'],
  ])('wraps %p as a corrupt-file error', (thrown, message) => {
    const e = toImportError(thrown);
    expect(e).toBeInstanceOf(ImportError);
    expect(e.code).toBe('corrupt');
    expect(e.message).toBe(message);
  });
});
