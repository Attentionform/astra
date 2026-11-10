/**
 * isPathInsideDir: a folder that ends in a separator, a drive or filesystem root, and Windows
 * spellings. The `platform` argument selects the path rules, so these run the same on every system.
 */
import { describe, it, expect } from 'vitest';
import { isPathInsideDir } from '../../packages/engine/src/sandbox/paths';

describe('isPathInsideDir: Windows rules (run on any host)', () => {
  const win = (c: string, d: string) => isPathInsideDir(c, d, 'win32');

  it('accepts a file under a folder given with or without a trailing separator', () => {
    expect(win('C:\\Users\\Ana\\doc.docx', 'C:\\Users\\Ana')).toBe(true);
    expect(win('C:\\Users\\Ana\\doc.docx', 'C:\\Users\\Ana\\')).toBe(true);
    expect(win('C:\\Users\\Ana\\doc.docx', 'C:/Users/Ana/')).toBe(true);
    expect(win('C:/Users/Ana/doc.docx', 'C:\\Users\\Ana\\')).toBe(true);
  });

  it('accepts a file under a drive root', () => {
    expect(win('C:\\doc.docx', 'C:\\')).toBe(true);
    expect(win('C:\\a\\b.docx', 'C:\\')).toBe(true);
    expect(win('C:/a/b.docx', 'C:/')).toBe(true);
  });

  it('is not case sensitive', () => {
    expect(win('c:\\users\\ana\\DOC.docx', 'C:\\Users\\ANA\\')).toBe(true);
  });

  it('refuses a sibling folder that shares the prefix, and other drives', () => {
    expect(win('C:\\Users\\Anabel\\doc.docx', 'C:\\Users\\Ana')).toBe(false);
    expect(win('C:\\Users\\Anabel\\doc.docx', 'C:\\Users\\Ana\\')).toBe(false);
    expect(win('D:\\doc.docx', 'C:\\')).toBe(false);
  });

  it('refuses a path that climbs out of the folder', () => {
    expect(win('C:\\Users\\Ana\\..\\Bob\\doc.docx', 'C:\\Users\\Ana\\')).toBe(false);
  });

  it('counts the folder itself as inside, with or without a trailing separator', () => {
    expect(win('C:\\Users\\Ana', 'C:\\Users\\Ana\\')).toBe(true);
    expect(win('C:\\Users\\Ana\\', 'C:\\Users\\Ana')).toBe(true);
  });
});

describe('isPathInsideDir: POSIX rules (run on any host)', () => {
  const posix = (c: string, d: string) => isPathInsideDir(c, d, 'linux');

  it('accepts a file under a folder given with or without a trailing slash', () => {
    expect(posix('/tmp/a/doc.docx', '/tmp/a')).toBe(true);
    expect(posix('/tmp/a/doc.docx', '/tmp/a/')).toBe(true);
    expect(posix('/tmp/a/doc.docx', '/tmp/a//')).toBe(true);
  });

  it('accepts a file under the filesystem root', () => {
    expect(posix('/doc.docx', '/')).toBe(true);
    expect(posix('/tmp/x/doc.docx', '/')).toBe(true);
  });

  it('is case sensitive and treats a backslash as a name character', () => {
    expect(posix('/tmp/A/doc.docx', '/tmp/a')).toBe(false);
    expect(posix('/tmp/a\\b/doc.docx', '/tmp/a')).toBe(false);
  });

  it('refuses a sibling folder that shares the prefix, and a climb out', () => {
    expect(posix('/tmp/ab/doc.docx', '/tmp/a')).toBe(false);
    expect(posix('/tmp/ab/doc.docx', '/tmp/a/')).toBe(false);
    expect(posix('/tmp/a/../b/doc.docx', '/tmp/a')).toBe(false);
  });
});
