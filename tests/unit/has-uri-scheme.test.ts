/**
 * A Windows drive letter (C:\ or C:/) is a path. The checks that refuse remote references
 * (https:, data:, file:) used to match it as a one-letter scheme, so an absolute Windows path to an
 * image, an asset or a style guide was refused before the jail could look at it.
 */
import { describe, it, expect } from 'vitest';
import { hasUriScheme } from '../../packages/engine/src/sandbox/paths';
import { resolveStyleGuide } from '../../packages/engine/src/style-guide/resolve';
import { resolveImageFilePath } from '../../packages/engine/src/operations/image-file';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

describe('hasUriScheme', () => {
  it.each([
    'https://example.org/a.png',
    'http://x',
    'data:image/png;base64,AAAA',
    'file:///etc/passwd',
    'ftp:evil.docx',
    'mailto:a@b.c',
    'javascript:alert(1)',
    'C:a.png', // relative to the working folder of drive C: not a file inside a named folder
    'C:',
  ])('treats %s as a URI scheme', (ref) => {
    expect(hasUriScheme(ref)).toBe(true);
  });

  it.each([
    'C:\\Users\\Ana\\logo.png',
    'c:\\logo.png',
    'D:/images/logo.png',
    'C:\\',
    'C:/',
    'logo.png',
    'sub/logo.png',
    '/abs/logo.png',
    '.\\logo.png',
    '\\\\server\\share\\logo.png', // UNC
  ])('treats %s as a path', (ref) => {
    expect(hasUriScheme(ref)).toBe(false);
  });
});

describe('references that start with a drive letter reach the jail', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-scheme-'));

  it('an image path with a drive letter is no longer refused as a remote reference', () => {
    // On a POSIX host this is a plain file name that does not exist, so the failure must be the
    // missing file or the jail, never the scheme refusal that said "keep image files inside".
    let message = '';
    try {
      resolveImageFilePath(dir, 'C:\\pictures\\logo.png');
    } catch (err) {
      message = err instanceof Error ? err.message : String(err);
    }
    expect(message).not.toMatch(/outside the allowed directory/);
  });

  it('a remote image reference is still refused', () => {
    expect(() => resolveImageFilePath(dir, 'https://example.org/logo.png')).toThrow(/outside the allowed directory/);
    expect(() => resolveImageFilePath(dir, 'C:logo.png')).toThrow(/outside the allowed directory/);
  });

  it('a style guide path with a drive letter is looked up as a file, not refused as a URL', () => {
    let message = '';
    try {
      resolveStyleGuide('C:\\guides\\house.md', { allowedDir: dir });
    } catch (err) {
      message = err instanceof Error ? err.message : String(err);
    }
    expect(message).not.toMatch(/remote URL/);
    expect(() => resolveStyleGuide('https://example.org/house.md', { allowedDir: dir })).toThrow(/remote URL/);
  });
});
