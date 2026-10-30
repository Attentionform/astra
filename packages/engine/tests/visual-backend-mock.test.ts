/**
 * Visual backend capability gate — mocked, no live binaries.
 *
 * Proves the headless gate fails closed without spawning processes:
 * bogus binary overrides and an empty PATH both yield E_DEPENDENCY,
 * and empty input yields E_VALIDATION before any backend probe.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  findOnPath,
  hasPdfBackend,
  assertPdfBackend,
  VISUAL_DEPENDENCY_REMEDY,
} from '../src/visual/backend';
import { renderDocxToImages, getVisualArtifactsDir } from '../src/visual/renderer';
import { DocxParser } from '../src/parser';

let savedPath: string | undefined;
const tmpDirs: string[] = [];
const artifactDirs: string[] = [];

function mkTmpDir(prefix: string): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), prefix));
  tmpDirs.push(dir);
  return dir;
}

beforeEach(() => {
  savedPath = process.env.PATH;
});

afterEach(() => {
  if (savedPath !== undefined) process.env.PATH = savedPath;
  else delete process.env.PATH;
  while (tmpDirs.length > 0) {
    const d = tmpDirs.pop() as string;
    try {
      fs.rmSync(d, { recursive: true, force: true });
    } catch {
      /* best-effort */
    }
  }
  while (artifactDirs.length > 0) {
    const d = artifactDirs.pop() as string;
    try {
      fs.rmSync(d, { recursive: true, force: true });
    } catch {
      /* best-effort */
    }
  }
});

function blankDocx(): Buffer {
  return DocxParser.createNewDocument('blank').rezip();
}

describe('visual backend gate (mocked)', () => {
  it('findOnPath returns null for bogus explicit paths', () => {
    expect(findOnPath('/nonexistent/soffice-astra-mock')).toBeNull();
    expect(findOnPath('/nonexistent/pdftoppm-astra-mock')).toBeNull();
  });

  it('findOnPath returns null when PATH is an empty dir', () => {
    const emptyDir = mkTmpDir('astra-visual-mock-path-');
    process.env.PATH = emptyDir;
    expect(findOnPath('soffice')).toBeNull();
    expect(findOnPath('pdftoppm')).toBeNull();
    expect(findOnPath('libreoffice')).toBeNull();
  });

  it('hasPdfBackend reports unavailable with bogus overrides (no spawn)', () => {
    const status = hasPdfBackend({
      sofficeBin: '/nonexistent/soffice-astra-mock',
      pdftoppmBin: '/nonexistent/pdftoppm-astra-mock',
    });
    expect(status.available).toBe(false);
    expect(status.backend).toBe('none');
    expect(status.missing).toContain('soffice');
    expect(status.missing).toContain('pdftoppm');
  });

  it('hasPdfBackend reports unavailable when PATH hides binaries', () => {
    const emptyDir = mkTmpDir('astra-visual-mock-empty-');
    process.env.PATH = emptyDir;
    const status = hasPdfBackend();
    expect(status.available).toBe(false);
    expect(status.backend).toBe('none');
    expect(status.missing.length).toBeGreaterThan(0);
  });

  it('assertPdfBackend throws E_DEPENDENCY with remedy (no spawn)', () => {
    try {
      assertPdfBackend({
        sofficeBin: '/nonexistent/soffice-astra-mock',
        pdftoppmBin: '/nonexistent/pdftoppm-astra-mock',
      });
      throw new Error('expected E_DEPENDENCY throw');
    } catch (err: unknown) {
      expect((err as { code?: unknown })?.code).toBe('E_DEPENDENCY');
      expect(String((err as { remedy?: unknown })?.remedy)).toMatch(/LibreOffice/i);
    }
    expect(VISUAL_DEPENDENCY_REMEDY).toMatch(/LibreOffice/i);
  });

  it('render rejects empty buffer with E_VALIDATION before backend probe', async () => {
    await expect(renderDocxToImages(Buffer.alloc(0), {})).rejects.toMatchObject({
      code: 'E_VALIDATION',
    });
  });

  it('render with bogus bins fails closed E_DEPENDENCY without spawning', async () => {
    await expect(
      renderDocxToImages(blankDocx(), {
        sofficeBin: '/nonexistent/soffice-astra-mock',
        pdftoppmBin: '/nonexistent/pdftoppm-astra-mock',
      })
    ).rejects.toMatchObject({ code: 'E_DEPENDENCY' });
  });

  it('artifacts dir helper stays inside base without binaries', () => {
    const dir = getVisualArtifactsDir('mock-check');
    artifactDirs.push(dir);
    expect(fs.existsSync(dir)).toBe(true);
    expect(path.resolve(dir).includes(path.join('ASTra', 'visual'))).toBe(true);
  });
});
