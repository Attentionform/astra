/**
 * Visual backend resolver on each platform's rules — no live binaries.
 *
 * Fake executables are created in temp folders and the platform, PATHEXT and
 * environment are injected, so the Windows and macOS lookup rules run on any
 * host. The real render on Linux, Windows and macOS runners is covered by the
 * `visual-render` CI job.
 */
import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { findOnPath, installFolders, hasPdfBackend, VISUAL_DEPENDENCY_REMEDY } from '../src/visual/backend';

const tmpDirs: string[] = [];

function mkTmp(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-vis-plat-'));
  tmpDirs.push(dir);
  return dir;
}

function touchExe(dir: string, file: string): string {
  fs.mkdirSync(dir, { recursive: true });
  const p = path.join(dir, file);
  fs.writeFileSync(p, '#!/bin/sh\nexit 0\n');
  fs.chmodSync(p, 0o755);
  return p;
}

afterEach(() => {
  while (tmpDirs.length > 0) {
    fs.rmSync(tmpDirs.pop() as string, { recursive: true, force: true });
  }
});

describe('findOnPath on Windows rules', () => {
  it('finds a bare name through .exe', () => {
    const bin = mkTmp();
    const exe = touchExe(bin, 'soffice.exe');
    expect(findOnPath('soffice', bin, { platform: 'win32', pathExt: '.COM;.EXE;.BAT', fallbackDirs: false })).toBe(exe);
  });

  it('finds .com when no .exe exists', () => {
    const bin = mkTmp();
    const com = touchExe(bin, 'pdftoppm.com');
    expect(findOnPath('pdftoppm', bin, { platform: 'win32', fallbackDirs: false })).toBe(com);
  });

  it('prefers .exe over .com, whatever the PATHEXT order', () => {
    const bin = mkTmp();
    touchExe(bin, 'soffice.com');
    const exe = touchExe(bin, 'soffice.exe');
    expect(findOnPath('soffice', bin, { platform: 'win32', pathExt: '.COM;.EXE', fallbackDirs: false })).toBe(exe);
  });

  it('ignores .cmd and .bat, which execFile cannot spawn', () => {
    const bin = mkTmp();
    touchExe(bin, 'soffice.cmd');
    touchExe(bin, 'soffice.bat');
    expect(findOnPath('soffice', bin, { platform: 'win32', pathExt: '.CMD;.BAT', fallbackDirs: false })).toBeNull();
  });

  it('accepts a name that already has .exe', () => {
    const bin = mkTmp();
    const exe = touchExe(bin, 'soffice.exe');
    expect(findOnPath('soffice.exe', bin, { platform: 'win32', fallbackDirs: false })).toBe(exe);
  });

  it('does not add extensions on other platforms', () => {
    const bin = mkTmp();
    touchExe(bin, 'soffice.exe');
    expect(findOnPath('soffice', bin, { platform: 'linux', fallbackDirs: false })).toBeNull();
  });

  it('strips quotes around a PATH entry', () => {
    const bin = mkTmp();
    const exe = touchExe(bin, 'soffice.exe');
    expect(findOnPath('soffice', `"${bin}"`, { platform: 'win32', fallbackDirs: false })).toBe(exe);
  });

  it('falls back to Program Files when PATH has no LibreOffice', () => {
    const programFiles = mkTmp();
    const exe = touchExe(path.join(programFiles, 'LibreOffice', 'program'), 'soffice.exe');
    const env = { ProgramFiles: programFiles, PATH: '' } as NodeJS.ProcessEnv;
    // The folder list is computed from the injected environment, then narrowed to the fake install.
    const dirs = installFolders('soffice', 'win32', env).filter((d) => d.startsWith(programFiles));
    expect(dirs).toEqual([path.join(programFiles, 'LibreOffice', 'program')]);
    expect(findOnPath('soffice', undefined, { platform: 'win32', env, installDirs: dirs })).toBe(exe);
  });
});

describe('findOnPath on macOS rules', () => {
  it('falls back to the LibreOffice application bundle in the home folder', () => {
    const home = mkTmp();
    const soffice = touchExe(path.join(home, 'Applications', 'LibreOffice.app', 'Contents', 'MacOS'), 'soffice');
    const env = { HOME: home, PATH: '' } as NodeJS.ProcessEnv;
    // Only the home-folder entry is searched, so a real /Applications/LibreOffice.app on the host cannot win.
    const dirs = installFolders('soffice', 'darwin', env).filter((d) => d.startsWith(home));
    expect(dirs).toEqual([path.join(home, 'Applications', 'LibreOffice.app', 'Contents', 'MacOS')]);
    expect(findOnPath('soffice', undefined, { platform: 'darwin', env, installDirs: dirs })).toBe(soffice);
  });

  it('searches the folders in order: system bundle before the home bundle', () => {
    const dirs = installFolders('soffice', 'darwin', { HOME: '/Users/x' } as NodeJS.ProcessEnv);
    expect(dirs.indexOf('/Applications/LibreOffice.app/Contents/MacOS')).toBeLessThan(
      dirs.indexOf(path.join('/Users/x', 'Applications', 'LibreOffice.app', 'Contents', 'MacOS'))
    );
  });

  it('lists the system bundle, Homebrew and MacPorts folders', () => {
    const dirs = installFolders('soffice', 'darwin', { HOME: '/Users/x' } as NodeJS.ProcessEnv);
    expect(dirs).toContain('/Applications/LibreOffice.app/Contents/MacOS');
    expect(dirs).toContain('/opt/homebrew/bin');
    expect(dirs).toContain('/usr/local/bin');
    expect(dirs).toContain('/opt/local/bin');
    const poppler = installFolders('pdftoppm', 'darwin', { HOME: '/Users/x' } as NodeJS.ProcessEnv);
    expect(poppler).not.toContain('/Applications/LibreOffice.app/Contents/MacOS');
    expect(poppler).toContain('/opt/homebrew/bin');
  });
});

describe('relative PATH entries', () => {
  it('returns an absolute path', () => {
    const bin = mkTmp();
    touchExe(bin, 'pdftoppm');
    const found = findOnPath('pdftoppm', path.relative(process.cwd(), bin), { platform: 'linux', fallbackDirs: false });
    expect(found).not.toBeNull();
    expect(path.isAbsolute(found as string)).toBe(true);
  });
});

describe('fallback switches', () => {
  it('an explicit PATH string turns the fallback off', () => {
    const home = mkTmp();
    touchExe(path.join(home, 'Applications', 'LibreOffice.app', 'Contents', 'MacOS'), 'soffice');
    const env = { HOME: home } as NodeJS.ProcessEnv;
    expect(findOnPath('soffice', '', { platform: 'darwin', env })).toBeNull();
  });

  it('ASTRA_VISUAL_NO_FALLBACK_DIRS=1 turns the fallback off', () => {
    const home = mkTmp();
    touchExe(path.join(home, 'Applications', 'LibreOffice.app', 'Contents', 'MacOS'), 'soffice');
    const env = { HOME: home, PATH: '', ASTRA_VISUAL_NO_FALLBACK_DIRS: '1' } as NodeJS.ProcessEnv;
    expect(findOnPath('soffice', undefined, { platform: 'darwin', env })).toBeNull();
  });

  it('hasPdfBackend with an empty explicit PATH reports both binaries missing', () => {
    const status = hasPdfBackend({ pathEnv: mkTmp() });
    expect(status.available).toBe(false);
    expect(status.missing).toEqual(['soffice', 'pdftoppm']);
  });
});

describe('remedy text', () => {
  it('names an install route for each platform', () => {
    expect(VISUAL_DEPENDENCY_REMEDY).toContain('apt install');
    expect(VISUAL_DEPENDENCY_REMEDY).toContain('brew install');
    expect(VISUAL_DEPENDENCY_REMEDY).toContain('winget install');
    expect(VISUAL_DEPENDENCY_REMEDY).toContain('pdftoppm.exe');
  });
});
