/**
 * visual_inspect `outputDir`: rendered images go to a folder the caller picks
 * (inside the allowed folders) instead of the server's temp artifacts folder,
 * for clients that cannot read that folder.
 */
import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DocxParser, hasVisualPdfBackend } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { handleVisualInspect, placeImages } from '../../packages/mcp-server/src/tools/visual';
import { withLoSlot } from '../helpers/lo-slot';

const RUN_LIVE = process.env.ASTRA_RUN_SLOW === '1' && hasVisualPdfBackend().available;

const tmpDirs: string[] = [];
function tmp(prefix: string): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), prefix));
  tmpDirs.push(dir);
  return dir;
}
afterEach(() => {
  for (const dir of tmpDirs.splice(0)) fs.rmSync(dir, { recursive: true, force: true });
});

async function openBlank(label: string) {
  const store = createMcpSessionStore();
  const handle = await store.openBuffer(DocxParser.createNewDocument('blank').rezip(), label);
  return { store, sessionId: handle.sessionId };
}

describe('visual_inspect outputDir', () => {
  it('needs withVisual', async () => {
    const { store, sessionId } = await openBlank('vo-1.docx');
    const res = await handleVisualInspect(store, { sessionId, outputDir: tmp('astra-vo-') });
    expect(res.ok).toBe(false);
    if (!res.ok) {
      expect(res.code).toBe('E_VALIDATION');
      expect(res.message).toMatch(/outputDir/);
    }
  });

  it.each([
    ['a folder outside the allowed roots', '/proc/astra-visual-out'],
    ['a path with ..', path.join(os.tmpdir(), 'x', '..', '..', 'etc')],
  ])('refuses %s', async (_name, dir) => {
    const { store, sessionId } = await openBlank('vo-2.docx');
    const res = await handleVisualInspect(store, { sessionId, withVisual: true, outputDir: dir });
    expect(res.ok).toBe(false);
    if (!res.ok) expect(res.code).toBe('E_SANDBOX');
  });

  it.runIf(RUN_LIVE)('writes the page images into the folder and reports those paths', async () => {
    const { store, sessionId } = await openBlank('vo-3.docx');
    const out = path.join(tmp('astra-vo-'), 'renders');
    const res = await withLoSlot(() => handleVisualInspect(store, { sessionId, withVisual: true, outputDir: out }));
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.outDir).toBe(fs.realpathSync(out));
    expect(res.images?.length).toBeGreaterThanOrEqual(1);
    for (const img of res.images ?? []) {
      expect(path.dirname(img.path)).toBe(fs.realpathSync(out));
      expect(path.basename(img.path)).toBe(`vo-3-page-${img.page}.png`);
      expect(fs.statSync(img.path).size).toBe(img.bytes);
    }
  }, 300_000);

  it.runIf(RUN_LIVE)('refuses to write through a symlink at the image name', async () => {
    const { store, sessionId } = await openBlank('vo-4.docx');
    const out = tmp('astra-vo-');
    const victim = path.join(tmp('astra-vo-victim-'), 'victim.txt');
    fs.writeFileSync(victim, 'keep');
    fs.symlinkSync(victim, path.join(out, 'vo-4-page-1.png'));
    const res = await withLoSlot(() => handleVisualInspect(store, { sessionId, withVisual: true, outputDir: out }));
    expect(res.ok).toBe(false);
    if (!res.ok) expect(res.code).toBe('E_SANDBOX');
    expect(fs.readFileSync(victim, 'utf8')).toBe('keep');
  }, 300_000);
});

describe('placing rendered images', () => {
  const rendered = (name: string, text: string) => {
    const file = path.join(tmp('astra-vo-src-'), name);
    fs.writeFileSync(file, text);
    return { page: 1, path: file, format: 'png' as const, bytes: text.length };
  };

  it('writes <document>-page-<n>.<ext> into the folder and removes the temp copy', () => {
    const out = fs.realpathSync(tmp('astra-vo-'));
    const img = rendered('page-1.png', 'first');
    const [placed] = placeImages([img], out, '/somewhere/My Report.docx');
    expect(placed.path).toBe(path.join(out, 'My-Report-page-1.png'));
    expect(fs.readFileSync(placed.path, 'utf8')).toBe('first');
    expect(fs.existsSync(img.path)).toBe(false);
  });

  it('replaces the previous render of the same page', () => {
    const out = fs.realpathSync(tmp('astra-vo-'));
    placeImages([rendered('a.png', 'old')], out, 'doc.docx');
    const [placed] = placeImages([rendered('b.png', 'new')], out, 'doc.docx');
    expect(fs.readFileSync(placed.path, 'utf8')).toBe('new');
  });

  it('refuses a symlink at the image name and leaves its target alone', () => {
    const out = fs.realpathSync(tmp('astra-vo-'));
    const victim = path.join(tmp('astra-vo-victim-'), 'victim.txt');
    fs.writeFileSync(victim, 'keep');
    fs.symlinkSync(victim, path.join(out, 'doc-page-1.png'));
    expect(() => placeImages([rendered('a.png', 'x')], out, 'doc.docx')).toThrow(/symlink/i);
    expect(fs.readFileSync(victim, 'utf8')).toBe('keep');
  });

  it('refuses a dangling symlink at the image name and creates nothing behind it', () => {
    const out = fs.realpathSync(tmp('astra-vo-'));
    const target = path.join(tmp('astra-vo-victim-'), 'not-yet.txt');
    fs.symlinkSync(target, path.join(out, 'doc-page-1.png'));
    expect(() => placeImages([rendered('a.png', 'x')], out, 'doc.docx')).toThrow(/symlink/i);
    expect(fs.existsSync(target)).toBe(false);
  });
});
