import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import path from 'path';
import fs from 'fs';
import fsp from 'fs/promises';
import os from 'os';
import {
  ensureAstraGitRepo,
  getRepoDir,
  initAstraRepo,
  commitChange,
  getHistory,
  getSidecarHistoryWarning,
  clearSidecarHistoryWarning,
  SIDECAR_TOMBSTONE_FILENAME,
  DocxParser,
  parseFontSize,
  FONT_SIZE_PT_REGEX,
  replaceOpSchema,
  styleOpSchema,
  OperationSchema,
  validateOperationStrict,
  executeStyle,
  validateDocxBuffer,
} from '../../packages/engine/src';
import { readDocxFileBuffer } from '../../packages/engine/src/parser';
import { ZIP_BOMB_MAX_UNCOMPRESSED_BYTES } from '../../packages/engine/src/zip-hardening';
import {
  stripAnsi,
  shouldStripColor,
  formatCliOutput,
  ANSI_REGEX,
  stringifyJsonPayload,
} from '../../packages/cli/src/cli';

function blankDocxBuf(): Buffer {
  return DocxParser.createNewDocument('blank').rezip();
}

const POSIX_ONLY = process.platform === 'win32' ? describe.skip : describe;

describe('Hostile-review sweep: perms, pt, parser path, NO_COLOR', () => {
  // =========================================================================
  // 1. Perms: pre-existing 0755 migration + writeAndAdd subdir modes (POSIX)
  // =========================================================================
  POSIX_ONLY('perms: 0700 migration and subdir modes', () => {
    let tmpDir: string;
    let docPath: string;

    beforeEach(async () => {
      tmpDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-lows-perms-'));
      docPath = path.join(tmpDir, 'document.docx');
    });

    afterEach(async () => {
      await fsp.rm(tmpDir, { recursive: true, force: true });
    });

    it('ensureAstraGitRepo migrates a pre-existing 0755 tree to 0700', async () => {
      const repoDir = getRepoDir(docPath);
      const nested = path.join(repoDir, 'word', 'media');
      await fsp.mkdir(nested, { recursive: true });
      await fsp.chmod(repoDir, 0o755);
      await fsp.chmod(path.dirname(repoDir), 0o755);
      await fsp.chmod(path.join(repoDir, 'word'), 0o755);
      await fsp.chmod(nested, 0o755);

      const returned = await ensureAstraGitRepo(docPath);
      expect(returned).toBe(repoDir);
      expect((await fsp.stat(repoDir)).mode & 0o777).toBe(0o700);
      expect((await fsp.stat(path.dirname(repoDir))).mode & 0o777).toBe(0o700);
      expect((await fsp.stat(nested)).mode & 0o777).toBe(0o700);
    });

    it('initAstraRepo/commitChange writeAndAdd subdirs land on 0700', async () => {
      const xml = '<w:document><w:body><w:p>Hi</w:p></w:body></w:document>';
      await initAstraRepo(docPath, xml, { 'word/media/pic.png': Buffer.from([1, 2, 3]) });
      const repoDir = getRepoDir(docPath);
      expect((await fsp.stat(path.join(repoDir, 'word', 'media'))).mode & 0o777).toBe(0o700);

      await commitChange(docPath, xml, 'second', { 'word/charts/c1.xml': '<c/>' });
      expect((await fsp.stat(path.join(repoDir, 'word', 'charts'))).mode & 0o777).toBe(0o700);
      expect((await fsp.stat(repoDir)).mode & 0o777).toBe(0o700);
    });

    it('migration skips files and symlinked dirs without following (no escape)', async () => {
      const repoDir = getRepoDir(docPath);
      const realSub = path.join(repoDir, 'word', 'media');
      await fsp.mkdir(realSub, { recursive: true });
      await fsp.writeFile(path.join(realSub, 'a.bin'), 'hi');
      const outside = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-lows-outside-'));
      try {
        await fsp.chmod(outside, 0o755);
        await fsp.symlink(outside, path.join(repoDir, 'word', 'linkdir'));
        await fsp.symlink(
          path.join(realSub, 'a.bin'),
          path.join(repoDir, 'word', 'linkfile')
        );
        await fsp.chmod(repoDir, 0o755);
        await fsp.chmod(path.join(repoDir, 'word'), 0o755);
        await fsp.chmod(realSub, 0o755);
        await ensureAstraGitRepo(docPath);
        expect((await fsp.stat(repoDir)).mode & 0o777).toBe(0o700);
        expect((await fsp.stat(path.join(repoDir, 'word'))).mode & 0o777).toBe(0o700);
        expect((await fsp.stat(realSub)).mode & 0o777).toBe(0o700);
        expect((await fsp.lstat(path.join(repoDir, 'word', 'linkdir'))).isSymbolicLink()).toBe(true);
        expect((await fsp.stat(outside)).mode & 0o777).toBe(0o755);
      } finally {
        await fsp.rm(outside, { recursive: true, force: true });
      }
    });
  });

  // =========================================================================
  // 2. pt: strict regex with real capture group, garbage rejection, semantics
  // =========================================================================
  describe('pt: strict regex + documented half-point semantics', () => {
    it('regex exposes a capture group holding the points value', () => {
      expect(FONT_SIZE_PT_REGEX.source).toContain('(');
      const m = '10.5pt'.match(FONT_SIZE_PT_REGEX);
      expect(m?.[1]).toBe('10.5');
      expect('12PT'.match(FONT_SIZE_PT_REGEX)?.[1]).toBe('12');
    });

    it('converts pt strings to half-points, passes numbers through', () => {
      expect(parseFontSize('12pt')).toBe(24);
      expect(parseFontSize('10.5 pt')).toBe(21);
      expect(parseFontSize('0pt')).toBe(0);
      expect(parseFontSize('0.5pt')).toBe(1);
      // Documented semantics: bare numbers are half-points, never points.
      expect(parseFontSize(12)).toBe(12);
      expect(parseFontSize(24)).toBe(24);
      expect(parseFontSize(0)).toBe(0);
    });

    it('rounds bare float numbers to integer half-points, integers pass through', () => {
      expect(parseFontSize(10.5)).toBe(11);
      expect(parseFontSize(10.4)).toBe(10);
      expect(parseFontSize(10.6)).toBe(11);
      expect(parseFontSize(0.5)).toBe(1);
      expect(parseFontSize(12)).toBe(12);
      expect(parseFontSize('12pt')).toBe(24);
    });

    it('negative numbers fail closed with a clean validation error', () => {
      expect(parseFontSize(-3)).toBeUndefined();
      expect(styleOpSchema.safeParse({ action: 'style', target: 'p_1', fontSize: -3 }).success).toBe(false);
      expect(() =>
        validateOperationStrict({ action: 'style', target: 'p_1', fontSize: -3 })
      ).toThrow(/E_VALIDATION/);
    });

    it('bare float font size lands as an integer w:sz and the document validates', async () => {
      const parser = DocxParser.createNewDocument('blank');
      const ast = await parser.getAST();
      const target = ast.sections[0].nodes[0].id;
      executeStyle({ action: 'style', target, fontSize: 10.5 } as never, ast);
      const xml = parser.readDocumentXml();
      expect(xml).toContain('<w:sz w:val="11"/>');
      expect(xml).not.toContain('w:val="10.5"');
      const res = await validateDocxBuffer(parser.rezip());
      expect(res.ok).toBe(true);
      expect(res.errors).toEqual([]);
    });

    it('rejects bare numeric strings to match the schema gate', () => {
      for (const bare of ['+5', '1e3', '.5', '24', '-5', '12', '0']) {
        expect(parseFontSize(bare), bare).toBeUndefined();
      }
      expect(parseFontSize('+5pt')).toBeUndefined();
    });

    it('rejects negatives, floats-garbage and wrong units', () => {
      for (const bad of [
        '-12pt', '+5pt', '.5pt', '12.pt', '12.5.3pt', '12px',
        'pt', 'pt12', '1e3pt', 'NaNpt', 'Infinitypt', '', '   ',
      ]) {
        expect(parseFontSize(bad), bad).toBeUndefined();
      }
      expect(parseFontSize(-3)).toBeUndefined();
      expect(parseFontSize(Number.NaN)).toBeUndefined();
      expect(parseFontSize(Number.POSITIVE_INFINITY)).toBeUndefined();
      expect(parseFontSize('-5')).toBeUndefined();
    });

    it('watermark accepts 0 and "0pt" like body text (nonnegative uniformity)', () => {
      expect(OperationSchema.safeParse({ action: 'watermark', fontSize: 0 }).success).toBe(true);
      expect(OperationSchema.safeParse({ action: 'watermark', fontSize: '0pt' }).success).toBe(true);
      expect(OperationSchema.safeParse({ action: 'watermark', sz: 0 }).success).toBe(true);
      expect(OperationSchema.safeParse({ action: 'watermark', sz: '0pt' }).success).toBe(true);
      expect(parseFontSize('0pt')).toBe(0);
      expect(parseFontSize(0)).toBe(0);
    });

    it('schema gate rejects negatives and pt-garbage, accepts canonical shapes', () => {
      expect(styleOpSchema.safeParse({ action: 'style', target: 'p_1', sz: '12pt' }).success).toBe(true);
      expect(styleOpSchema.safeParse({ action: 'style', target: 'p_1', fontSize: 24 }).success).toBe(true);
      expect(styleOpSchema.safeParse({ action: 'style', target: 'p_1', sz: '-5pt' }).success).toBe(false);
      expect(styleOpSchema.safeParse({ action: 'style', target: 'p_1', sz: '12.5.3pt' }).success).toBe(false);
      expect(styleOpSchema.safeParse({ action: 'style', target: 'p_1', sz: -3 }).success).toBe(false);
      expect(styleOpSchema.safeParse({ action: 'style', target: 'p_1', sz: '24' }).success).toBe(false);
      expect(styleOpSchema.safeParse({ action: 'style', target: 'p_1', fontSize: '+5' }).success).toBe(false);
      expect(replaceOpSchema.safeParse({ action: 'replace', target: 'p_1', text: 'x', fontSize: '-5pt' }).success).toBe(false);
      expect(replaceOpSchema.safeParse({ action: 'replace', target: 'p_1', text: 'x', fontSize: '12pt' }).success).toBe(true);
      expect(replaceOpSchema.safeParse({ action: 'replace', target: 'p_1', text: 'x', fontSize: '24' }).success).toBe(false);
    });
  });

  // =========================================================================
  // 3. Parser string path: sandbox gates + stat/isFile + size + EISDIR
  // =========================================================================
  describe('parser: string-path sandbox gates', () => {
    let tmpDir: string;

    beforeEach(async () => {
      tmpDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-lows-parser-'));
    });

    afterEach(async () => {
      await fsp.rm(tmpDir, { recursive: true, force: true });
    });

    it('parses a valid file path', async () => {
      const p = path.join(tmpDir, 'doc.docx');
      await fsp.writeFile(p, blankDocxBuf());
      const parser = new DocxParser(p);
      const ast = await parser.getAST();
      expect(ast.sections.length).toBeGreaterThan(0);
    });

    it('missing path keeps the legacy "File not found: <path>" message', () => {
      const fake = path.join(tmpDir, 'nope.docx');
      expect(() => new DocxParser(fake)).toThrow(`File not found: ${fake}`);
      expect(() => readDocxFileBuffer(fake)).toThrow(`File not found: ${fake}`);
    });

    it('directory path is an E_VALIDATION error, never raw EISDIR', () => {
      expect(() => new DocxParser(tmpDir)).toThrow(/E_VALIDATION/);
      expect(() => readDocxFileBuffer(tmpDir)).toThrow(/E_VALIDATION/);
      try {
        readDocxFileBuffer(tmpDir);
        expect.unreachable();
      } catch (err: unknown) {
        expect(String((err as Error).message)).not.toMatch(/EISDIR/);
      }
    });

    it('null-byte path is E_VALIDATION', () => {
      expect(() => readDocxFileBuffer('a\0b.docx')).toThrow(/E_VALIDATION/);
    });

    it('oversize file fails closed with E_LIMIT before reading', async () => {
      const p = path.join(tmpDir, 'huge.docx');
      await fsp.writeFile(p, 'x');
      await fsp.truncate(p, ZIP_BOMB_MAX_UNCOMPRESSED_BYTES + 1);
      expect(() => readDocxFileBuffer(p)).toThrow(/E_LIMIT/);
    });

    it('Buffer input still works', async () => {
      const parser = new DocxParser(blankDocxBuf());
      expect((await parser.getAST()).sections.length).toBeGreaterThan(0);
    });

    it('non-regular file is E_VALIDATION (POSIX)', () => {
      if (process.platform === 'win32') return;
      if (!fs.existsSync('/dev/null')) return;
      expect(() => readDocxFileBuffer('/dev/null')).toThrow(/E_VALIDATION/);
    });
  });

  // =========================================================================
  // 4. NO_COLOR: spec non-empty check + OSC/charset/colon + JSON stripping
  // =========================================================================
  describe('cli: spec-compliant NO_COLOR and ANSI coverage', () => {
    const savedNoColor = process.env.NO_COLOR;

    afterEach(() => {
      if (savedNoColor !== undefined) process.env.NO_COLOR = savedNoColor;
      else delete process.env.NO_COLOR;
    });

    it('empty NO_COLOR does NOT strip on a TTY (spec non-empty check)', () => {
      process.env.NO_COLOR = '';
      expect(shouldStripColor({ isTTY: true })).toBe(false);
      const colored = '\u001b[34mInfo message\u001b[0m';
      expect(formatCliOutput(colored, { isTTY: true })).toBe(colored);
    });

    it('non-empty NO_COLOR strips even on a TTY', () => {
      process.env.NO_COLOR = '1';
      expect(shouldStripColor({ isTTY: true })).toBe(true);
      expect(shouldStripColor({ isTTY: false })).toBe(true);
      expect(formatCliOutput('\u001b[34mInfo\u001b[0m', { isTTY: true })).toBe('Info');
    });

    it('TTY without NO_COLOR keeps color; non-TTY strips', () => {
      delete process.env.NO_COLOR;
      const colored = '\u001b[34mInfo\u001b[0m';
      expect(shouldStripColor({ isTTY: true })).toBe(false);
      expect(shouldStripColor({ isTTY: false })).toBe(true);
      expect(formatCliOutput(colored, { isTTY: false })).toBe('Info');
    });

    it('strips OSC hyperlinks, charset designators and colon truecolor', () => {
      expect(stripAnsi('\u001b]8;;https://example.com\u0007link\u001b]8;;\u0007')).toBe('link');
      expect(stripAnsi('\u001b(Bplain')).toBe('plain');
      expect(stripAnsi('\u001b[38:2::255:0:0mred\u001b[0m')).toBe('red');
      expect(stripAnsi('\u001b[4:3munder\u001b[0m')).toBe('under');
      expect(stripAnsi('\u001b[1;31mbold red\u001b[0m')).toBe('bold red');
    });

    it('ANSI_REGEX matches the extended sequences', () => {
      const fresh = new RegExp(ANSI_REGEX.source);
      for (const seq of ['\u001b[32m', '\u001b]8;;https://x.com\u0007', '\u001b(B', '\u001b[38:2::1:2:3m', '\u001b[4:3m']) {
        fresh.lastIndex = 0;
        expect(fresh.test(seq), JSON.stringify(seq)).toBe(true);
      }
    });

    it('stringifyJsonPayload strips ANSI when color is disabled', () => {
      process.env.NO_COLOR = '1';
      const payload = { msg: '\u001b[31mhi\u001b[0m' };
      const out = stringifyJsonPayload(payload);
      expect(out).toBe(JSON.stringify({ msg: 'hi' }));
      expect(JSON.parse(out)).toEqual({ msg: 'hi' });
    });
  });
});

describe('R-6: sidecar-removal tombstone warning', () => {
  const R6_XML = '<w:document><w:body><w:p>Hi</w:p></w:body></w:document>';
  type HistoryWithWarning = Array<{
    oid: string;
    message: string;
    author: string;
    timestamp: Date;
  }> & { warning: string | null };
  function readWarning(history: unknown): string | null {
    return (history as HistoryWithWarning).warning ?? null;
  }

  let tmpDir: string;
  let docPath: string;

  beforeEach(async () => {
    tmpDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-r6-'));
    docPath = path.join(tmpDir, 'memo.docx');
  });

  afterEach(async () => {
    await fsp.rm(tmpDir, { recursive: true, force: true });
  });

  it('fresh open records history with no warning', async () => {
    await initAstraRepo(docPath, R6_XML);
    expect(getSidecarHistoryWarning(docPath)).toBeNull();
    const history = await getHistory(docPath);
    expect(history).toHaveLength(1);
    expect(readWarning(history)).toBeNull();
    // Registry lives outside `.astra/` so whole-dir removal cannot take it.
    expect(fs.existsSync(path.join(tmpDir, SIDECAR_TOMBSTONE_FILENAME))).toBe(true);
  });

  it('per-document sidecar deletion + reopen warns in response fields, never blocks', async () => {
    await initAstraRepo(docPath, R6_XML);
    await commitChange(docPath, R6_XML, 'second');
    await fsp.rm(getRepoDir(docPath), { recursive: true, force: true });

    await initAstraRepo(docPath, R6_XML); // reopen must not throw

    expect(getSidecarHistoryWarning(docPath)).toMatch(/possible history removal/);
    const history = await getHistory(docPath);
    expect(history).toHaveLength(1);
    expect(readWarning(history)).toMatch(/possible history removal/);
  });

  it('whole `.astra/` directory deletion + reopen warns', async () => {
    await initAstraRepo(docPath, R6_XML);
    await fsp.rm(path.join(tmpDir, '.astra'), { recursive: true, force: true });
    expect(fs.existsSync(path.join(tmpDir, SIDECAR_TOMBSTONE_FILENAME))).toBe(true);

    await initAstraRepo(docPath, R6_XML);
    expect(getSidecarHistoryWarning(docPath)).toMatch(/possible history removal/);
  });

  it('normal second open on an intact chain stays silent', async () => {
    await initAstraRepo(docPath, R6_XML);
    await commitChange(docPath, R6_XML, 'second');
    await ensureAstraGitRepo(docPath);
    expect(getSidecarHistoryWarning(docPath)).toBeNull();
    const history = await getHistory(docPath);
    expect(history).toHaveLength(2);
    expect(readWarning(history)).toBeNull();
  });

  it('repoDir-form lookup agrees and the clear hook resets the field', async () => {
    await initAstraRepo(docPath, R6_XML);
    await fsp.rm(getRepoDir(docPath), { recursive: true, force: true });
    const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    try {
      await initAstraRepo(docPath, R6_XML);
    } finally {
      errSpy.mockRestore();
    }
    expect(getSidecarHistoryWarning(getRepoDir(docPath))).toMatch(/possible history removal/);
    clearSidecarHistoryWarning(docPath);
    expect(getSidecarHistoryWarning(docPath)).toBeNull();
    expect(readWarning(await getHistory(docPath))).toBeNull();
  });

  it('restart-equivalent: clearing in-memory state then re-ensuring still warns', async () => {
    await initAstraRepo(docPath, R6_XML);
    await commitChange(docPath, R6_XML, 'second');
    await fsp.rm(getRepoDir(docPath), { recursive: true, force: true });
    await initAstraRepo(docPath, R6_XML);
    expect(getSidecarHistoryWarning(docPath)).toMatch(/possible history removal/);
    // Simulate a fresh process: drop in-memory warnings only, registry stays.
    clearSidecarHistoryWarning(docPath);
    expect(getSidecarHistoryWarning(docPath)).toBeNull();
    await ensureAstraGitRepo(docPath);
    expect(getSidecarHistoryWarning(docPath)).toMatch(/possible history removal/);
    expect(readWarning(await getHistory(docPath))).toMatch(/possible history removal/);
  });

  it('warning is enumerable and survives JSON.stringify', async () => {
    await initAstraRepo(docPath, R6_XML);
    await fsp.rm(getRepoDir(docPath), { recursive: true, force: true });
    const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    try {
      await initAstraRepo(docPath, R6_XML);
    } finally {
      errSpy.mockRestore();
    }
    const history = await getHistory(docPath);
    expect(Object.getOwnPropertyDescriptor(history, 'warning')?.enumerable).toBe(true);
    expect(JSON.stringify(history)).toMatch(/possible history removal/);
    const parsed = JSON.parse(JSON.stringify(history)) as { warning: string | null };
    expect(parsed.warning).toMatch(/possible history removal/);
  });
});
