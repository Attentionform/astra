/**
 * Engine library log-silence: library calls emit zero console output.
 * Typed channels (warnings[], skipped[], errors + rollback flags, history
 * warning fields, trust registry) carry the signal instead. CLI remains the
 * logging surface and is intentionally not asserted here.
 */
import { describe, it, expect, vi, afterEach } from 'vitest';
import path from 'path';
import fsp from 'fs/promises';
import os from 'os';
import {
  DocxParser,
  createBlankDocxBuffer,
  OperationEngine,
  executeTable,
  executeDelete,
  executeReplace,
  executeInsert,
  initAstraRepo,
  getRepoDir,
  getSidecarHistoryWarning,
  clearSidecarHistoryWarning,
  registerTrustedDocPath,
  listTrustedDocPaths,
  clearTrustedDocPaths,
} from '@attentionform/astra-engine';

function silence() {
  const log = vi.spyOn(console, 'log').mockImplementation(() => undefined);
  const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
  const error = vi.spyOn(console, 'error').mockImplementation(() => undefined);
  const info = vi.spyOn(console, 'info').mockImplementation(() => undefined);
  const debug = (typeof (console as unknown as Record<string, unknown>).debug === 'function'
    ? vi.spyOn(console as unknown as { debug: (...a: unknown[]) => void }, 'debug').mockImplementation(() => undefined)
    : null);
  return {
    spies: [log, warn, error, info, ...(debug ? [debug] : [])],
    calls() {
      return [log, warn, error, info, ...(debug ? [debug] : [])].reduce((n, s) => n + s.mock.calls.length, 0);
    },
    restore() {
      for (const s of [log, warn, error, info, ...(debug ? [debug] : [])]) s.mockRestore();
    },
  };
}

afterEach(() => {
  vi.restoreAllMocks();
});

async function freshTable(): Promise<{ parser: DocxParser; ast: any; tableId: string }> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id;
  await executeInsert({
    action: 'insert', type: 'table', anchor, position: 'after',
    headers: ['A', 'B'], rows: [['1', '2'], ['3', '4']],
  } as any, ast, parser);
  const tbl = ast.sections[0].nodes.find((n: any) => n.type === 'table');
  expect(tbl).toBeDefined();
  return { parser, ast, tableId: tbl.id as string };
}

describe('engine library emits zero console output', () => {
  it('table contrast warning rides in warnings[], not on console', async () => {
    const { ast, tableId } = await freshTable();
    const gate = silence();
    try {
      const res: any = await executeTable({
        action: 'table', target: tableId,
        params: { op: 'setCellShading', rowIndex: 1, colIndex: 1, fill: '000080' },
      } as any, ast);
      expect(res.warnings.length).toBe(1);
      expect(res.warnings[0]).toMatch(/contrast/);
      expect(gate.calls()).toBe(0);
    } finally {
      gate.restore();
    }
  });

  it('idempotent delete miss rides in skipped[], not on console', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const gate = silence();
    try {
      const res = executeDelete({ action: 'delete', target: 'ghost_silent' } as any, ast, parser);
      expect(res.skipped).toEqual(['ghost_silent']);
      expect(res.deleted).toEqual([]);
      expect(gate.calls()).toBe(0);
    } finally {
      gate.restore();
    }
  });

  it('strict delete miss throws typed error with zero console output', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const gate = silence();
    try {
      expect(() =>
        executeDelete({ action: 'delete', target: 'ghost_strict', strict: true } as any, ast, parser)
      ).toThrow(/not found for deletion/);
      expect(gate.calls()).toBe(0);
    } finally {
      gate.restore();
    }
  });

  it('replace returns typed warnings with zero console output', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const target = ast.sections[0].nodes[0].id;
    const gate = silence();
    try {
      const out: any = executeReplace({ action: 'replace', target, text: 'silent hello' } as any, ast);
      expect(Array.isArray(out.warnings)).toBe(true);
      expect(ast.sections[0].nodes[0].metadata.text).toBe('silent hello');
      expect(gate.calls()).toBe(0);
    } finally {
      gate.restore();
    }
  });

  it('batch failure records typed errors + rollback with zero console output', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const gate = silence();
    try {
      const res = await engine.executeBatch({
        operations: [
          { action: 'delete', target: 'ghost_before_fail' } as any,
          { action: 'delete', target: 'ghost_boom', strict: true } as any,
        ],
        sequential: true,
      });
      expect(res.success).toBe(false);
      expect(res.rollbackTriggered).toBe(true);
      expect(res.errors.length).toBeGreaterThan(0);
      expect(gate.calls()).toBe(0);
    } finally {
      gate.restore();
    }
  });

  it('parser integrity check stays silent', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const gate = silence();
    try {
      const out = parser.validateZipIntegrity();
      expect(out.isValid).toBe(true);
      expect(gate.calls()).toBe(0);
    } finally {
      gate.restore();
    }
  });

  it('history-removal signal rides in the typed warning field, not on console', async () => {
    const tmpDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-silence-'));
    const docPath = path.join(tmpDir, 'silent.docx');
    const xml = '<w:document><w:body><w:p>Hi</w:p></w:body></w:document>';
    try {
      await initAstraRepo(docPath, xml);
      await fsp.rm(getRepoDir(docPath), { recursive: true, force: true });
      const gate = silence();
      try {
        await initAstraRepo(docPath, xml);
        expect(gate.calls()).toBe(0);
      } finally {
        gate.restore();
      }
      expect(getSidecarHistoryWarning(docPath)).toMatch(/possible history removal/);
      clearSidecarHistoryWarning(docPath);
    } finally {
      await fsp.rm(tmpDir, { recursive: true, force: true });
    }
  });

  it('trust registration records in the registry with zero console output', async () => {
    const tmpDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-silence-trust-'));
    const candidate = path.join(tmpDir, 'trusted.docx');
    await fsp.writeFile(candidate, 'x');
    const gate = silence();
    try {
      const resolved = registerTrustedDocPath(candidate, 'dialog');
      expect(typeof resolved).toBe('string');
      expect(listTrustedDocPaths().some((e) => e.path === resolved)).toBe(true);
      expect(gate.calls()).toBe(0);
    } finally {
      gate.restore();
      clearTrustedDocPaths();
      await fsp.rm(tmpDir, { recursive: true, force: true });
    }
  });
});
