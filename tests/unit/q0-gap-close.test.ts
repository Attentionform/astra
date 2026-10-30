/**
 * Leftover narrowing close-out — gap-close proof.
 *
 * GAP 1: MCP compose/new-document error tables still narrow (KNOWN_COMPOSE_CODES
 * dropped E_LIMIT/E_SECURITY/E_DEPENDENCY/E_AUTH; new-document 5-code union
 * same effect). Both now resolve via the shared engine table
 * (getEngineErrorCode from @attentionform/astra-engine). Every engine throw survives both
 * paths verbatim.
 *
 * GAP 2: renderer jail missing win32 branch. renderer.ts now imports
 * isPathInsideDir from sandbox/paths (single source of truth). This test
 * pins old-vs-new convergence incl. the win32 branch + live traversal
 * refusal via resolveJailedVisualOutDir.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  ENGINE_ERROR_CODES,
  getEngineErrorCode,
  isPathInsideDir,
  toError,
  type EngineErrorCode,
} from '@attentionform/astra-engine';
import { DocxParser } from '@attentionform/astra-engine';
import { getVisualArtifactsDir, resolveJailedVisualOutDir } from '@attentionform/astra-engine';
import { composeErrorFrom, handleComposeDocument } from '../../packages/mcp-server/src/tools/compose-document';
import { handleNewDocument, resolveNewDocumentCode } from '../../packages/mcp-server/src/tools/new-document';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';

const ALL: readonly EngineErrorCode[] = ENGINE_ERROR_CODES;

function codedErr(code: EngineErrorCode): Error {
  return toError(code, `probe ${code}`);
}
function messageOnlyErr(code: EngineErrorCode): Error {
  return new Error(`[${code}] probe failure without a code prop`);
}
function codePropOnlyErr(code: EngineErrorCode): Error {
  return Object.assign(new Error('probe without code text'), { code });
}

/** Narrow POSIX-only copy of the removed renderer local (no win32 branch). */
function oldNarrowIsInside(candidate: string, dir: string): boolean {
  const normC = path.normalize(candidate);
  const normD = path.normalize(dir);
  return normC === normD || normC.startsWith(normD + path.sep);
}

describe('Q0 gap-close: compose/new-document widened codes', () => {
  it('composeErrorFrom preserves all 12 codes (coded, message-only, code-prop-only)', () => {
    expect([...ALL].sort()).toEqual(
      [
        'E_VALIDATION', 'E_SESSION', 'E_QUOTA', 'E_IO', 'E_CONFLICT',
        'E_NOT_FOUND', 'E_SANDBOX', 'E_LIMIT', 'E_SECURITY', 'E_STATE',
        'E_DEPENDENCY', 'E_AUTH',
      ].sort()
    );
    for (const code of ALL) {
      expect(composeErrorFrom(codedErr(code), 'fallback').code, `coded ${code}`).toBe(code);
      expect(composeErrorFrom(messageOnlyErr(code), 'fallback').code, `message-only ${code}`).toBe(code);
      expect(composeErrorFrom(codePropOnlyErr(code), 'fallback').code, `code-prop ${code}`).toBe(code);
    }
    // Uncoded fallback unchanged.
    expect(composeErrorFrom(new Error('boom'), 'fallback').code).toBe('E_VALIDATION');
  });

  it('FOCUS: previously-dropped E_LIMIT/E_SECURITY/E_DEPENDENCY/E_AUTH survive compose', () => {
    for (const code of ['E_LIMIT', 'E_SECURITY', 'E_DEPENDENCY', 'E_AUTH'] as const) {
      expect(composeErrorFrom(codedErr(code), 'fallback').code).toBe(code);
      expect(composeErrorFrom(messageOnlyErr(code), 'fallback').code).toBe(code);
    }
  });

  it('resolveNewDocumentCode preserves all 12 codes with per-site fallbacks', () => {
    for (const code of ALL) {
      expect(resolveNewDocumentCode(codedErr(code), 'E_IO'), `coded ${code}`).toBe(code);
      expect(resolveNewDocumentCode(messageOnlyErr(code), 'E_IO'), `message-only ${code}`).toBe(code);
      expect(resolveNewDocumentCode(codePropOnlyErr(code), 'E_IO'), `code-prop ${code}`).toBe(code);
      expect(resolveNewDocumentCode(codedErr(code), 'E_SANDBOX'), `sandbox-fallback ${code}`).toBe(code);
      expect(resolveNewDocumentCode(codedErr(code), 'E_SESSION'), `session-fallback ${code}`).toBe(code);
      expect(resolveNewDocumentCode(codedErr(code), 'E_VALIDATION'), `validation-fallback ${code}`).toBe(code);
    }
    expect(resolveNewDocumentCode(new Error('boom'), 'E_IO')).toBe('E_IO');
    expect(resolveNewDocumentCode(new Error('boom'), 'E_SANDBOX')).toBe('E_SANDBOX');
    expect(resolveNewDocumentCode(new Error('boom'), 'E_SESSION')).toBe('E_SESSION');
  });

  it('LIVE compose path: E_LIMIT via withSession survives verbatim', async () => {
    const store = createMcpSessionStore();
    const buf = DocxParser.createNewDocument('blank').rezip();
    const handle = await store.openBuffer(buf, 'memory.docx');
    const orig = store.withSession.bind(store);
    store.withSession = (async () => {
      throw toError('E_LIMIT', 'probe E_LIMIT over cap');
    }) as typeof store.withSession;
    try {
      const res = await handleComposeDocument(store, { sessionId: handle.sessionId, markdown: '# Hi\n\nBody.\n' });
      expect(res.ok).toBe(false);
      if (res.ok) throw new Error('expected failure');
      expect(res.code).toBe('E_LIMIT');
    } finally {
      store.withSession = orig;
    }
  });

  it('LIVE new-document path: E_LIMIT via openDocument survives verbatim', async () => {
    const store = createMcpSessionStore();
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-newdoc-gap-'));
    const target = path.join(dir, 'gap.docx');
    const orig = store.openDocument.bind(store);
    store.openDocument = (async () => {
      throw toError('E_LIMIT', 'probe E_LIMIT on open');
    }) as typeof store.openDocument;
    try {
      const res = await handleNewDocument(store, { outPath: target });
      expect(res.ok).toBe(false);
      if (res.ok) throw new Error('expected failure');
      expect(res.code).toBe('E_LIMIT');
    } finally {
      store.openDocument = orig;
      try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* ignore */ }
    }
  });
});

describe('Q0 gap-close: renderer jail converges with shared module', () => {
  it('shared jail matches narrow local on POSIX corpus, adds win32 branch', () => {
    const dir = path.resolve(os.tmpdir(), 'q0-gap-jail-root');
    const corpus: Array<[string, string]> = [
      [path.join(dir, 'a.docx'), dir],
      [path.join(dir, 'sub', 'b.docx'), dir],
      [path.join(os.tmpdir(), 'outside.docx'), dir],
      [dir, dir],
      [path.join(dir, '..', 'evil.docx'), dir],
      [dir + '2/x.docx', dir],
    ];
    for (const [candidate, d] of corpus) {
      expect(isPathInsideDir(candidate, d, process.platform)).toBe(oldNarrowIsInside(candidate, d));
    }
    // win32 branch: case-insensitive containment (narrow local lacks it).
    // NOTE: forward slashes so the branch is observable on a POSIX host
    // (path.sep is '/' here; backslash cases agree trivially on both sides).
    const winDir = '/Tmp/JailRoot';
    const winCand = '/tmp/jailroot/sub/a.docx';
    expect(isPathInsideDir(winCand, winDir, 'win32')).toBe(true);
    expect(oldNarrowIsInside(winCand, winDir)).toBe(false);
    // getEngineErrorCode sanity (shared table importable from engine root).
    expect(getEngineErrorCode(codedErr('E_SECURITY'))).toBe('E_SECURITY');
  });

  it('live: traversal outDir refused with E_SANDBOX, valid subdir accepted', () => {
    expect(() => resolveJailedVisualOutDir('/etc/passwd-visual-escape')).toThrow(/E_SANDBOX/);
    expect(() => resolveJailedVisualOutDir('../escape')).toThrow(/E_SANDBOX/);
    // Valid run-scoped subdir resolves inside the artifacts base.
    const runDir = getVisualArtifactsDir('gap-close-valid');
    const jailed = resolveJailedVisualOutDir(runDir);
    expect(isPathInsideDir(jailed, runDir)).toBe(true);
  });
});
