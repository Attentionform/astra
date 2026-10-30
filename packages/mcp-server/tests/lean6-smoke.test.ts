import { describe, it, expect, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createAstraMcpServer, SERVER_NAME } from '../src/index';
import { createMcpSessionStore } from '../src/lib/session-store';
import {
  LEAN6_DESCRIPTIONS,
  LEAN6_HANDLERS,
  MCP_ENVELOPE_NOTE,
  getLean6ToolNames,
  sanitizeMcpClientMessage,
} from '../src/lib/mcp-tool-adapter';
import { DocxParser } from '@attentionform/astra-engine';

const tmpFiles: string[] = [];

function makeTmpDocx(prefix: string): string {
  const buf = DocxParser.createNewDocument('blank').rezip();
  const file = path.join(os.tmpdir(), `${prefix}-${process.pid}-${Date.now()}.docx`);
  fs.writeFileSync(file, buf);
  tmpFiles.push(file);
  return file;
}

afterEach(() => {
  while (tmpFiles.length > 0) {
    const f = tmpFiles.pop() as string;
    try {
      fs.unlinkSync(f);
    } catch {
      /* best-effort */
    }
  }
});

describe('mcp-server lean-6 surface', () => {
  it('exposes exactly the locked lean-6 tool set (open/inspect/search/batch/save/close)', () => {
    expect(getLean6ToolNames()).toEqual([
      'open_document',
      'inspect_document',
      'search_content',
      'execute_batch',
      'save_document',
      'close_document',
    ]);
    expect(Object.keys(LEAN6_HANDLERS).sort()).toEqual([...getLean6ToolNames()].sort());
  });

  it('registers core tools on the MCP server (lean-6 + export_document + visual_inspect + git history + preview, no research, no single-op)', () => {
    const server = createAstraMcpServer(createMcpSessionStore());
    const registered = (server as unknown as { _registeredTools: Record<string, unknown> })._registeredTools;
    expect(SERVER_NAME).toBe('astra-mcp-server');
    expect(Object.keys(registered).sort()).toEqual([
      'close_document',
      'compose_document',
      'execute_batch',
      'export_document',
      'get_document_history',
      'get_stats',
      'import_document',
      'inspect_document',
      'new_document',
      'normalize_document',
      'open_document',
      'preview_revert',
      'redo_document',
      'revert_document',
      'sanitize_document',
      'save_document',
      'search_content',
      'validate_document',
      'visual_inspect',
    ]);
  });

  it('open_document rejects invalid input fail-closed (E_VALIDATION, ok:false)', async () => {
    const store = createMcpSessionStore();
    const res = await LEAN6_HANDLERS.open_document(store, { path: '' });
    expect(res.ok).toBe(false);
    if (!res.ok) {
      expect(res.code).toBe('E_VALIDATION');
      expect(typeof res.remedy).toBe('string');
    }
  });

  it('full lean-6 dispatch smoke: open -> inspect -> search -> batch(invalid) -> save -> close', async () => {
    const store = createMcpSessionStore();
    const src = makeTmpDocx('astra-mcp-open');
    const out = path.join(os.tmpdir(), `astra-mcp-save-${process.pid}-${Date.now()}.docx`);
    tmpFiles.push(out);

    const opened = await LEAN6_HANDLERS.open_document(store, { path: src });
    expect(opened.ok).toBe(true);
    if (!opened.ok) return;
    const sessionId = opened.sessionId;
    expect(typeof sessionId).toBe('string');

    const inspected = await LEAN6_HANDLERS.inspect_document(store, { sessionId });
    expect(inspected.ok).toBe(true);
    if (inspected.ok) {
      expect(typeof inspected.toonText).toBe('string');
      expect(inspected.toonText).toContain('fidelity');
      expect(inspected.rowCount).toBeGreaterThanOrEqual(0);
    }

    const searched = await LEAN6_HANDLERS.search_content(store, { sessionId, query: 'a' });
    expect(searched.ok).toBe(true);
    if (searched.ok) expect(Array.isArray(searched.matches)).toBe(true);

    // Unknown action must fail closed with remedy (never applied).
    const badBatch = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'definitely_not_a_real_op' }],
    });
    expect(badBatch.ok).toBe(false);
    if (!badBatch.ok) {
      expect(['E_VALIDATION', 'E_QUOTA']).toContain(badBatch.code);
    }

    const saved = await LEAN6_HANDLERS.save_document(store, { sessionId, outPath: out });
    expect(saved.ok).toBe(true);
    expect(fs.existsSync(out)).toBe(true);

    const closed = await LEAN6_HANDLERS.close_document(store, { sessionId });
    expect(closed.ok).toBe(true);
    if (!closed.ok) return;
    expect(closed.sessionId).toBe(sessionId);
    expect(closed.dirty).toBe(false);
    expect(closed.discardedRevisions).toBe(0);
    expect(closed.warning).toBeUndefined();

    // Double close maps to E_SESSION (already closed), never throws.
    const closedAgain = await LEAN6_HANDLERS.close_document(store, { sessionId });
    expect(closedAgain.ok).toBe(false);
    if (!closedAgain.ok) expect(closedAgain.code).toBe('E_SESSION');
  });

  it('executes visual_inspect via MCP server (telemetry-only default and opt-in images withVisual:true)', async () => {
    const store = createMcpSessionStore();
    const server = createAstraMcpServer(store);
    const registered = (
      server as unknown as {
        _registeredTools: Record<
          string,
          { handler: (args: unknown) => Promise<{ content: Array<{ type: string; text: string }>; isError?: boolean }> }
        >;
      }
    )._registeredTools;
    const visualTool = registered['visual_inspect'];
    expect(visualTool).toBeDefined();

    const src = makeTmpDocx('astra-mcp-visual');
    const opened = await LEAN6_HANDLERS.open_document(store, { path: src });
    expect(opened.ok).toBe(true);
    if (!opened.ok) return;
    const sessionId = opened.sessionId;

    // 1. Default (withVisual omitted / false) -> telemetry-only (<500 tokens, wireframe, no images)
    const resDefault = await visualTool.handler({ sessionId });
    expect(resDefault.isError).toBeFalsy();
    const dataDefault = JSON.parse(resDefault.content[0].text);
    expect(dataDefault.ok).toBe(true);
    expect(dataDefault.telemetry).toBeDefined();
    expect(typeof dataDefault.telemetry.wireframe).toBe('string');
    expect(dataDefault.telemetry.tokens).toBeLessThanOrEqual(500);
    expect(dataDefault.images).toBeUndefined();

    // 2. Explicit withVisual: false -> telemetry-only
    const resNoVisual = await visualTool.handler({ sessionId, withVisual: false });
    expect(resNoVisual.isError).toBeFalsy();
    const dataNoVisual = JSON.parse(resNoVisual.content[0].text);
    expect(dataNoVisual.ok).toBe(true);
    expect(dataNoVisual.telemetry).toBeDefined();
    expect(dataNoVisual.images).toBeUndefined();

    // 3. Opt-in withVisual: true -> renders file-path images (or E_DEPENDENCY if host lacks binaries)
    const resVisual = await visualTool.handler({ sessionId, withVisual: true });
    const dataVisual = JSON.parse(resVisual.content[0].text);
    if (dataVisual.ok) {
      expect(resVisual.isError).toBeFalsy();
      expect(Array.isArray(dataVisual.images)).toBe(true);
      expect(dataVisual.images.length).toBeGreaterThanOrEqual(1);
      for (const img of dataVisual.images) {
        expect(img.path.endsWith('.png')).toBe(true);
        expect(fs.existsSync(img.path)).toBe(true);
      }
      expect(dataVisual.outDir).toBeTruthy();
    } else {
      expect(resVisual.isError).toBe(true);
      expect(dataVisual.code).toBe('E_DEPENDENCY');
    }

    await LEAN6_HANDLERS.close_document(store, { sessionId });
  }, 60000);

  it('envelope: tool docs carry no ticket notes; execute_batch client message strips stack + /home paths', async () => {
    expect(MCP_ENVELOPE_NOTE).toBe('');
    expect(LEAN6_DESCRIPTIONS.execute_batch).not.toMatch(/-32602|P2-1/);
    const raw = 'execute_batch failed: Error: Node p_999999 not found in AST\n at executeReplace (/app/packages/mcp-server/dist/index.js:4352:11)\n at OperationEngine.routeOperation (/app/dist/index.js:1:2)';
    const clean = sanitizeMcpClientMessage(raw);
    expect(clean).not.toMatch(/\bat\s+/);
    expect(clean).not.toMatch(/\/home\//);
    expect(clean).toMatch(/Node p_999999 not found/);
    // Handler path: bad target returns sanitized E_VALIDATION (no stack/paths)
    const store = createMcpSessionStore();
    const src = makeTmpDocx('astra-mcp-sanitize');
    const opened = await LEAN6_HANDLERS.open_document(store, { path: src });
    expect(opened.ok).toBe(true);
    if (!opened.ok) return;
    const bad = await LEAN6_HANDLERS.execute_batch(store, {
      sessionId: opened.sessionId,
      operations: [{ action: 'replace', target: 'p_999999', text: 'x' }],
    });
    expect(bad.ok).toBe(false);
    if (!bad.ok) {
      expect(bad.code).toBe('E_VALIDATION');
      expect(bad.message).not.toMatch(/\bat\s+/);
      expect(bad.message).not.toMatch(/\/home\//);
      expect(typeof bad.remedy).toBe('string');
    }
    await LEAN6_HANDLERS.close_document(store, { sessionId: opened.sessionId });
  });
});
