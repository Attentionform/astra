/**
 * Table cells and composed ids over the MCP handlers: search_content hands
 * back the cell address of a match, execute_batch accepts it as a target, and
 * compose_document says which paragraph ids a saved file will carry.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip } from '@attentionform/astra-engine';
import { createMcpSessionStore } from '../../src/lib/session-store';
import { LEAN6_HANDLERS } from '../../src/lib/mcp-tool-adapter';
import { handleNewDocument } from '../../src/tools/new-document';
import { handleComposeDocument } from '../../src/tools/compose-document';
import { LEAN6_DESCRIPTIONS } from '../../src/lib/mcp/descriptions';

const SOURCE = [
  '# Quarterly figures',
  '',
  'Opening paragraph.',
  '',
  '| Region | Revenue |',
  '| --- | --- |',
  '| North | 120 |',
  '| South | 98 |',
  '',
  'Closing paragraph.',
  '',
].join('\n');

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cell-mcp-'));
});
afterEach(() => {
  fs.rmSync(dir, { recursive: true, force: true });
});

async function composed() {
  const store = createMcpSessionStore();
  const file = path.join(dir, 'figures.docx');
  const created = await handleNewDocument(store, { outPath: file, template: 'blank' });
  if (!created.ok) throw new Error(`new_document failed: ${JSON.stringify(created)}`);
  const result = await handleComposeDocument(store, { sessionId: created.sessionId, markdown: SOURCE });
  if (!result.ok) throw new Error(`compose failed: ${JSON.stringify(result)}`);
  return { store, file, sessionId: created.sessionId, result };
}

describe('search_content cell addresses', () => {
  it('lists the cell that holds the match on a table hit, and the address edits it', async () => {
    const { store, sessionId } = await composed();
    const found = (await LEAN6_HANDLERS.search_content(store, { sessionId, query: 'South' })) as {
      ok: boolean;
      matches: Array<{ id: string; cells?: string[] }>;
    };
    expect(found.ok).toBe(true);
    const hit = found.matches.find((m) => m.cells !== undefined);
    expect(hit).toBeDefined();
    expect(hit?.cells).toEqual([`${hit?.id}:r3c1`]);

    const edited = (await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'replace', target: hit?.cells?.[0], text: 'South-East' }],
    })) as { ok: boolean };
    expect(edited.ok).toBe(true);
    const xml = new AdmZip(store.get(sessionId).parser.rezip()).readAsText('word/document.xml');
    expect(xml).toContain('South-East');
    expect(xml).toContain('>North<');
  });

  it('regex search reports cells too, and paragraph hits carry none', async () => {
    const { store, sessionId } = await composed();
    const found = (await LEAN6_HANDLERS.search_content(store, { sessionId, query: 'North|South', regex: true })) as {
      ok: boolean;
      matches: Array<{ id: string; preview: string; cells?: string[] }>;
    };
    expect(found.ok).toBe(true);
    const tableHit = found.matches.find((m) => m.cells !== undefined);
    expect(tableHit?.cells).toEqual([`${tableHit?.id}:r2c1`, `${tableHit?.id}:r3c1`]);

    const prose = (await LEAN6_HANDLERS.search_content(store, { sessionId, query: 'Closing' })) as {
      matches: Array<{ cells?: string[] }>;
    };
    expect(prose.matches.length).toBeGreaterThan(0);
    expect(prose.matches.every((m) => m.cells === undefined)).toBe(true);
  });
});

describe('tool descriptions', () => {
  it('tell the caller how to address a cell', () => {
    expect(LEAN6_DESCRIPTIONS.inspect_document).toContain('<tableId>:r<row>c<col>');
    expect(LEAN6_DESCRIPTIONS.inspect_document).toContain('1-based');
    expect(LEAN6_DESCRIPTIONS.search_content).toContain('cells');
  });
});

describe('compose_document persistentIds', () => {
  it('maps each created paragraph to the id the saved file carries', async () => {
    const { store, file, sessionId, result } = await composed();
    const ids = result.persistentIds ?? {};
    const mapped = Object.keys(ids);
    expect(mapped.length).toBeGreaterThanOrEqual(3);
    for (const [sessionIdOfNode, saved] of Object.entries(ids)) {
      expect(sessionIdOfNode).toMatch(/^md_\d{4}$/);
      expect(saved).toMatch(/^[0-9A-F]{8}$/);
    }
    // The table has no paragraph id, so it is not listed.
    const created = result.createdNodeIds;
    expect(created.length).toBeGreaterThan(mapped.length);

    const saved = (await LEAN6_HANDLERS.save_document(store, { sessionId, outPath: file })) as { ok: boolean };
    expect(saved.ok).toBe(true);

    // A session opened on the saved file answers to the mapped ids.
    const reopened = (await LEAN6_HANDLERS.open_document(store, { path: file })) as { ok: boolean; sessionId: string };
    expect(reopened.ok).toBe(true);
    const handle = store.get(reopened.sessionId);
    const live = new Set(handle.ast.sections.flatMap((s) => s.nodes.map((n) => n.id)));
    for (const savedId of Object.values(ids)) expect(live.has(savedId)).toBe(true);

    // Same order as the document: the first mapped id is the heading.
    const heading = Object.entries(ids)[0];
    const headingNode = handle.ast.sections[0].nodes.find((n) => n.id === heading[1]);
    expect(String(headingNode?.metadata.text)).toBe('Quarterly figures');
  });

  it('keeps the md_ ids working in the session after a save', async () => {
    const { store, file, sessionId, result } = await composed();
    const [firstMd] = Object.keys(result.persistentIds ?? {});
    await LEAN6_HANDLERS.save_document(store, { sessionId, outPath: file });
    const edited = (await LEAN6_HANDLERS.execute_batch(store, {
      sessionId,
      operations: [{ action: 'replace', target: firstMd, text: 'Renamed heading' }],
    })) as { ok: boolean };
    expect(edited.ok).toBe(true);
  });
});
