import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, revisionsRemedy } from '@attentionform/astra-engine';

const STAMP = 'w:author="Alice" w:date="2026-01-01T00:00:00Z"';
// A table property exception revision the resolver reports as unresolved.
const UNRESOLVABLE = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tblPrEx><w:tPrChange w:id="7" ${STAMP}><w:tblPrEx/></w:tPrChange></w:tblPrEx><w:tc><w:p/></w:tc></w:tr></w:tbl>`;
const RESOLVABLE = `<w:p><w:ins w:id="1" ${STAMP}><w:r><w:t>added</w:t></w:r></w:ins></w:p>`;
// A whole-paragraph move: its elements are the moveFrom/moveTo wrappers and the paragraph mark markers.
const MOVE = `<w:p><w:pPr><w:rPr><w:moveFrom w:id="11" ${STAMP}/></w:rPr></w:pPr><w:moveFromRangeStart w:id="10" w:name="m1" ${STAMP}/><w:moveFrom w:id="12" ${STAMP}><w:r><w:t>moved</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="10"/></w:p>`
  + `<w:p><w:pPr><w:rPr><w:moveTo w:id="14" ${STAMP}/></w:rPr></w:pPr><w:moveToRangeStart w:id="13" w:name="m1" ${STAMP}/><w:moveTo w:id="15" ${STAMP}><w:r><w:t>moved</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="13"/></w:p>`;

async function open(body: string) {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  const parser = new DocxParser(zip.toBuffer());
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
}
const run = (engine: OperationEngine, operation: string) => engine.executeBatch({ operations: [{ action: 'trackedChange', operation }] } as never);

describe('revisionsRemedy without a resolution (legacy wording)', () => {
  it('still names acceptAll, trackedChange and --accept-all-revisions', () => {
    const text = revisionsRemedy(3);
    expect(text).toContain('acceptAll');
    expect(text).toContain('trackedChange');
    expect(text).toContain('--accept-all-revisions');
    expect(revisionsRemedy(0)).toBeUndefined();
  });
});

describe('FX-07: remedy after a resolution that left revisions behind', () => {
  it('a second acceptAll on an unresolvable document does not tell the agent to run acceptAll', async () => {
    const { engine } = await open(`${UNRESOLVABLE}<w:p/>`);
    for (const round of [1, 2]) {
      const result = await run(engine, 'acceptAll');
      expect(result.success, `round ${round}`).toBe(true);
      expect(result.revisionsPending).toBe(1);
      expect(result.revisionsUnresolved).toBe(1);
      expect(result.revisionsRemedy).toBeDefined();
      expect(result.revisionsRemedy).not.toMatch(/run \{ action: 'trackedChange', operation: 'acceptAll' \}/);
      expect(result.revisionsRemedy).toMatch(/cannot be resolved automatically/);
      expect(result.revisionsRemedy).toContain('tPrChange');
      expect(result.revisionsRemedy).toMatch(/not (?:resolve|help)|again/i);
    }
  });

  it('a mixed document reports the resolved part and only the stuck part as unresolved', async () => {
    const { engine } = await open(`${RESOLVABLE}${UNRESOLVABLE}<w:p/>`);
    expect(engine.countPendingRevisions()).toBe(2);
    const result = await run(engine, 'acceptAll');
    expect(result.revisionResolution?.resolved).toBe(1);
    expect(result.revisionsPending).toBe(1);
    expect(result.revisionsUnresolved).toBe(1);
    expect(result.revisionsRemedy).toContain('tPrChange');
  });

  it('the unresolved count never exceeds the pending count', async () => {
    const { engine } = await open(`${RESOLVABLE}${UNRESOLVABLE}<w:p/>`);
    const result = await run(engine, 'rejectAll');
    expect(result.revisionsUnresolved ?? 0).toBeLessThanOrEqual(result.revisionsPending);
  });

  it('when everything resolves there is no remedy and no unresolved count', async () => {
    const { engine } = await open(`${RESOLVABLE}<w:p/>`);
    const result = await run(engine, 'acceptAll');
    expect(result.revisionsPending).toBe(0);
    expect(result.revisionsRemedy).toBeUndefined();
    expect(result.revisionsUnresolved).toBeUndefined();
  });

  it('a batch that never resolves keeps the legacy remedy that names acceptAll', async () => {
    const { engine } = await open(`${RESOLVABLE}<w:p/>`);
    const result = await engine.executeBatch({ operations: [{ action: 'metadata', operation: 'setProperties', title: 'x' }] } as never);
    expect(result.revisionsPending).toBe(1);
    expect(result.revisionsUnresolved).toBeUndefined();
    expect(result.revisionsRemedy).toContain('acceptAll');
  });

  it('two resolving operations in one batch do not double count the same stuck revision', async () => {
    const { engine } = await open(`${UNRESOLVABLE}<w:p/>`);
    const result = await engine.executeBatch({ operations: [
      { action: 'trackedChange', operation: 'acceptAll' },
      { action: 'trackedChange', operation: 'acceptAll' },
    ] } as never);
    expect(result.revisionsPending).toBe(1);
    expect(result.revisionsUnresolved).toBe(1);
  });
});

describe('revisionsPending is the honest number of revision elements (FX-07)', () => {
  it('counts a whole-paragraph move as each of its revision elements', async () => {
    const { engine } = await open(`${MOVE}<w:p/>`);
    // two paragraph-mark markers and two wrappers; range markers are not revisions
    expect(engine.countPendingRevisions()).toBe(4);
    const result = await engine.executeBatch({ operations: [{ action: 'metadata', operation: 'setProperties', title: 'x' }] } as never);
    expect(result.revisionsPending).toBe(4);
  });

  it('resolving the move clears every one of those elements', async () => {
    const { engine } = await open(`${MOVE}<w:p/>`);
    const result = await run(engine, 'acceptAll');
    expect(result.success).toBe(true);
    expect(result.revisionsPending).toBe(0);
  });
});

describe('FX-07 over the MCP handler', () => {
  it('passes revisionsUnresolved and the unresolved-aware remedy through execute_batch', async () => {
    const { createMcpSessionStore } = await import('../../packages/mcp-server/src/lib/session-store');
    const { handleExecuteBatch } = await import('../../packages/mcp-server/src/lib/mcp-tool-adapter');
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${UNRESOLVABLE}<w:p/>$1`)));
    const store = createMcpSessionStore();
    const handle = await store.openBuffer(zip.toBuffer(), 'stuck.docx');
    const out = await handleExecuteBatch(store, { sessionId: handle.sessionId, operations: [{ action: 'trackedChange', operation: 'acceptAll' }] } as never) as Record<string, unknown>;
    expect(out.ok, JSON.stringify(out)).toBe(true);
    expect(out.revisionsPending).toBe(1);
    expect(out.revisionsUnresolved).toBe(1);
    expect(String(out.revisionsRemedy)).toContain('tPrChange');
    expect(String(out.revisionsRemedy)).not.toMatch(/run \{ action: 'trackedChange', operation: 'acceptAll' \}/);
  });
});
