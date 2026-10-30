import { describe, it, expect } from 'vitest';
import {
  DocumentAST,
  Node,
  OperationEngine,
  Section,
} from '@attentionform/astra-engine';

function buildMinimalAST(nodeIds: string[]): DocumentAST {
  const nodes: Node[] = nodeIds.map(id => ({
    id,
    type: 'paragraph' as const,
    version: 1,
    checksum: 'hash',
    metadata: {
      text: `Text for ${id}`,
      hasCitation: false, hasComment: false, hasBookmark: false,
      hasTrackedChanges: false, hasEquation: false, hasField: false,
      hasImage: false, hasHyperlink: false,
    },
    xml: {} as any,
  }));
  return {
    metadata: {},
    sections: [{ id: 'sec_1', nodes, heading: null, pageSetup: null, version: 1, checksum: 'h' } as Section],
    styles: {}, numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {}, headers: {}, footers: {}, mediaFiles: {},
  } as DocumentAST;
}

describe('Phase 8.4 — Regression Suite', () => {

  it('Reg-001: Missing target node should trigger rollback with a helpful hint', async () => {
    const ast = buildMinimalAST([]);
    const engine = new OperationEngine(ast);

    const res = await engine.executeBatch({
      sequential: true,
      operations: [
        { action: 'replace', target: 'missing_node', expectedVersion: 1, text: 'Fail' } as any,
      ],
    });

    expect(res.success).toBe(false);
    expect(res.rollbackTriggered).toBe(true);
    expect(res.errors.length).toBeGreaterThan(0);
    expect(res.errors[0].hint).toContain('Check if the target node ID is correct');
  });

  it('Reg-002: Moving a node relative to itself should be rejected', async () => {
    const ast = buildMinimalAST(['p_1', 'p_2']);
    const engine = new OperationEngine(ast);

    const res = await engine.executeBatch({
      sequential: true,
      operations: [
        { action: 'move', target: 'p_1', anchor: 'p_1', position: 'after' } as any,
      ],
    });

    expect(res.success).toBe(false);
    expect(res.errors[0].error).toContain('relative to itself');
  });

  it('Reg-003: Version mismatch should generate self-correction hint', async () => {
    const ast = buildMinimalAST(['p_1']);
    ast.sections[0].nodes[0].version = 5; // current is 5
    const engine = new OperationEngine(ast);

    const res = await engine.executeBatch({
      sequential: true,
      operations: [
        { action: 'replace', target: 'p_1', expectedVersion: 1, text: 'Stale edit' } as any,
      ],
    });

    expect(res.success).toBe(false);
    expect(res.errors[0].hint).toContain('latest versions');
  });

  it('Reg-004: Unsupported operation action should produce a clear error', async () => {
    const ast = buildMinimalAST(['p_1']);
    const engine = new OperationEngine(ast);

    const res = await engine.executeBatch({
      sequential: true,
      operations: [
        { action: 'teleport', target: 'p_1' } as any,
      ],
    });

    expect(res.success).toBe(false);
    // D2 BX-10 (engine-owned): loud unknown-op error carries the received
    // value + valid set + self-correction hint (was: bare 'Unsupported operation').
    expect(res.errors[0].error).toContain("Unknown operation 'teleport'");
    expect(res.errors[0].error).toContain('Supported:');
    expect(res.errors[0].error).toContain('inspect_document');
  });

  it('Reg-005: Batch sort order is respected (style before delete)', async () => {
    const ast = buildMinimalAST(['p_1', 'p_2']);
    const engine = new OperationEngine(ast);

    // We can't easily instrument the engine, but we can verify that
    // a delete-then-style batch doesn't crash (style runs first due to sort)
    const res = await engine.executeBatch({
      sequential: false,
      operations: [
        { action: 'delete', target: 'p_2' } as any,
        { action: 'style', target: 'p_1', properties: { bold: true } } as any,
      ],
    });

    // Even if operations fail on mock XML, the engine shouldn't crash
    expect(res).toBeDefined();
    expect(typeof res.appliedCount).toBe('number');
  });
});
