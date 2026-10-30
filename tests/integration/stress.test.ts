import { describe, it, expect } from 'vitest';
import {
  DocumentAST,
  Node,
  OperationEngine,
  Section,
  ToonGenerator,
} from '@attentionform/astra-engine';

function buildLargeAST(sectionCount: number, nodesPerSection: number): DocumentAST {
  const sections: Section[] = [];
  let idCounter = 0;
  for (let s = 0; s < sectionCount; s++) {
    const nodes: Node[] = [];
    for (let p = 0; p < nodesPerSection; p++) {
      idCounter++;
      nodes.push({
        id: `p_${idCounter}`,
        type: 'paragraph',
        version: 1,
        checksum: 'dummy_hash',
        metadata: {
          text: `Paragraph ${p} in section ${s}. `.repeat(5),
          hasCitation: false, hasComment: false, hasBookmark: false,
          hasTrackedChanges: false, hasEquation: false, hasField: false,
          hasImage: false, hasHyperlink: false,
        },
        xml: {} as any,
      } as Node);
    }
    sections.push({
      id: `sec_${s}`, nodes,
      heading: null, pageSetup: null, version: 1, checksum: 'dummy',
    } as Section);
  }
  return {
    metadata: {}, sections, styles: {}, numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {}, headers: {}, footers: {}, mediaFiles: {},
  } as DocumentAST;
}

describe('Phase 8.3 — Stress Testing', () => {

  it('should generate TOON for a 50-page document (1000 nodes) in <500ms', () => {
    const ast = buildLargeAST(50, 20);
    const toon = new ToonGenerator();

    const start = performance.now();
    const result = toon.generateToon(ast);
    const elapsed = performance.now() - start;

    expect(result.rows.length).toBe(1000);
    expect(elapsed).toBeLessThan(500);
  });

  it('should execute a 100-operation batch in <10 seconds', async () => {
    const ast = buildLargeAST(1, 200);
    const engine = new OperationEngine(ast);
    const ops: any[] = [];
    for (let i = 1; i <= 100; i++) {
      ops.push({ action: 'replace', target: `p_${i}`, expectedVersion: 1, text: `Replaced ${i}` });
    }

    const start = performance.now();
    const res = await engine.executeBatch({ operations: ops as any, sequential: false });
    const elapsed = performance.now() - start;

    expect(elapsed).toBeLessThan(10000);
    // Even if individual ops fail (due to mock xml), the engine should handle them
    expect(res.appliedCount + res.errors.length).toBe(100);
  });

  it('should produce consistent TOON diffs after edits', () => {
    const ast = buildLargeAST(1, 10);
    const toon = new ToonGenerator();

    const before = toon.generateToon(ast);

    // Simulate a version bump on node p_3
    const node = ast.sections[0].nodes[2];
    node.version = 2;
    node.metadata.text = 'Updated text';

    const after = toon.generateToon(ast);
    const diff = toon.generateDiff(before, after);

    expect(diff.rows.length).toBe(1);
    expect(diff.rows[0].id).toBe('p_3');
    expect(diff.rows[0].version).toBe(2);
  });
});
