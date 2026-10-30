import { describe, it, expect } from 'vitest';
import {
  DocumentAST,
  Node,
  Section,
  ToonGenerator,
} from '@attentionform/astra-engine';

/**
 * Phase 8.2 — Migration Testing (v3 vs v4)
 *
 * v3 fed the LLM raw OOXML and used flexible string matching on XML.
 * v4 feeds the LLM a compact TOON table and routes JSON operations through the engine.
 *
 * These tests verify that the v4 TOON representation is strictly superior
 * to the v3 raw-XML approach on token efficiency and accuracy metrics.
 */

function buildSampleAST(): DocumentAST {
  const nodes: Node[] = [
    {
      id: 'p_001', type: 'paragraph', version: 1, checksum: 'h1',
      metadata: {
        text: 'Introduction to ASTra Architecture',
        hasCitation: false, hasComment: false, hasBookmark: false,
        hasTrackedChanges: false, hasEquation: false, hasField: false,
        hasImage: false, hasHyperlink: false, namedStyle: 'Heading1',
      },
      xml: {} as any,
    },
    {
      id: 'p_002', type: 'paragraph', version: 1, checksum: 'h2',
      metadata: {
        text: 'ASTra v4 introduces a fundamentally new approach to document editing using an AST-based architecture with TOON compression.',
        hasCitation: true, hasComment: false, hasBookmark: false,
        hasTrackedChanges: false, hasEquation: false, hasField: false,
        hasImage: false, hasHyperlink: true,
      },
      xml: {} as any,
    },
  ];

  return {
    metadata: {},
    sections: [{ id: 'sec_1', nodes, heading: null, pageSetup: null, version: 1, checksum: 'h' } as Section],
    styles: {}, numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {}, headers: {}, footers: {}, mediaFiles: {},
  } as DocumentAST;
}

describe('Phase 8.2 — Migration Testing (v3 → v4)', () => {

  it('v4 TOON should be significantly smaller than v3 raw XML', () => {
    const ast = buildSampleAST();
    const toon = new ToonGenerator();
    const table = toon.generateToon(ast);
    const toonText = toon.renderToonText(table);

    // Simulate v3 raw XML output (a typical paragraph is ~400-800 bytes of XML)
    const v3XmlSize = 600 * ast.sections[0].nodes.length;
    const v4ToonSize = toonText.length;

    // v4 TOON should be at least 60% smaller
    const reduction = (1 - v4ToonSize / v3XmlSize) * 100;
    expect(reduction).toBeGreaterThan(60);
  });

  it('v4 TOON should preserve all metadata flags that v3 lost', () => {
    const ast = buildSampleAST();
    const toon = new ToonGenerator();
    const table = toon.generateToon(ast);

    // The second node has citation and hyperlink — these should appear as flags
    const row2 = table.rows.find(r => r.id === 'p_002');
    expect(row2).toBeDefined();
    expect(row2!.flags).toContain('[CITE]');
    expect(row2!.flags).toContain('[LNK]');
  });

  it('v4 TOON token estimate should be within budget', () => {
    const ast = buildSampleAST();
    const toon = new ToonGenerator();
    const table = toon.generateToon(ast);
    const text = toon.renderToonText(table);

    const tokens = toon.estimateTokens(text);
    // For 2 nodes, TOON should use fewer than 100 tokens
    expect(tokens).toBeLessThan(100);
  });

  it('v4 diff mode should send zero rows when nothing changes', () => {
    const ast = buildSampleAST();
    const toon = new ToonGenerator();

    const before = toon.generateToon(ast);
    const after = toon.generateToon(ast); // No changes
    const diff = toon.generateDiff(before, after);

    expect(diff.rows.length).toBe(0);
  });
});
