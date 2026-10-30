import { describe, it, expect } from 'vitest';
import { DocxParser, createBlankDocxBuffer } from '../src/parser';
import { OperationEngine } from '../src/operation-engine';
import { executeCoverPage } from '../src/operations/coverPage';
import { validateOperationStrict, validateOperationsStrict } from '../src/operation-schema';
import type { Node } from '../src/ast';

async function createFreshDoc(title: string = 'Cover Spacing Test') {
  const parser = new DocxParser(createBlankDocxBuffer(title));
  const ast = await parser.getAST();
  return { parser, ast };
}

function findCoverNodeBySlot(nodes: Node[], slot: string): Node | undefined {
  return nodes.find((n) => n.id.startsWith(`cover_${slot}_`));
}

function getSpacingElement(node: Node): Element | null {
  const pPr = node.xml.getElementsByTagName('w:pPr')[0];
  if (!pPr) return null;
  return pPr.getElementsByTagName('w:spacing')[0] ?? null;
}

describe('Cover Page Geometry and Spacing Parameterization', () => {
  describe('Dead code fix: top paragraph properties attachment', () => {
    it('attaches w:pPr and w:spacing to the top spacing paragraph', async () => {
      const { parser, ast } = await createFreshDoc();
      executeCoverPage(
        {
          action: 'coverPage',
          title: 'Document Title',
          subtitle: 'Document Subtitle',
        },
        ast,
        parser
      );

      const topNode = findCoverNodeBySlot(ast.sections[0].nodes, 'top');
      expect(topNode).toBeDefined();
      expect(topNode!.xml.tagName).toBe('w:p');

      const pPrChildren = topNode!.xml.getElementsByTagName('w:pPr');
      expect(pPrChildren.length).toBe(1);

      const pPr = pPrChildren[0];
      expect(pPr.parentNode).toBe(topNode!.xml);

      const spacingElements = pPr.getElementsByTagName('w:spacing');
      expect(spacingElements.length).toBe(1);

      const spacing = spacingElements[0];
      expect(spacing.parentNode).toBe(pPr);
      expect(spacing.getAttribute('w:before')).toBe('1440');
    });
  });

  describe('Calibrated defaults across templates', () => {
    it('applies calibrated defaults for modern_minimalist template', async () => {
      const { parser, ast } = await createFreshDoc();
      executeCoverPage(
        {
          action: 'coverPage',
          template: 'modern_minimalist',
          title: 'Minimalist Report',
        },
        ast,
        parser
      );

      const topNode = findCoverNodeBySlot(ast.sections[0].nodes, 'top')!;
      const divNode = findCoverNodeBySlot(ast.sections[0].nodes, 'div')!;
      const metaNode = findCoverNodeBySlot(ast.sections[0].nodes, 'meta')!;

      const topSpacing = getSpacingElement(topNode);
      const divSpacing = getSpacingElement(divNode);
      const metaSpacing = getSpacingElement(metaNode);

      expect(topSpacing?.getAttribute('w:before')).toBe('1440');
      expect(divSpacing?.getAttribute('w:after')).toBe('720');
      expect(metaSpacing?.getAttribute('w:after')).toBe('720');
    });

    it('applies calibrated defaults for corporate_navy template', async () => {
      const { parser, ast } = await createFreshDoc();
      executeCoverPage(
        {
          action: 'coverPage',
          template: 'corporate_navy',
          title: 'Corporate Brief',
        },
        ast,
        parser
      );

      const topNode = findCoverNodeBySlot(ast.sections[0].nodes, 'top')!;
      const divNode = findCoverNodeBySlot(ast.sections[0].nodes, 'div')!;
      const metaNode = findCoverNodeBySlot(ast.sections[0].nodes, 'meta')!;

      const topSpacing = getSpacingElement(topNode);
      const divSpacing = getSpacingElement(divNode);
      const metaSpacing = getSpacingElement(metaNode);

      expect(topSpacing?.getAttribute('w:before')).toBe('1440');
      expect(divSpacing?.getAttribute('w:after')).toBe('720');
      expect(metaSpacing?.getAttribute('w:after')).toBe('720');
    });

    it('applies calibrated defaults for academic template', async () => {
      const { parser, ast } = await createFreshDoc();
      executeCoverPage(
        {
          action: 'coverPage',
          template: 'academic',
          title: 'Research Paper',
        },
        ast,
        parser
      );

      const topNode = findCoverNodeBySlot(ast.sections[0].nodes, 'top')!;
      const divNode = findCoverNodeBySlot(ast.sections[0].nodes, 'div')!;
      const metaNode = findCoverNodeBySlot(ast.sections[0].nodes, 'meta')!;

      const topSpacing = getSpacingElement(topNode);
      const divSpacing = getSpacingElement(divNode);
      const metaSpacing = getSpacingElement(metaNode);

      expect(topSpacing?.getAttribute('w:before')).toBe('1440');
      expect(divSpacing?.getAttribute('w:after')).toBe('720');
      expect(metaSpacing?.getAttribute('w:after')).toBe('720');
    });

    it('applies calibrated default of 360 DXA metaSpacingAfter for executive template', async () => {
      const { parser, ast } = await createFreshDoc();
      executeCoverPage(
        {
          action: 'coverPage',
          template: 'executive',
          title: 'Executive Summary',
        },
        ast,
        parser
      );

      const topNode = findCoverNodeBySlot(ast.sections[0].nodes, 'top')!;
      const divNode = findCoverNodeBySlot(ast.sections[0].nodes, 'div')!;
      const metaNode = findCoverNodeBySlot(ast.sections[0].nodes, 'meta')!;

      const topSpacing = getSpacingElement(topNode);
      const divSpacing = getSpacingElement(divNode);
      const metaSpacing = getSpacingElement(metaNode);

      expect(topSpacing?.getAttribute('w:before')).toBe('1440');
      expect(divSpacing?.getAttribute('w:after')).toBe('720');
      expect(metaSpacing?.getAttribute('w:after')).toBe('360');
    });
  });

  describe('Custom parameter overrides', () => {
    it('respects custom spacing parameters passed flat', async () => {
      const { parser, ast } = await createFreshDoc();
      executeCoverPage(
        {
          action: 'coverPage',
          title: 'Custom Spacing Document',
          topSpacingBefore: 2880,
          dividerSpacingAfter: 1440,
          metaSpacingAfter: 1080,
        },
        ast,
        parser
      );

      const topNode = findCoverNodeBySlot(ast.sections[0].nodes, 'top')!;
      const divNode = findCoverNodeBySlot(ast.sections[0].nodes, 'div')!;
      const metaNode = findCoverNodeBySlot(ast.sections[0].nodes, 'meta')!;

      expect(getSpacingElement(topNode)?.getAttribute('w:before')).toBe('2880');
      expect(getSpacingElement(divNode)?.getAttribute('w:after')).toBe('1440');
      expect(getSpacingElement(metaNode)?.getAttribute('w:after')).toBe('1080');
    });

    it('overrides executive default meta spacing when an explicit value is passed', async () => {
      const { parser, ast } = await createFreshDoc();
      executeCoverPage(
        {
          action: 'coverPage',
          template: 'executive',
          title: 'Executive Document Overridden',
          metaSpacingAfter: 900,
        },
        ast,
        parser
      );

      const metaNode = findCoverNodeBySlot(ast.sections[0].nodes, 'meta')!;
      expect(getSpacingElement(metaNode)?.getAttribute('w:after')).toBe('900');
    });
  });

  describe('Explicit 0-value preservation', () => {
    it('preserves flat 0 values and does not fall back to defaults', async () => {
      const { parser, ast } = await createFreshDoc();
      executeCoverPage(
        {
          action: 'coverPage',
          title: 'Zero Spacing Test',
          topSpacingBefore: 0,
          dividerSpacingAfter: 0,
          metaSpacingAfter: 0,
        },
        ast,
        parser
      );

      const topNode = findCoverNodeBySlot(ast.sections[0].nodes, 'top')!;
      const divNode = findCoverNodeBySlot(ast.sections[0].nodes, 'div')!;
      const metaNode = findCoverNodeBySlot(ast.sections[0].nodes, 'meta')!;

      expect(getSpacingElement(topNode)?.getAttribute('w:before')).toBe('0');
      expect(getSpacingElement(divNode)?.getAttribute('w:after')).toBe('0');
      expect(getSpacingElement(metaNode)?.getAttribute('w:after')).toBe('0');
    });

    it('preserves nested 0 values in params', async () => {
      const { parser, ast } = await createFreshDoc();
      executeCoverPage(
        {
          action: 'coverPage',
          title: 'Nested Zero Spacing Test',
          params: {
            topSpacingBefore: 0,
            dividerSpacingAfter: 0,
            metaSpacingAfter: 0,
          },
        },
        ast,
        parser
      );

      const topNode = findCoverNodeBySlot(ast.sections[0].nodes, 'top')!;
      const divNode = findCoverNodeBySlot(ast.sections[0].nodes, 'div')!;
      const metaNode = findCoverNodeBySlot(ast.sections[0].nodes, 'meta')!;

      expect(getSpacingElement(topNode)?.getAttribute('w:before')).toBe('0');
      expect(getSpacingElement(divNode)?.getAttribute('w:after')).toBe('0');
      expect(getSpacingElement(metaNode)?.getAttribute('w:after')).toBe('0');
    });
  });

  describe('Nested params support and precedence', () => {
    it('resolves spacing parameters from nested op.params', async () => {
      const { parser, ast } = await createFreshDoc();
      executeCoverPage(
        {
          action: 'coverPage',
          title: 'Nested Params Test',
          params: {
            topSpacingBefore: 2000,
            dividerSpacingAfter: 1000,
            metaSpacingAfter: 500,
          },
        },
        ast,
        parser
      );

      const topNode = findCoverNodeBySlot(ast.sections[0].nodes, 'top')!;
      const divNode = findCoverNodeBySlot(ast.sections[0].nodes, 'div')!;
      const metaNode = findCoverNodeBySlot(ast.sections[0].nodes, 'meta')!;

      expect(getSpacingElement(topNode)?.getAttribute('w:before')).toBe('2000');
      expect(getSpacingElement(divNode)?.getAttribute('w:after')).toBe('1000');
      expect(getSpacingElement(metaNode)?.getAttribute('w:after')).toBe('500');
    });

    it('prioritizes nested params over flat properties', async () => {
      const { parser, ast } = await createFreshDoc();
      executeCoverPage(
        {
          action: 'coverPage',
          title: 'Precedence Test',
          topSpacingBefore: 1200,
          dividerSpacingAfter: 600,
          metaSpacingAfter: 300,
          params: {
            topSpacingBefore: 2400,
            dividerSpacingAfter: 1200,
            metaSpacingAfter: 600,
          },
        },
        ast,
        parser
      );

      const topNode = findCoverNodeBySlot(ast.sections[0].nodes, 'top')!;
      const divNode = findCoverNodeBySlot(ast.sections[0].nodes, 'div')!;
      const metaNode = findCoverNodeBySlot(ast.sections[0].nodes, 'meta')!;

      expect(getSpacingElement(topNode)?.getAttribute('w:before')).toBe('2400');
      expect(getSpacingElement(divNode)?.getAttribute('w:after')).toBe('1200');
      expect(getSpacingElement(metaNode)?.getAttribute('w:after')).toBe('600');
    });

    it('rejects negative spacing values inside nested params with typed E_VALIDATION', async () => {
      const { parser, ast } = await createFreshDoc();
      const testCases = [
        { topSpacingBefore: -1 },
        { dividerSpacingAfter: -50 },
        { metaSpacingAfter: -10 },
      ];

      for (const params of testCases) {
        try {
          executeCoverPage(
            {
              action: 'coverPage',
              title: 'Negative Nested Spacing',
              params,
            },
            ast,
            parser
          );
          expect.unreachable('Should have thrown for negative nested spacing');
        } catch (err: unknown) {
          const error = err as { code?: string; message?: string; remedy?: string };
          expect(error.code).toBe('E_VALIDATION');
          expect(error.message).toContain('E_VALIDATION');
          expect(error.remedy).toBeDefined();
        }
      }
    });

    it('rejects fractional spacing values inside nested params with typed E_VALIDATION', async () => {
      const { parser, ast } = await createFreshDoc();
      const testCases = [
        { topSpacingBefore: 1440.5 },
        { dividerSpacingAfter: 720.25 },
        { metaSpacingAfter: 360.75 },
      ];

      for (const params of testCases) {
        try {
          executeCoverPage(
            {
              action: 'coverPage',
              title: 'Fractional Nested Spacing',
              params,
            },
            ast,
            parser
          );
          expect.unreachable('Should have thrown for fractional nested spacing');
        } catch (err: unknown) {
          const error = err as { code?: string; message?: string; remedy?: string };
          expect(error.code).toBe('E_VALIDATION');
          expect(error.message).toContain('E_VALIDATION');
          expect(error.remedy).toBeDefined();
        }
      }
    });

    it('rejects invalid nested params spacing through OperationEngine with typed E_VALIDATION', async () => {
      const { parser, ast } = await createFreshDoc();
      const engine = new OperationEngine(ast, parser);

      await expect(
        engine.execute({
          action: 'coverPage',
          title: 'Engine Spacing Validation',
          params: { topSpacingBefore: -1 },
        })
      ).rejects.toMatchObject({ code: 'E_VALIDATION' });

      await expect(
        engine.execute({
          action: 'coverPage',
          title: 'Engine Spacing Validation',
          params: { dividerSpacingAfter: 1.5 },
        })
      ).rejects.toMatchObject({ code: 'E_VALIDATION' });
    });
  });

  describe('Execution via OperationEngine', () => {
    it('executes coverPage through OperationEngine batch workflow', async () => {
      const { parser, ast } = await createFreshDoc();
      const engine = new OperationEngine(ast, parser);

      await engine.execute({
        action: 'coverPage',
        title: 'Engine Batch Title',
        subtitle: 'Engine Subtitle',
        topSpacingBefore: 1800,
        dividerSpacingAfter: 900,
        metaSpacingAfter: 450,
      });

      const topNode = findCoverNodeBySlot(ast.sections[0].nodes, 'top')!;
      const divNode = findCoverNodeBySlot(ast.sections[0].nodes, 'div')!;
      const metaNode = findCoverNodeBySlot(ast.sections[0].nodes, 'meta')!;

      expect(getSpacingElement(topNode)?.getAttribute('w:before')).toBe('1800');
      expect(getSpacingElement(divNode)?.getAttribute('w:after')).toBe('900');
      expect(getSpacingElement(metaNode)?.getAttribute('w:after')).toBe('450');
    });
  });

  describe('Schema validation', () => {
    it('validates valid coverPage operations with spacing fields', () => {
      const validOp = {
        action: 'coverPage',
        title: 'Valid Title',
        template: 'corporate_navy',
        topSpacingBefore: 1440,
        dividerSpacingAfter: 720,
        metaSpacingAfter: 720,
      };

      const result = validateOperationStrict(validOp);
      expect(result.ok).toBe(true);

      const batchResult = validateOperationsStrict([validOp]);
      expect(batchResult.ok).toBe(true);
    });

    it('accepts zero values in schema validation', () => {
      const zeroOp = {
        action: 'coverPage',
        title: 'Zero Values Operation',
        topSpacingBefore: 0,
        dividerSpacingAfter: 0,
        metaSpacingAfter: 0,
      };

      const result = validateOperationStrict(zeroOp);
      expect(result.ok).toBe(true);
    });

    it('rejects negative spacing values at the schema gate', () => {
      expect(() =>
        validateOperationStrict({
          action: 'coverPage',
          title: 'Negative Spacing',
          topSpacingBefore: -100,
        })
      ).toThrow(/E_VALIDATION/);

      expect(() =>
        validateOperationStrict({
          action: 'coverPage',
          title: 'Negative Spacing',
          dividerSpacingAfter: -50,
        })
      ).toThrow(/E_VALIDATION/);

      expect(() =>
        validateOperationStrict({
          action: 'coverPage',
          title: 'Negative Spacing',
          metaSpacingAfter: -1,
        })
      ).toThrow(/E_VALIDATION/);
    });

    it('rejects non-integer spacing values at the schema gate', () => {
      expect(() =>
        validateOperationStrict({
          action: 'coverPage',
          title: 'Float Spacing',
          topSpacingBefore: 1440.5,
        })
      ).toThrow(/E_VALIDATION/);
    });

    it('rejects negative spacing values inside nested params at schema gate', () => {
      expect(() =>
        validateOperationStrict({
          action: 'coverPage',
          title: 'Negative Spacing in params',
          params: { topSpacingBefore: -100 },
        })
      ).toThrow(/E_VALIDATION/);

      expect(() =>
        validateOperationStrict({
          action: 'coverPage',
          title: 'Negative Spacing in params',
          params: { dividerSpacingAfter: -50 },
        })
      ).toThrow(/E_VALIDATION/);

      expect(() =>
        validateOperationStrict({
          action: 'coverPage',
          title: 'Negative Spacing in params',
          params: { metaSpacingAfter: -1 },
        })
      ).toThrow(/E_VALIDATION/);
    });

    it('rejects non-integer spacing values inside nested params at schema gate', () => {
      expect(() =>
        validateOperationStrict({
          action: 'coverPage',
          title: 'Float Spacing in params',
          params: { topSpacingBefore: 1440.5 },
        })
      ).toThrow(/E_VALIDATION/);

      expect(() =>
        validateOperationStrict({
          action: 'coverPage',
          title: 'Float Spacing in params',
          params: { dividerSpacingAfter: 720.5 },
        })
      ).toThrow(/E_VALIDATION/);

      expect(() =>
        validateOperationStrict({
          action: 'coverPage',
          title: 'Float Spacing in params',
          params: { metaSpacingAfter: 360.25 },
        })
      ).toThrow(/E_VALIDATION/);
    });
  });

  describe('OOXML validity and save/re-open round-trip', () => {
    it('successfully serializes and re-opens a document with cover page spacing', async () => {
      const { parser, ast } = await createFreshDoc('Round Trip Document');
      executeCoverPage(
        {
          action: 'coverPage',
          template: 'modern_minimalist',
          title: 'Round Trip Title',
          subtitle: 'A Verification Document',
          author: 'Test Author',
          organization: 'Test Org',
          date: '2026-10-23',
          topSpacingBefore: 1800,
          dividerSpacingAfter: 900,
          metaSpacingAfter: 540,
        },
        ast,
        parser
      );

      // Re-zip verifies pre-save integrity, canonicalizes DOM according to ECMA-376,
      // and verifies post-assembly integrity
      const docxBuffer = parser.rezip();
      expect(docxBuffer).toBeInstanceOf(Buffer);
      expect(docxBuffer.length).toBeGreaterThan(0);

      // Re-load saved buffer to verify parseability and OOXML schema conformance
      const reloadedParser = new DocxParser(docxBuffer);
      const reloadedAst = await reloadedParser.getAST();

      expect(reloadedAst.sections.length).toBeGreaterThanOrEqual(1);
      const reloadedNodes = reloadedAst.sections[0].nodes;
      expect(reloadedNodes.length).toBeGreaterThanOrEqual(5);

      // Check top paragraph in the reloaded DOM
      const firstNode = reloadedNodes[0];
      expect(firstNode.type).toBe('paragraph');
      const topPPr = firstNode.xml.getElementsByTagName('w:pPr')[0];
      expect(topPPr).toBeDefined();
      const topSpacing = topPPr.getElementsByTagName('w:spacing')[0];
      expect(topSpacing).toBeDefined();
      expect(topSpacing.getAttribute('w:before')).toBe('1800');

      // Verify divider line paragraph in reloaded DOM
      // Divider has w:pBdr
      const divParagraphs = reloadedNodes.filter((n) => {
        const pPr = n.xml.getElementsByTagName('w:pPr')[0];
        return pPr && pPr.getElementsByTagName('w:pBdr').length > 0;
      });
      expect(divParagraphs.length).toBe(1);
      const divSpacing = divParagraphs[0].xml.getElementsByTagName('w:spacing')[0];
      expect(divSpacing).toBeDefined();
      expect(divSpacing.getAttribute('w:after')).toBe('900');
    });

    it('round-trips zero-spacing cover page without corruption', async () => {
      const { parser, ast } = await createFreshDoc('Zero Spacing Round Trip');
      executeCoverPage(
        {
          action: 'coverPage',
          title: 'Zero Spacing In Rezip',
          topSpacingBefore: 0,
          dividerSpacingAfter: 0,
          metaSpacingAfter: 0,
        },
        ast,
        parser
      );

      const buffer = parser.rezip();
      const reloadedParser = new DocxParser(buffer);
      const reloadedAst = await reloadedParser.getAST();

      const topNode = reloadedAst.sections[0].nodes[0];
      const topSpacing = topNode.xml.getElementsByTagName('w:spacing')[0];
      expect(topSpacing).toBeDefined();
      expect(topSpacing.getAttribute('w:before')).toBe('0');
    });

    it('round-trips executive template with default 360 meta spacing', async () => {
      const { parser, ast } = await createFreshDoc('Executive Round Trip');
      executeCoverPage(
        {
          action: 'coverPage',
          template: 'executive',
          title: 'Executive Specimen',
        },
        ast,
        parser
      );

      const buffer = parser.rezip();
      const reloadedParser = new DocxParser(buffer);
      const reloadedAst = await reloadedParser.getAST();

      const topNode = reloadedAst.sections[0].nodes[0];
      const topSpacing = topNode.xml.getElementsByTagName('w:spacing')[0];
      expect(topSpacing.getAttribute('w:before')).toBe('1440');
    });
  });
});
