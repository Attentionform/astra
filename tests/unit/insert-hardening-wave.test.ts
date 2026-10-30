import { describe, it, expect } from 'vitest';
import {
  DocxParser,
  createBlankDocxBuffer,
  executeInsert,
  executeDelete,
  executeTable,
  isValidParaId,
  OperationEngine,
} from '@attentionform/astra-engine';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

async function freshDoc(): Promise<{ parser: DocxParser; ast: any; anchor: string; trailingNodeId: string }> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id as string;
  const trailingNodeId = ast.sections[0].nodes[1]?.id as string;
  return { parser, ast, anchor, trailingNodeId };
}

function getNodeSequence(ast: any): string[] {
  return ast.sections.flatMap((s: any) => s.nodes.map((n: any) => n.id));
}

function getDomSiblingSequence(startNodeXml: Element): string[] {
  const ids: string[] = [];
  let curr: Element | null = startNodeXml;
  while (curr) {
    if (curr.nodeType === 1 && curr.tagName !== 'w:sectPr') {
      const id = curr.getAttribute('data-astra-id');
      if (id) ids.push(id);
    }
    curr = curr.nextSibling as Element | null;
  }
  return ids;
}

describe('Wave 1E: Insert Hardening (P0-9, P0-12, P0-13, T-1)', () => {

  describe('P0-9: Same-Anchor "after" Inserts FIFO Ordering', () => {
    it('preserves FIFO ordering in both AST and DOM for consecutive direct inserts', async () => {
      const { ast, anchor, trailingNodeId } = await freshDoc();
      const anchorNode = ast.sections[0].nodes[0];

      await executeInsert({
        action: 'insert',
        type: 'paragraph',
        anchor,
        position: 'after',
        text: 'Item A',
        id: 'node-A'
      }, ast);

      await executeInsert({
        action: 'insert',
        type: 'paragraph',
        anchor,
        position: 'after',
        text: 'Item B',
        id: 'node-B'
      }, ast);

      await executeInsert({
        action: 'insert',
        type: 'paragraph',
        anchor,
        position: 'after',
        text: 'Item C',
        id: 'node-C'
      }, ast);

      // Verify AST order: anchor -> node-A -> node-B -> node-C -> trailingNodeId
      const astIds = getNodeSequence(ast);
      expect(astIds).toEqual([anchor, 'node-A', 'node-B', 'node-C', trailingNodeId]);

      // Verify DOM sibling order: anchor -> node-A -> node-B -> node-C -> trailingNodeId
      const domIds = getDomSiblingSequence(anchorNode.xml);
      expect(domIds).toEqual([anchor, 'node-A', 'node-B', 'node-C', trailingNodeId]);
    });

    it('preserves FIFO ordering with existing subsequent siblings', async () => {
      const { ast, anchor, trailingNodeId } = await freshDoc();
      const anchorNode = ast.sections[0].nodes[0];

      // Insert multiple nodes consecutively after anchor
      await executeInsert({
        action: 'insert',
        type: 'paragraph',
        anchor,
        position: 'after',
        text: 'Alpha',
        id: 'node-alpha'
      }, ast);

      await executeInsert({
        action: 'insert',
        type: 'paragraph',
        anchor,
        position: 'after',
        text: 'Beta',
        id: 'node-beta'
      }, ast);

      await executeInsert({
        action: 'insert',
        type: 'paragraph',
        anchor,
        position: 'after',
        text: 'Gamma',
        id: 'node-gamma'
      }, ast);

      // Expected order: anchor -> alpha -> beta -> gamma -> trailingNodeId
      const astIds = getNodeSequence(ast);
      expect(astIds).toEqual([anchor, 'node-alpha', 'node-beta', 'node-gamma', trailingNodeId]);

      const domIds = getDomSiblingSequence(anchorNode.xml);
      expect(domIds).toEqual([anchor, 'node-alpha', 'node-beta', 'node-gamma', trailingNodeId]);
    });

    it('preserves FIFO ordering when executed through OperationEngine batch', async () => {
      const { parser, ast, anchor, trailingNodeId } = await freshDoc();
      const engine = new OperationEngine(ast, parser);

      const result = await engine.executeBatch({
        operations: [
          { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Op 1', id: 'batch-1' },
          { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Op 2', id: 'batch-2' },
          { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Op 3', id: 'batch-3' },
        ]
      });

      expect(result.success).toBe(true);
      expect(result.appliedCount).toBe(3);

      const astIds = getNodeSequence(ast);
      expect(astIds).toEqual([anchor, 'batch-1', 'batch-2', 'batch-3', trailingNodeId]);

      const anchorXml = ast.sections[0].nodes[0].xml;
      const domIds = getDomSiblingSequence(anchorXml);
      expect(domIds).toEqual([anchor, 'batch-1', 'batch-2', 'batch-3', trailingNodeId]);
    });
  });

  describe('T-1: Multi-Root XML Fragment Handling', () => {
    it('imports all root elements in sequence with zero silent drops', async () => {
      const { ast, anchor, trailingNodeId } = await freshDoc();
      const anchorNode = ast.sections[0].nodes[0];

      const multiRootXml =
        `<w:p><w:r><w:t>First Paragraph</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>Second Paragraph</w:t></w:r></w:p>` +
        `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell 1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;

      await executeInsert({
        action: 'insert',
        type: 'xml',
        anchor,
        position: 'after',
        xml: multiRootXml,
        id: 'multi-root'
      }, ast);

      // Verify AST contains all 3 roots plus initial anchor and trailing node
      const nodes = ast.sections[0].nodes;
      expect(nodes.length).toBe(5);
      expect(nodes[0].id).toBe(anchor);
      expect(nodes[1].id).toBe('multi-root');
      expect(nodes[1].type).toBe('paragraph');
      expect(nodes[1].metadata.text).toBe('First Paragraph');
      expect(nodes[2].id).toBe('multi-root_1');
      expect(nodes[2].type).toBe('paragraph');
      expect(nodes[2].metadata.text).toBe('Second Paragraph');
      expect(nodes[3].id).toBe('multi-root_2');
      expect(nodes[3].type).toBe('table');
      expect(nodes[4].id).toBe(trailingNodeId);

      // Verify DOM sequence
      const domIds = getDomSiblingSequence(anchorNode.xml);
      expect(domIds).toEqual([anchor, 'multi-root', 'multi-root_1', 'multi-root_2', trailingNodeId]);
    });

    it('interops with FIFO same-anchor inserts after multi-root fragment', async () => {
      const { ast, anchor, trailingNodeId } = await freshDoc();
      const anchorNode = ast.sections[0].nodes[0];

      const multiRootXml =
        `<w:p><w:r><w:t>Root A</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>Root B</w:t></w:r></w:p>`;

      await executeInsert({
        action: 'insert',
        type: 'xml',
        anchor,
        position: 'after',
        xml: multiRootXml,
        id: 'frag'
      }, ast);

      await executeInsert({
        action: 'insert',
        type: 'paragraph',
        anchor,
        position: 'after',
        text: 'Follow-up C',
        id: 'follow-up'
      }, ast);

      // Expected order: anchor -> frag (Root A) -> frag_1 (Root B) -> follow-up (Follow-up C) -> trailingNodeId
      const astIds = getNodeSequence(ast);
      expect(astIds).toEqual([anchor, 'frag', 'frag_1', 'follow-up', trailingNodeId]);

      const domIds = getDomSiblingSequence(anchorNode.xml);
      expect(domIds).toEqual([anchor, 'frag', 'frag_1', 'follow-up', trailingNodeId]);
    });

    it('stamps valid w14:paraId on every paragraph across all roots', async () => {
      const { ast, anchor } = await freshDoc();

      const multiRootXml =
        `<w:p><w:r><w:t>P1</w:t></w:r></w:p>` +
        `<w:tbl>` +
          `<w:tr><w:tc><w:p><w:r><w:t>C1</w:t></w:r></w:p></w:tc></w:tr>` +
          `<w:tr><w:tc><w:p><w:r><w:t>C2</w:t></w:r></w:p></w:tc></w:tr>` +
        `</w:tbl>` +
        `<w:p><w:r><w:t>P2</w:t></w:r></w:p>`;

      await executeInsert({
        action: 'insert',
        type: 'xml',
        anchor,
        position: 'after',
        xml: multiRootXml,
        id: 'multi-stamp'
      }, ast);

      const allParas = (ast.sections[0].nodes[0].xml.ownerDocument as Document).getElementsByTagName('w:p');
      expect(allParas.length).toBe(6); // 2 original + 4 inserted

      const seenIds = new Set<string>();
      for (let i = 0; i < allParas.length; i++) {
        const pid = (allParas[i] as unknown as Element).getAttribute('w14:paraId');
        expect(pid).toBeDefined();
        expect(isValidParaId(pid)).toBe(true);
        expect(seenIds.has(pid!.toUpperCase())).toBe(false);
        seenIds.add(pid!.toUpperCase());
      }
      expect(seenIds.size).toBe(6);
    });
  });

  describe('P0-13: Non-OOXML Tag Defense', () => {
    it('rejects <script> tag with E_VALIDATION and actionable remedy', async () => {
      const { ast, anchor } = await freshDoc();
      const initialNodeCount = ast.sections[0].nodes.length;

      let err: any;
      try {
        await executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<script>alert("xss")</script>'
        }, ast);
      } catch (e: any) {
        err = e;
      }

      expect(err).toBeDefined();
      expect(err.code).toBe('E_VALIDATION');
      expect(err.retryable).toBe(false);
      expect(err.remedy).toBeDefined();
      expect(err.remedy.length).toBeGreaterThan(10);
      expect(err.message).toContain('E_VALIDATION');
      expect(err.message).toContain('script');
      // Assert fail-closed: no nodes added
      expect(ast.sections[0].nodes.length).toBe(initialNodeCount);
    });

    it('rejects <iframe> tag with E_VALIDATION', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<iframe src="https://evil.com"></iframe>'
        }, ast)
      ).rejects.toMatchObject({
        code: 'E_VALIDATION',
        retryable: false
      });
    });

    it('rejects nested dangerous tags inside valid OOXML wrapper', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p><w:r><script>alert(1)</script></w:r></w:p>'
        }, ast)
      ).rejects.toMatchObject({
        code: 'E_VALIDATION',
        retryable: false
      });
    });

    it('rejects generic HTML markup (div, span, p)', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<div><span>HTML block</span></div>'
        }, ast)
      ).rejects.toMatchObject({
        code: 'E_VALIDATION',
        retryable: false
      });
    });

    it('rejects unknown external XML namespaces', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<custom:data xmlns:custom="http://evil.com/ns">payload</custom:data>'
        }, ast)
      ).rejects.toMatchObject({
        code: 'E_VALIDATION',
        retryable: false
      });
    });

    it('rejects fake tags with OOXML prefix like w:script', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:script>alert(1)</w:script>'
        }, ast)
      ).rejects.toMatchObject({
        code: 'E_VALIDATION',
        retryable: false
      });
    });

    it('allows valid OOXML namespaces (w:*, m:*, wp:*, a:*, pic:*, w14:*)', async () => {
      const { ast, anchor } = await freshDoc();
      const validMath =
        `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">` +
        `<m:r><m:t>x=y+1</m:t></m:r>` +
        `</m:oMath>`;

      await executeInsert({
        action: 'insert',
        type: 'xml',
        anchor,
        position: 'after',
        xml: validMath,
        id: 'math-node'
      }, ast);

      const mathNode = ast.sections[0].nodes.find((n: any) => n.id === 'math-node');
      expect(mathNode).toBeDefined();
    });
  });

  describe('P0-12: Reject Malformed / Truncated XML with E_VALIDATION', () => {
    it('rejects truncated XML with unclosed tags', async () => {
      const { ast, anchor } = await freshDoc();
      let err: any;
      try {
        await executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p><w:r><w:t>truncated text'
        }, ast);
      } catch (e: any) {
        err = e;
      }

      expect(err).toBeDefined();
      expect(err.code).toBe('E_VALIDATION');
      expect(err.retryable).toBe(false);
      expect(err.remedy).toBeDefined();
      expect(err.message).toContain('E_VALIDATION');
    });

    it('rejects mismatched closing tags', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p><w:r></w:p>'
        }, ast)
      ).rejects.toMatchObject({
        code: 'E_VALIDATION',
        retryable: false
      });
    });

    it('rejects invalid XML syntax', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p <<<'
        }, ast)
      ).rejects.toMatchObject({
        code: 'E_VALIDATION',
        retryable: false
      });
    });

    it('rejects empty or whitespace-only XML string', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '   \n  \t  '
        }, ast)
      ).rejects.toMatchObject({
        code: 'E_VALIDATION',
        retryable: false
      });
    });

    it('rejects dangling text outside OOXML root elements', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p><w:r><w:t>Valid</w:t></w:r></w:p> dangling text outside'
        }, ast)
      ).rejects.toMatchObject({
        code: 'E_VALIDATION',
        retryable: false
      });
    });

    it('rejects XXE and DTD declarations with E_VALIDATION', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<!DOCTYPE test [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><w:p><w:r><w:t>&xxe;</w:t></w:r></w:p>'
        }, ast)
      ).rejects.toMatchObject({
        code: 'E_VALIDATION',
        retryable: false
      });
    });
  });

  describe('ParaId Identity Stamps on Insert', () => {
    it('preserves caller-supplied valid w14:paraId and mints valid on invalid', async () => {
      const { ast, anchor } = await freshDoc();

      const xml =
        `<w:p xmlns:w14="${W14}" w14:paraId="1A2B3C4D"><w:r><w:t>Valid Id</w:t></w:r></w:p>` +
        `<w:p xmlns:w14="${W14}" w14:paraId="invalid_id"><w:r><w:t>Invalid Id</w:t></w:r></w:p>`;

      await executeInsert({
        action: 'insert',
        type: 'xml',
        anchor,
        position: 'after',
        xml,
        id: 'para-id-test'
      }, ast);

      const node1 = ast.sections[0].nodes[1];
      const node2 = ast.sections[0].nodes[2];

      const pid1 = (node1.xml as Element).getAttribute('w14:paraId');
      const pid2 = (node2.xml as Element).getAttribute('w14:paraId');

      expect(pid1).toBe('1A2B3C4D');
      expect(pid2).not.toBe('invalid_id');
      expect(isValidParaId(pid2)).toBe(true);
      expect(pid2).toMatch(/^[0-9A-Fa-f]{8}$/);
    });
  });

  describe('Stale-tail invalidation on delete', () => {
    it('delete tail then same-anchor after lands after latest live link', async () => {
      const { parser, ast, anchor, trailingNodeId } = await freshDoc();
      const anchorNode = ast.sections[0].nodes[0];

      await executeInsert({ action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'A', id: 't-A' }, ast, parser);
      await executeInsert({ action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'B', id: 't-B' }, ast, parser);
      expect(getNodeSequence(ast)).toEqual([anchor, 't-A', 't-B', trailingNodeId]);

      executeDelete({ action: 'delete', target: 't-B' }, ast, parser);
      expect(getNodeSequence(ast)).toEqual([anchor, 't-A', trailingNodeId]);

      await executeInsert({ action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'C', id: 't-C' }, ast, parser);

      expect(getNodeSequence(ast)).toEqual([anchor, 't-A', 't-C', trailingNodeId]);
      expect(getDomSiblingSequence(anchorNode.xml)).toEqual([anchor, 't-A', 't-C', trailingNodeId]);
    });

    it('delete middle link keeps FIFO tail attached to the live chain', async () => {
      const { parser, ast, anchor, trailingNodeId } = await freshDoc();
      const anchorNode = ast.sections[0].nodes[0];

      await executeInsert({ action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'A', id: 'm-A' }, ast, parser);
      await executeInsert({ action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'B', id: 'm-B' }, ast, parser);
      await executeInsert({ action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'C', id: 'm-C' }, ast, parser);

      executeDelete({ action: 'delete', target: 'm-B' }, ast, parser);

      await executeInsert({ action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'D', id: 'm-D' }, ast, parser);

      expect(getNodeSequence(ast)).toEqual([anchor, 'm-A', 'm-C', 'm-D', trailingNodeId]);
      expect(getDomSiblingSequence(anchorNode.xml)).toEqual([anchor, 'm-A', 'm-C', 'm-D', trailingNodeId]);
    });
  });

  describe('Strict well-formedness and central XXE assert', () => {
    it('rejects a missing final angle bracket instead of repairing it', async () => {
      const { ast, anchor } = await freshDoc();
      const before = ast.sections[0].nodes.length;
      let err: any;
      try {
        await executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p><w:r><w:t>hi</w:t></w:r></w:p'
        }, ast);
      } catch (e: any) {
        err = e;
      }
      expect(err).toBeDefined();
      expect(err.code).toBe('E_VALIDATION');
      expect(err.retryable).toBe(false);
      expect(ast.sections[0].nodes.length).toBe(before);
    });

    it('rejects lowercase doctype declarations with E_VALIDATION', async () => {
      const { ast, anchor } = await freshDoc();
      const before = ast.sections[0].nodes.length;
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<!doctype evil><w:p><w:r><w:t>hi</w:t></w:r></w:p>'
        }, ast)
      ).rejects.toMatchObject({ code: 'E_VALIDATION', retryable: false });
      expect(ast.sections[0].nodes.length).toBe(before);
    });

    it('rejects a lowercase doctype smuggled after the payload', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p><w:r><w:t>hi</w:t></w:r></w:p><!doctype evil>'
        }, ast)
      ).rejects.toMatchObject({ code: 'E_VALIDATION', retryable: false });
    });

    it('rejects lowercase entity declarations with E_VALIDATION', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<!entity xxe "evil"><w:p><w:r><w:t>hi</w:t></w:r></w:p>'
        }, ast)
      ).rejects.toMatchObject({ code: 'E_VALIDATION', retryable: false });
    });
  });

  describe('Inserted-element attribute allowlist', () => {
    it('rejects onclick event-handler attributes on the root', async () => {
      const { ast, anchor } = await freshDoc();
      const before = ast.sections[0].nodes.length;
      let err: any;
      try {
        await executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p onclick="alert(1)"><w:r><w:t>hi</w:t></w:r></w:p>'
        }, ast);
      } catch (e: any) {
        err = e;
      }
      expect(err).toBeDefined();
      expect(err.code).toBe('E_VALIDATION');
      expect(err.message).toContain('onclick');
      expect(ast.sections[0].nodes.length).toBe(before);
    });

    it('rejects style attributes on the root', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p style="color:red"><w:r><w:t>hi</w:t></w:r></w:p>'
        }, ast)
      ).rejects.toMatchObject({ code: 'E_VALIDATION', retryable: false });
    });

    it('rejects nested event-handler attributes', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p><w:r><w:t onmouseover="x()">hi</w:t></w:r></w:p>'
        }, ast)
      ).rejects.toMatchObject({ code: 'E_VALIDATION', retryable: false });
    });

    it('rejects unknown non-allowlisted attributes', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p foo="bar"><w:r><w:t>hi</w:t></w:r></w:p>'
        }, ast)
      ).rejects.toMatchObject({ code: 'E_VALIDATION', retryable: false });
    });

    it('accepts engine-emitted formatting attributes without false positives', async () => {
      const { ast, anchor } = await freshDoc();
      await executeInsert({
        action: 'insert',
        type: 'xml',
        anchor,
        position: 'after',
        xml: '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve"> hi </w:t></w:r></w:p>',
        id: 'fmt-ok'
      }, ast);
      const node = ast.sections[0].nodes.find((n: any) => n.id === 'fmt-ok');
      expect(node).toBeDefined();
      expect(node.metadata.text).toBe(' hi ');
    });

    it('accepts drawing docPr bare id/name attributes', async () => {
      const { ast, anchor } = await freshDoc();
      await executeInsert({
        action: 'insert',
        type: 'xml',
        anchor,
        position: 'after',
        xml: '<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="pic"/></wp:inline></w:drawing></w:r></w:p>',
        id: 'draw-ok'
      }, ast);
      const node = ast.sections[0].nodes.find((n: any) => n.id === 'draw-ok');
      expect(node).toBeDefined();
    });

    it('rejects prefixed w:onclick on the root', async () => {
      const { ast, anchor } = await freshDoc();
      const before = ast.sections[0].nodes.length;
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p w:onclick="alert(1)"><w:r><w:t>hi</w:t></w:r></w:p>'
        }, ast)
      ).rejects.toMatchObject({ code: 'E_VALIDATION', retryable: false });
      expect(ast.sections[0].nodes.length).toBe(before);
    });

    it('rejects prefixed xml:onclick on the root', async () => {
      const { ast, anchor } = await freshDoc();
      const before = ast.sections[0].nodes.length;
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p xml:onclick="alert(1)"><w:r><w:t>hi</w:t></w:r></w:p>'
        }, ast)
      ).rejects.toMatchObject({ code: 'E_VALIDATION', retryable: false });
      expect(ast.sections[0].nodes.length).toBe(before);
    });

    it('rejects prefixed w:style on the root', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p w:style="color:red"><w:r><w:t>hi</w:t></w:r></w:p>'
        }, ast)
      ).rejects.toMatchObject({ code: 'E_VALIDATION', retryable: false });
    });

    it('rejects nested prefixed w:onmouseover on w:t', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p><w:r><w:t w:onmouseover="x()">hi</w:t></w:r></w:p>'
        }, ast)
      ).rejects.toMatchObject({ code: 'E_VALIDATION', retryable: false });
    });

    it('rejects case-variant prefixed handler W:ONCLICK', async () => {
      const { ast, anchor } = await freshDoc();
      await expect(
        executeInsert({
          action: 'insert',
          type: 'xml',
          anchor,
          position: 'after',
          xml: '<w:p W:ONCLICK="alert(1)"><w:r><w:t>hi</w:t></w:r></w:p>'
        }, ast)
      ).rejects.toMatchObject({ code: 'E_VALIDATION', retryable: false });
    });

    it('accepts legit w:val attribute', async () => {
      const { ast, anchor } = await freshDoc();
      await executeInsert({
        action: 'insert',
        type: 'xml',
        anchor,
        position: 'after',
        xml: '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>hi</w:t></w:r></w:p>',
        id: 'wval-ok'
      }, ast);
      expect(ast.sections[0].nodes.find((n: any) => n.id === 'wval-ok')).toBeDefined();
    });

    it('accepts legit xml:space attribute', async () => {
      const { ast, anchor } = await freshDoc();
      await executeInsert({
        action: 'insert',
        type: 'xml',
        anchor,
        position: 'after',
        xml: '<w:p><w:r><w:t xml:space="preserve"> hi </w:t></w:r></w:p>',
        id: 'sp-ok'
      }, ast);
      const node = ast.sections[0].nodes.find((n: any) => n.id === 'sp-ok');
      expect(node).toBeDefined();
      expect(node.metadata.text).toBe(' hi ');
    });

    it('accepts legit wp:docPr id/name drawing fragment', async () => {
      const { ast, anchor } = await freshDoc();
      await executeInsert({
        action: 'insert',
        type: 'xml',
        anchor,
        position: 'after',
        xml: '<w:p><w:r><w:drawing><wp:inline><wp:docPr id="7" name="pic7"/></wp:inline></w:drawing></w:r></w:p>',
        id: 'docpr-ok'
      }, ast);
      expect(ast.sections[0].nodes.find((n: any) => n.id === 'docpr-ok')).toBeDefined();
    });

    it('accepts legit bare data-astra-id attribute', async () => {
      const { ast, anchor } = await freshDoc();
      await executeInsert({
        action: 'insert',
        type: 'xml',
        anchor,
        position: 'after',
        xml: '<w:p data-astra-id="keep-me"><w:r><w:t>hi</w:t></w:r></w:p>',
        id: 'astra-ok'
      }, ast);
      expect(ast.sections[0].nodes.find((n: any) => n.id === 'astra-ok')).toBeDefined();
    });
  });

  describe('Caller-supplied paraId collision', () => {
    it('mints fresh ids for intra-fragment duplicates of a valid id', async () => {
      const { ast, anchor } = await freshDoc();
      const xml =
        `<w:p xmlns:w14="${W14}" w14:paraId="1A2B3C4D"><w:r><w:t>one</w:t></w:r></w:p>` +
        `<w:p xmlns:w14="${W14}" w14:paraId="1A2B3C4D"><w:r><w:t>two</w:t></w:r></w:p>`;
      await executeInsert({ action: 'insert', type: 'xml', anchor, position: 'after', xml, id: 'dup' }, ast);

      const pid1 = (ast.sections[0].nodes[1].xml as Element).getAttribute('w14:paraId');
      const pid2 = (ast.sections[0].nodes[2].xml as Element).getAttribute('w14:paraId');
      expect(pid1).toBe('1A2B3C4D');
      expect(isValidParaId(pid2)).toBe(true);
      expect(pid2).not.toBe('1A2B3C4D');

      const doc = ast.sections[0].nodes[0].xml.ownerDocument as Document;
      const all = doc.getElementsByTagName('w:p');
      const seen = new Set<string>();
      for (let i = 0; i < all.length; i++) {
        const v = (all[i] as unknown as Element).getAttribute('w14:paraId');
        expect(v).toBeTruthy();
        expect(seen.has(v!.toUpperCase())).toBe(false);
        seen.add(v!.toUpperCase());
      }
    });

    it('mints a fresh id on collision with a live-document id', async () => {
      const { ast, anchor } = await freshDoc();
      const doc = ast.sections[0].nodes[0].xml.ownerDocument as Document;
      const livePid = ((doc.getElementsByTagName('w:p')[0] as unknown as Element).getAttribute('w14:paraId') || '') as string;
      expect(isValidParaId(livePid)).toBe(true);

      const xml = `<w:p xmlns:w14="${W14}" w14:paraId="${livePid}"><w:r><w:t>collision</w:t></w:r></w:p>`;
      await executeInsert({ action: 'insert', type: 'xml', anchor, position: 'after', xml, id: 'collide' }, ast);

      const inserted = ast.sections[0].nodes.find((n: any) => n.id === 'collide');
      const minted = (inserted.xml as Element).getAttribute('w14:paraId');
      expect(isValidParaId(minted)).toBe(true);
      expect(minted).not.toBe(livePid);

      const all = doc.getElementsByTagName('w:p');
      const seen = new Set<string>();
      for (let i = 0; i < all.length; i++) {
        const v = (all[i] as unknown as Element).getAttribute('w14:paraId');
        expect(seen.has(v!.toUpperCase())).toBe(false);
        seen.add(v!.toUpperCase());
      }
    });
  });

  describe('After-table inserts land past a trailing caption sibling', () => {
    async function tableWithBelowCaption() {
      const { ast, anchor } = await freshDoc();
      await executeInsert({
        action: 'insert',
        type: 'table',
        anchor,
        position: 'after',
        headers: ['A', 'B'],
        rows: [['1', '2']],
        captionText: 'Table 1: Results',
        captionPosition: 'below',
        id: 'cap-table',
      }, ast);
      const tableNode = ast.sections[0].nodes.find((n: any) => n.id === 'cap-table');
      expect(tableNode).toBeDefined();
      return { ast, anchor, tableId: tableNode.id as string };
    }

    it('places an after-table paragraph past the below-caption in AST and DOM order', async () => {
      const { ast, tableId } = await tableWithBelowCaption();
      const tableNode = ast.sections[0].nodes.find((n: any) => n.id === tableId);
      await executeInsert({
        action: 'insert', type: 'paragraph', anchor: tableId, position: 'after',
        text: 'After table', id: 'node-after',
      }, ast);

      const astIds = getNodeSequence(ast);
      const tableIdx = astIds.indexOf(tableId);
      const captionIdx = astIds.findIndex((id) => id.startsWith('tcap_'));
      expect(captionIdx).toBe(tableIdx + 1);
      expect(astIds[tableIdx + 2]).toBe('node-after');

      const domIds = getDomSiblingSequence(tableNode.xml);
      expect(domIds.slice(0, 3)).toEqual([tableId, astIds[tableIdx + 1], 'node-after']);
    });

    it('leaves above-caption tables alone (after lands directly after the table)', async () => {
      const { ast, anchor } = await freshDoc();
      await executeInsert({
        action: 'insert',
        type: 'table',
        anchor,
        position: 'after',
        headers: ['A'],
        rows: [['1']],
        captionText: 'Table 2: Above',
        captionPosition: 'above',
        id: 'above-table',
      }, ast);
      await executeInsert({
        action: 'insert', type: 'paragraph', anchor: 'above-table', position: 'after',
        text: 'Directly after', id: 'node-direct',
      }, ast);
      const astIds = getNodeSequence(ast);
      expect(astIds.indexOf('node-direct')).toBe(astIds.indexOf('above-table') + 1);
    });

    it('keeps FIFO order past the caption for repeated after-table inserts', async () => {
      const { ast, tableId } = await tableWithBelowCaption();
      await executeInsert({
        action: 'insert', type: 'paragraph', anchor: tableId, position: 'after',
        text: 'First', id: 'node-first',
      }, ast);
      await executeInsert({
        action: 'insert', type: 'paragraph', anchor: tableId, position: 'after',
        text: 'Second', id: 'node-second',
      }, ast);
      const astIds = getNodeSequence(ast);
      const captionIdx = astIds.findIndex((id) => id.startsWith('tcap_'));
      expect(astIds.slice(captionIdx + 1, captionIdx + 3)).toEqual(['node-first', 'node-second']);
    });

    it('still splits correctly when the caption arrives after the first insert', async () => {
      const { ast, anchor } = await freshDoc();
      await executeInsert({
        action: 'insert', type: 'table', anchor, position: 'after',
        headers: ['A'], rows: [['1']], id: 'late-table',
      }, ast);
      await executeInsert({
        action: 'insert', type: 'paragraph', anchor: 'late-table', position: 'after',
        text: 'Early', id: 'node-early',
      }, ast);
      await executeTable({
        action: 'table', target: 'late-table',
        params: { op: 'setCaption', captionText: 'Late caption', captionPosition: 'below' },
      } as any, ast);
      const astIds = getNodeSequence(ast);
      const tableIdx = astIds.indexOf('late-table');
      const captionIdx = astIds.findIndex((id) => id.startsWith('tcap_'));
      expect(captionIdx).toBe(tableIdx + 1);
      await executeInsert({
        action: 'insert', type: 'paragraph', anchor: 'late-table', position: 'after',
        text: 'Late', id: 'node-late',
      }, ast);
      const afterIds = getNodeSequence(ast);
      expect(afterIds.indexOf('node-late')).toBeGreaterThan(afterIds.findIndex((id) => id.startsWith('tcap_')));
    });
  });

});
