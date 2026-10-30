import { createWordElement } from '../ooxml/ns';
import { toError } from '../errors';
import { DocumentAST, Node } from '../ast';
import { BaseOperation, findNode } from './base';
import { appendInSchemaOrder, PPR_CHILD_ORDER } from '../ooxml/schema-order';

export interface NamedStyleOperation extends BaseOperation {
  action: 'namedStyle';
  operation: 'apply' | 'create' | 'modify' | 'delete';
  styleName: string;
  target?: string | string[]; // Required for 'apply'
}

export function executeNamedStyle(op: NamedStyleOperation, ast: DocumentAST, parser?: { getStyleIds(): string[] }): void {
  if (op.operation === 'apply') {
    const targets = Array.isArray(op.target) ? op.target : (op.target ? [op.target] : []);
    for (const target of targets) {
      const targetNode: Node | undefined = findNode(ast, target);
      if (!targetNode) throw new Error(`Node ${target} not found`);

      if ((targetNode.xml as Element).tagName !== 'w:p') {
        throw toError(
          'E_VALIDATION',
          `namedStyle.apply applies a paragraph style, but '${target}' is a ${targetNode.type} node.`,
          "Target a paragraph id (see inspect_document). To style a table use action 'table'."
        );
      }

      const doc = targetNode.xml.ownerDocument;
      if (!doc) throw new Error(`Node ${target} has no owner document; cannot apply style '${op.styleName}'`);

      let pPr = targetNode.xml.getElementsByTagName('w:pPr')[0];
      if (!pPr) {
        pPr = createWordElement(doc, 'w:pPr');
        targetNode.xml.insertBefore(pPr, targetNode.xml.firstChild);
      }
      
      let pStyle = pPr.getElementsByTagName('w:pStyle')[0];
      if (!pStyle) {
        pStyle = createWordElement(doc, 'w:pStyle');
        pPr.appendChild(pStyle);
      }
      // Word matches w:val against w:styleId (e.g. 'Heading1'), not the display
      // name ('Heading 1'). Resolve against the real style sheet, else strip spaces.
      const declared = parser ? parser.getStyleIds() : [];
      const compact = String(op.styleName || '').replace(/\s+/g, '');
      const styleId =
        declared.find(id => id === op.styleName) ??
        declared.find(id => String(id || '').replace(/\s+/g, '').toLowerCase() === compact.toLowerCase()) ??
        compact;
      pStyle.setAttribute('w:val', styleId);

      // Headings stay with the text that follows: keep them glued to the
      // next paragraph so a heading is never stranded at a page bottom.
      // Additive only (an existing keepNext is left alone) and placed in
      // schema order. The save-time typography pass keeps this invariant
      // for headings arriving through other paths.
      const headingLike =
        /^heading\s*\d/i.test(styleId) ||
        /^heading\d/i.test(styleId) ||
        /^title/i.test(styleId) ||
        /^subtitle/i.test(styleId) ||
        pPr.getElementsByTagName('w:outlineLvl').length > 0;
      if (headingLike) {
        let hasKeepNext = false;
        for (let i = 0; i < pPr.childNodes.length; i++) {
          const c = pPr.childNodes[i] as unknown as Element;
          if ((c as unknown as { nodeType?: number }).nodeType === 1 && c.tagName === 'w:keepNext') {
            hasKeepNext = true;
            break;
          }
        }
        if (!hasKeepNext) {
          appendInSchemaOrder(pPr, createWordElement(doc, 'w:keepNext'), PPR_CHILD_ORDER);
        }
      }

      targetNode.metadata.namedStyle = styleId;
      targetNode.version += 1;
    }
  } else if (op.operation === 'create' || op.operation === 'modify' || op.operation === 'delete') {
    throw new Error(`namedStyle.${op.operation} is not supported yet`);
  }
}
