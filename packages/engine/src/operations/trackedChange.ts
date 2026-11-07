import { buildRunModel, applyChunks } from '../ooxml/run-model';
import { findNode, validateVersion } from './base';
import { createWordElement } from '../ooxml/ns';
import type { RevisionResolution } from '../ooxml/revisions';
import { DocumentAST } from '../ast';
import { BaseOperation } from './base';
import { DocxParser } from '../parser';
import { resolveRevisionStamp } from '../ooxml/track-changes-emitter';
import { PPR_CHILD_ORDER, sortChildrenBySchema } from '../ooxml/schema-order';

export interface TrackedChangeOperation extends BaseOperation {
  action: 'trackedChange';
  operation: 'acceptAll' | 'rejectAll' | 'accept' | 'reject' | 'insert' | 'delete';
  ids?: number[];
  target?: string;
  text?: string;
  occurrence?: number;
  author?: string;
  date?: string;
}

export function executeTrackedChange(op: TrackedChangeOperation, ast: DocumentAST, parser: DocxParser): RevisionResolution | { revisionIds: number[] } | void {
  const { author, date } = resolveRevisionStamp(op.author, op.date, ast.settings.defaultAuthor);

  if (op.operation === 'acceptAll') {
    return parser.resolvePackageRevisions('accept');
  } else if (op.operation === 'rejectAll') {
    return parser.resolvePackageRevisions('reject');
  } else if (op.operation === 'accept' || op.operation === 'reject') {
    return parser.resolvePackageRevisions(op.operation, {ids:op.ids, author:op.author, nodeId:op.target});
  } else if (op.operation === 'insert') {
    if (!op.target) {
      throw new Error('Tracked change "insert" requires a "target" node ID.');
    }
    if (typeof op.text !== 'string') {
      throw new Error('Tracked change "insert" requires "text" string content.');
    }
    const node = findNode(ast, op.target);
    if (node) validateVersion(node, op.expectedVersion);
    const id = ast.nextChangeId;
    parser.insertTrackedChange(op.target, op.text, author, date, undefined, () => ast.nextChangeId++);
    if (node && (node.xml as Element).tagName === 'w:p') {
      node.metadata.text += op.text;
      node.metadata.hasTrackedChanges = true;
      node.version += 1;
    }
    return {revisionIds:[id]};
  } else if (op.operation === 'delete') {
    if (!op.target) {
      throw new Error('Tracked change "delete" requires a "target" node ID.');
    }
    const node = findNode(ast, op.target);
    if (!node || !node.xml || (node.xml as Element).tagName !== 'w:p') {
      throw new Error('E_VALIDATION: Tracked deletion requires a paragraph target.');
    }
    validateVersion(node, op.expectedVersion);
    const paragraph = node.xml as Element;
    if (op.text === undefined) {
      const pPr = Array.from(paragraph.childNodes).find(n => n.nodeType === 1 && (n as Element).tagName === 'w:pPr') as Element | undefined;
      const rPr = pPr && Array.from(pPr.childNodes).find(n => n.nodeType === 1 && (n as Element).tagName === 'w:rPr') as Element | undefined;
      if (rPr && Array.from(rPr.childNodes).some(n => n.nodeType === 1 && ['w:ins','w:del'].includes((n as Element).tagName))) {
        throw new Error('E_VALIDATION: The paragraph mark already has a tracked revision; resolve it before deleting the paragraph.');
      }
    }
    const model = buildRunModel(paragraph, node.metadata.text);
    if (op.text === '') throw new Error('E_VALIDATION: text must not be empty; omit it to delete the paragraph.');
    let at = 0;
    let length = model.text.length;
    if (op.text !== undefined) {
      const hits: number[] = [];
      for (let i = 0; i <= model.text.length - op.text.length; i++) if (model.text.startsWith(op.text, i)) hits.push(i);
      if (!hits.length) throw new Error('E_VALIDATION: text not found in paragraph.');
      if (hits.length > 1 && op.occurrence === undefined) throw new Error(`E_VALIDATION: text occurs ${hits.length} times; provide occurrence.`);
      if (op.occurrence !== undefined && op.occurrence > hits.length) throw new Error('E_VALIDATION: occurrence is out of range.');
      at = hits[(op.occurrence ?? 1) - 1];
      length = op.text.length;
    }
    const chunks: {type:'equal'|'delete';text:string}[] = [
      {type:'equal',text:model.text.slice(0,at)},
      {type:'delete',text:model.text.slice(at,at+length)},
      {type:'equal',text:model.text.slice(at+length)},
    ];
    const revisionIds = applyChunks(model, chunks, {tracked:true, author, date, nextId:()=>ast.nextChangeId++});
    if (op.text === undefined) {
      const doc = paragraph.ownerDocument!;
      let pPr = Array.from(paragraph.childNodes).find(n => n.nodeType === 1 && (n as Element).tagName === 'w:pPr') as Element | undefined;
      if (!pPr) { pPr = createWordElement(doc, 'w:pPr'); paragraph.insertBefore(pPr, paragraph.firstChild); }
      let rPr = Array.from(pPr.childNodes).find(n => n.nodeType === 1 && (n as Element).tagName === 'w:rPr') as Element | undefined;
      if (!rPr) { rPr = createWordElement(doc, 'w:rPr'); pPr.appendChild(rPr); sortChildrenBySchema(pPr, PPR_CHILD_ORDER); }
      const mark = createWordElement(doc, 'w:del');
      const id = ast.nextChangeId++; revisionIds.push(id);
      const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      mark.setAttributeNS(ns, 'w:id', String(id)); mark.setAttributeNS(ns, 'w:author', author); mark.setAttributeNS(ns, 'w:date', date);
      const change = Array.from(rPr.childNodes).find(n => n.nodeType === 1 && (n as Element).tagName === 'w:rPrChange');
      rPr.insertBefore(mark, change ?? null);
    }
    node.metadata.text = model.text.slice(0,at) + model.text.slice(at+length);
    node.metadata.hasTrackedChanges = true;
    node.version += 1;
    return { revisionIds };
  } else {
    throw new Error(
      `Unsupported trackedChange operation: ${op.operation}. ` +
      `Supported: acceptAll, rejectAll, insert, delete. ` +
      `Re-read the current structure with inspect_document and retry with a supported operation.`
    );
  }
}

