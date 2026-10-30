import { DocumentAST, Node } from '../ast';
import { BaseOperation, validateVersion, findNode } from './base';
import {
  addComment,
  addCommentReply,
  anchorCommentReply,
  findLastTopLevelCommentId,
} from '../ooxml/comments';
import { createWordElement } from '../ooxml/ns';
import { DocxParser } from '../parser';
import { DOMParser } from '@xmldom/xmldom';
import { resolveDeterministicAuthor } from '../ooxml/track-changes-emitter';
import { toError } from '../errors';

export interface CommentOperation extends BaseOperation {
  action: 'comment';
  operation: 'insert' | 'reply';
  target: string; // The paragraph node ID or parent comment ID for replies
  text: string;
  author?: string;
  date?: string;
}

function firstDirectChild(parent: Element, tag: string): Element | null {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i] as Element;
    if (child.nodeType === 1 && child.tagName === tag) return child;
  }
  return null;
}

/**
 * A-3: comment markers (w:commentRangeStart/w:r) are only valid inside a
 * w:p. A table node routes to its first-cell paragraph
 * (w:tr[0]->w:tc[0]->w:p[0], created if absent) instead of emitting
 * w:r-under-w:tbl. Degenerate cell-less tables reject loudly.
 */
function resolveCommentHost(targetNode: Node): Element {
  if (targetNode.type !== 'table') return targetNode.xml;
  const firstRow = firstDirectChild(targetNode.xml, 'w:tr');
  const firstCell = firstRow ? firstDirectChild(firstRow, 'w:tc') : null;
  if (!firstCell) {
    throw new Error(
      `Cannot comment on table node '${targetNode.id}': table has no rows/cells to host the comment. Remedy: target a paragraph node id from inspect_document.`
    );
  }
  let firstPara = firstDirectChild(firstCell, 'w:p');
  if (!firstPara) {
    firstPara = createWordElement(firstCell.ownerDocument, 'w:p');
    firstCell.insertBefore(firstPara, firstCell.firstChild);
  }
  return firstPara;
}

/**
 * The comment writer adds its relationships (comments, commentsExtended, people) straight to
 * document.xml.rels. Publish them to ast.relationships too, because the image and note
 * allocators pick the next free id from that map: without this, a later insert in the same
 * batch reuses an id the comment part already holds ('rId2' ID conflicts).
 */
function publishPackageRelationships(ast: DocumentAST, parser: DocxParser): void {
  const relsXml = parser.getZipAdapter().readTextOrNull('word/_rels/document.xml.rels');
  if (!relsXml) return;
  const doc = new DOMParser().parseFromString(relsXml, 'text/xml');
  const rels = doc.getElementsByTagName('Relationship');
  for (let i = 0; i < rels.length; i++) {
    const rel = rels[i] as Element;
    const id = rel.getAttribute('Id');
    if (!id || Object.prototype.hasOwnProperty.call(ast.relationships, id)) continue;
    ast.relationships[id] = {
      type: rel.getAttribute('Type') ?? '',
      target: rel.getAttribute('Target') ?? '',
      targetMode: rel.getAttribute('TargetMode') === 'External' ? 'External' : 'Internal',
    };
  }
}

export function executeComment(op: CommentOperation, ast: DocumentAST, parser: DocxParser): void {
  const zipAdapter = parser.getZipAdapter();
  const author = resolveDeterministicAuthor(op.author, ast.settings.defaultAuthor);

  if (op.operation === 'insert') {
    const targetNode: Node | undefined = findNode(ast, op.target);

    if (!targetNode) {
      throw new Error(`Node ${op.target} not found in AST`);
    }

    validateVersion(targetNode, op.expectedVersion);
    addComment(zipAdapter, resolveCommentHost(targetNode), author, op.text, op.date);
    publishPackageRelationships(ast, parser);
    
    targetNode.version += 1;
    targetNode.metadata.hasComment = true;
  } else if (op.operation === 'reply') {
    const targetNode: Node | undefined = findNode(ast, op.target);

    let parentCommentId: number;
    if (targetNode) {
      const foundId = findLastTopLevelCommentId(zipAdapter, resolveCommentHost(targetNode));
      if (foundId === null) {
        throw toError(
          'E_VALIDATION',
          `Node ${op.target} has no comment to reply to.`,
          'Target a paragraph that carries comments or pass a numeric comment id.'
        );
      }
      parentCommentId = foundId;
    } else if (/^\d+$/.test(op.target)) {
      parentCommentId = parseInt(op.target, 10);
    } else {
      throw toError(
        'E_VALIDATION',
        `Invalid parent for reply: '${op.target}' is neither a comment id nor a node id.`,
        'Target a paragraph node id from inspect_document or pass a numeric comment id.'
      );
    }

    const replyId = addCommentReply(zipAdapter, parentCommentId, author, op.text, op.date);
    publishPackageRelationships(ast, parser);
    anchorCommentReply(parser.getDocumentDom(), parentCommentId, replyId);
  }
}
