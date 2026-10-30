import { DocumentAST } from '../ast';
import { BaseOperation } from './base';
import { DocxParser } from '../parser';
import { resolveRevisionStamp } from '../ooxml/track-changes-emitter';

export interface TrackedChangeOperation extends BaseOperation {
  action: 'trackedChange';
  operation: 'acceptAll' | 'rejectAll' | 'insert' | 'delete';
  target?: string;
  text?: string;
  author?: string;
  date?: string;
}

export function executeTrackedChange(op: TrackedChangeOperation, ast: DocumentAST, parser: DocxParser): void {
  const { author, date } = resolveRevisionStamp(op.author, op.date, ast.settings.defaultAuthor);

  if (op.operation === 'acceptAll') {
    parser.applyTrackedChanges('accept');
  } else if (op.operation === 'rejectAll') {
    parser.applyTrackedChanges('reject');
  } else if (op.operation === 'insert') {
    if (!op.target) {
      throw new Error('Tracked change "insert" requires a "target" node ID.');
    }
    if (typeof op.text !== 'string') {
      throw new Error('Tracked change "insert" requires "text" string content.');
    }
    parser.insertTrackedChange(op.target, op.text, author, date);
  } else if (op.operation === 'delete') {
    if (!op.target) {
      throw new Error('Tracked change "delete" requires a "target" node ID.');
    }
    parser.deleteTrackedChange(op.target, author, date);
  } else {
    throw new Error(
      `Unsupported trackedChange operation: ${op.operation}. ` +
      `Supported: acceptAll, rejectAll, insert, delete. ` +
      `Re-read the current structure with inspect_document and retry with a supported operation.`
    );
  }
}

