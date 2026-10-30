import { DocumentAST, Node, Section } from '../ast';
import { BaseOperation, validateVersion } from './base';
import { DocxParser } from '../parser';
import { insertIntoSection } from '../parser/sections';

export interface MoveOperation extends BaseOperation {
  action: 'move';
  target: string | string[];
  anchor: string;
  position: 'before' | 'after';
  params?: Record<string, unknown>;
}

export function executeMove(op: MoveOperation, ast: DocumentAST, _parser?: DocxParser): void {
  void _parser;
  // OP-16: accept args nested under params:{} or flat on the operation (mirrors table.ts pick).
  const raw = (op.params || {}) as Record<string, unknown>;
  const pick = <T>(key: string, flat: T | undefined): T | undefined =>
    raw[key] !== undefined ? (raw[key] as T) : flat;
  const resolvedTarget = (pick('target', op.target) ?? op.target) as string | string[];
  const resolvedAnchor = String(pick('anchor', op.anchor) ?? op.anchor ?? '');
  const resolvedPosition = (pick('position', op.position) ?? op.position ?? 'after') as 'before' | 'after';
  const targets = Array.isArray(resolvedTarget) ? resolvedTarget : (resolvedTarget ? [resolvedTarget] : []);
  if (targets.length === 0) throw new Error("Move operation requires at least one target");

  if (targets.includes(resolvedAnchor)) {
    throw new Error("Cannot move a node relative to itself");
  }

  const nodesToMove: { node: Node, section: Section, index: number }[] = [];

  // 1. Detach targets
  for (const target of targets) {
    let found = false;
    for (const section of ast.sections) {
      const index = section.nodes.findIndex(n => n.id === target);
      if (index !== -1) {
        const node = section.nodes[index];
        validateVersion(node, op.expectedVersion);
        nodesToMove.push({ node, section, index });
        
        // Remove from DOM temporarily
        node.xml.parentNode?.removeChild(node.xml);
        // Remove from AST
        section.nodes.splice(index, 1);
        found = true;
        break;
      }
    }
    if (!found) throw new Error(`Target node ${target} not found`);
  }

  // 2. Find anchor and attach
  let anchorSection: Section | null = null;
  let anchorIndex = -1;
  let anchorNode: Node | undefined;

  for (const section of ast.sections) {
    anchorIndex = section.nodes.findIndex(n => n.id === resolvedAnchor);
    if (anchorIndex !== -1) {
      anchorSection = section;
      anchorNode = section.nodes[anchorIndex];
      break;
    }
  }

  if (!anchorNode || !anchorSection) throw new Error(`Anchor node ${resolvedAnchor} not found`);

  // Insert backwards if position is 'before' to maintain original target order
  const insertIndex = resolvedPosition === 'before' ? anchorIndex : anchorIndex + 1;
  
  for (const { node } of nodesToMove) {
    // Insert into DOM
    if (resolvedPosition === 'before') {
      anchorNode.xml.parentNode?.insertBefore(node.xml, anchorNode.xml);
    } else {
      // To keep correct order when inserting multiple nodes after an anchor,
      // we insert after the anchor and then advance the anchor reference.
      anchorNode.xml.parentNode?.insertBefore(node.xml, anchorNode.xml.nextSibling);
      anchorNode = node; // Move anchor forward
    }
    node.version += 1;
  }

  // Insert into AST, all at once and in order. After a section's break
  // paragraph is the start of the next section, as in the file.
  insertIntoSection(ast.sections, anchorSection, insertIndex, nodesToMove.map(({ node }) => node));
}
