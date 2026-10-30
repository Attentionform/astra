import type { DocumentAST, Node } from './ast';
import type { DocxParser } from './parser';

/**
 * One way to bring a live AST up to date after the package under it changed.
 *
 * The AST object is shared (the session, the engine and callers hold the same
 * reference), so a re-read of the document cannot replace it: its sections are
 * spliced and its maps refilled in place.
 */

const MAP_KEYS = ['styles', 'numbering', 'relationships', 'headers', 'footers', 'mediaFiles'] as const;

export interface RefreshOptions {
  /**
   * The AST as it was before the package was rewound to that same state (a
   * failed or previewed batch). The document read now is the document that AST
   * described, so its blocks and sections get the ids and versions they had,
   * instead of the ones the file derives. Leave this out when the package
   * really changed (revert, sanitize, compose): ids are then read from the file.
   */
  carryIdentityFrom?: DocumentAST;
}

const paraIdOf = (node: Node): string => (node.xml as Element).getAttribute?.('w14:paraId') ?? '';
const tagOf = (node: Node): string => (node.xml as Element).tagName ?? '';

/**
 * Gives the freshly read blocks and sections the ids and versions of the ones
 * they replace. Only done when the two lists are provably the same blocks (same
 * count, same element at each position, same paragraph id where both have one);
 * otherwise nothing is changed and the returned sentence says so.
 */
function carryIdentity(parser: DocxParser, before: DocumentAST, fresh: DocumentAST): string | null {
  const oldNodes = before.sections.flatMap((section) => section.nodes);
  const newNodes = fresh.sections.flatMap((section) => section.nodes);
  const lost =
    'Ids were re-read from the document after the rollback, so ids returned earlier in this session may no longer resolve. ' +
    'Re-read the document with inspect_document.';

  if (oldNodes.length !== newNodes.length) return lost;
  for (let i = 0; i < oldNodes.length; i++) {
    if (tagOf(oldNodes[i]) !== tagOf(newNodes[i])) return lost;
    const oldPara = paraIdOf(oldNodes[i]);
    if (oldPara !== '' && oldPara !== paraIdOf(newNodes[i])) return lost;
  }

  const renames = new Map<string, string>();
  for (let i = 0; i < oldNodes.length; i++) {
    if (oldNodes[i].id !== newNodes[i].id) renames.set(newNodes[i].id, oldNodes[i].id);
  }
  if (!parser.reassignNodeIds(renames)) return lost;
  for (let i = 0; i < oldNodes.length; i++) {
    newNodes[i].id = oldNodes[i].id;
    newNodes[i].version = oldNodes[i].version;
  }

  const sameSections =
    before.sections.length === fresh.sections.length &&
    before.sections.every((section, i) => section.nodes.length === fresh.sections[i].nodes.length);
  if (sameSections) {
    before.sections.forEach((section, i) => {
      fresh.sections[i].id = section.id;
      fresh.sections[i].version = section.version;
    });
    return null;
  }
  return 'Section ids were re-read from the document after the rollback.';
}

/**
 * Re-reads the document from the parser and updates `target` to match, in
 * place. Returns sentences to show the caller when something could not be kept.
 */
export async function refreshAstInPlace(
  target: DocumentAST,
  parser: DocxParser,
  options: RefreshOptions = {},
): Promise<string[]> {
  const fresh = await parser.getAST();
  const notes: string[] = [];
  if (options.carryIdentityFrom) {
    const note = carryIdentity(parser, options.carryIdentityFrom, fresh);
    if (note) notes.push(note);
  }

  target.sections.splice(0, target.sections.length, ...fresh.sections);
  for (const key of MAP_KEYS) {
    const map = target[key] as Record<string, unknown>;
    for (const k of Object.keys(map)) delete map[k];
    Object.assign(map, fresh[key] as Record<string, unknown>);
  }
  const metadata = target.metadata as Record<string, unknown>;
  for (const k of Object.keys(metadata)) delete metadata[k];
  Object.assign(metadata, fresh.metadata);
  target.nextChangeId = fresh.nextChangeId;
  // Bare rezip() reads numbering from the AST the parser last handed out; point it at the one callers hold.
  parser.adoptAstForNumbering(target);
  return notes;
}
