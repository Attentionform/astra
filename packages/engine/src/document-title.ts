import type { DocumentAST } from './ast';
import type { DocxParser } from './parser';
import { CorePropertiesManipulator } from './manipulators/core-properties';
import { newDocumentTitle } from './ooxml/new-document-parts';

/** The names of the built-in templates; each one except blank writes its own name into dc:title of a new document. */
const TEMPLATE_NAMES = [
  'blank', 'memo', 'report', 'letter', 'proposal', 'invoice', 'research', 'corporate', 'financial',
  'novel', 'academic', 'newsletter', 'api-spec', 'bar-chart-3d',
] as const;

const PLACEHOLDER_TITLES: ReadonlySet<string> = new Set(TEMPLATE_NAMES.map((name) => newDocumentTitle(name).toLowerCase()));

/** The longest title written from the content (the Word Info panel shows one line). */
export const DERIVED_TITLE_MAX = 200;

/** The tool name: the application a file was made by, and the creator older files may carry. A new file carries no creator. */
const DEFAULT_CREATOR = 'ASTra';

/** True when docProps/app.xml names ASTra as the application: the file was created by ASTra, not imported from elsewhere. */
function createdByAstra(parser: DocxParser | undefined): boolean {
  const app = parser?.getZipAdapter().readTextOrNull('docProps/app.xml');
  if (!app) return false;
  const name = /<(?:[\w.-]+:)?Application\b[^>]*>([^<]*)</.exec(app)?.[1];
  return name !== undefined && name.trim().toLowerCase() === DEFAULT_CREATOR.toLowerCase();
}

/** True for an empty title and for the name of a built-in template ("Report", "Memo"). */
export function isPlaceholderTitle(title: unknown): boolean {
  if (typeof title !== 'string') return true;
  const clean = title.replace(/\s+/g, ' ').trim().toLowerCase();
  return clean === '' || PLACEHOLDER_TITLES.has(clean);
}

function cleanTitle(text: unknown): string {
  const plain = typeof text === 'string' ? Array.from(text, (ch) => (ch.charCodeAt(0) < 32 || ch.charCodeAt(0) === 127 ? ' ' : ch)).join('').replace(/\s+/g, ' ').trim() : '';
  return plain.length > DERIVED_TITLE_MAX ? plain.slice(0, DERIVED_TITLE_MAX).trimEnd() : plain;
}

export interface DerivedTitle {
  title: string;
  /** Where it came from: the cover page, a paragraph in the Title style, or the first Heading 1. */
  from: 'cover page' | 'Title' | 'Heading 1';
}

/**
 * The title the content proposes: the cover page title, else the first paragraph in the Title style,
 * else the first Heading 1. Text that is itself a template placeholder is no candidate.
 */
export function deriveTitleFromContent(ast: DocumentAST): DerivedTitle | null {
  const nodes = ast.sections.flatMap((section) => section.nodes);
  type AstNode = (typeof nodes)[number];
  // The style and the text are read from the paragraph itself: the node's cached metadata is not
  // refreshed by every operation (a style applied by namedStyle leaves it stale).
  const styleOf = (node: AstNode): string | undefined => {
    const pPr = Array.from(node.xml.childNodes as unknown as ArrayLike<Element>).find((c) => c.nodeType === 1 && c.tagName === 'w:pPr');
    const pStyle = pPr ? Array.from(pPr.childNodes as unknown as ArrayLike<Element>).find((c) => c.nodeType === 1 && c.tagName === 'w:pStyle') : undefined;
    return pStyle?.getAttribute('w:val') ?? node.metadata?.namedStyle ?? undefined;
  };
  const textOf = (node: AstNode): string => {
    const runs = node.xml.getElementsByTagName('w:t');
    let text = '';
    for (let i = 0; i < runs.length; i++) text += runs[i].textContent ?? '';
    return text !== '' ? text : (node.metadata?.text ?? '');
  };
  const pick = (match: (node: AstNode) => boolean, from: DerivedTitle['from']): DerivedTitle | null => {
    for (const node of nodes) {
      if (node.type !== 'paragraph' || !match(node)) continue;
      const title = cleanTitle(textOf(node));
      if (title !== '' && !isPlaceholderTitle(title)) return { title, from };
    }
    return null;
  };
  return (
    pick((n) => n.id.startsWith('cover_title_'), 'cover page') ??
    pick((n) => styleOf(n) === 'Title', 'Title') ??
    pick((n) => styleOf(n) === 'Heading1', 'Heading 1')
  );
}

export interface DocumentIdentityResult {
  /** Set when dc:title was a placeholder and is now written from the content. */
  titleSetFrom?: DerivedTitle['from'];
  title?: string;
  /** Set when the creator was the tool default and a configured author replaced it. */
  creatorSet?: string;
}

/**
 * A saved document should never carry a placeholder title (the template's own name, in a file ASTra made), and a document made under a configured
 * author (--author, ASTRA_AUTHOR, or the session's author) should name that author. Both changes are
 * made only when the current value is the tool's own default: a title or creator that the caller or
 * the file already had is never touched. Returns what it changed so the caller can report it.
 */
export function applyDocumentIdentity(ast: DocumentAST, parser: DocxParser | undefined): DocumentIdentityResult {
  const result: DocumentIdentityResult = {};
  const patch: { title?: string; creator?: string; lastModifiedBy?: string } = {};

  // The name of a template ("Report", "Memo") is a placeholder only in a file ASTra made from that template; in a file
  // from another application it is a title somebody chose. A missing or empty title is replaceable either way.
  const title = ast.metadata?.title;
  const emptyTitle = typeof title !== 'string' || title.trim() === '';
  if (emptyTitle || (isPlaceholderTitle(title) && createdByAstra(parser))) {
    const derived = deriveTitleFromContent(ast);
    if (derived) {
      patch.title = derived.title;
      result.title = derived.title;
      result.titleSetFrom = derived.from;
    }
  }

  const author = typeof ast.settings?.defaultAuthor === 'string' ? ast.settings.defaultAuthor.trim() : '';
  const creator = typeof ast.metadata?.creator === 'string' ? ast.metadata.creator.trim() : '';
  // A file ASTra made carries no creator (or the tool name): a configured author fills it. A creator somebody set is kept.
  const creatorIsToolDefault = creator === DEFAULT_CREATOR || (creator === '' && createdByAstra(parser));
  if (author !== '' && author !== DEFAULT_CREATOR && creatorIsToolDefault) {
    patch.creator = author;
    patch.lastModifiedBy = author;
    result.creatorSet = author;
  }

  if (Object.keys(patch).length === 0) return result;
  new CorePropertiesManipulator(ast, parser?.getZipAdapter()).update(patch);
  return result;
}
