/**
 * Grep-based text search for document nodes
 * Fast, simple alternative to semantic search for structured documents
 */

/**
 * Options for `grepNodes`: `regex` reads the query as a regular expression, and `caseSensitive` matches case.
 */
export interface GrepOptions {
  regex?: boolean;
  caseSensitive?: boolean;
}

/**
 * One match: the node `id`, its `preview`, and `matchContext`, the text around the match.
 */
export interface GrepResult {
  id: string;
  preview: string;
  matchContext: string;
}

/**
 * Searches a list of nodes that you prepared for a text, and returns the matches with some text around each. `query` is plain text, or a regular expression when `options.regex` is true.
 * The search ignores case unless `options.caseSensitive` is true. An invalid regular expression gives an empty list.
 */
export function grepNodes(
  nodes: { id: string; text: string; preview: string }[],
  query: string,
  options: GrepOptions = {}
): GrepResult[] {
  // Don't use 'g' flag for test() - it causes issues with lastIndex
  const flags = options.caseSensitive ? '' : 'i';
  
  let pattern: RegExp;
  try {
    pattern = options.regex 
      ? new RegExp(query, flags)
      : new RegExp(escapeRegex(query), flags);
  } catch {
    // Invalid regex, return empty results
    return [];
  }
  
  return nodes
    .filter(n => pattern.test(n.text))
    .map(n => ({
      id: n.id,
      preview: n.preview,
      matchContext: extractMatchContext(n.text, pattern, 100)
    }));
}

/**
 * Extract context around the first match
 * @param text - Full text content
 * @param pattern - Regex pattern to match
 * @param contextChars - Number of characters to show before/after match
 * @returns Context string with match highlighted
 */
function extractMatchContext(text: string, pattern: RegExp, contextChars: number): string {
  // Reset regex state
  pattern.lastIndex = 0;
  
  const match = pattern.exec(text);
  if (!match) return '';
  
  const start = Math.max(0, match.index - contextChars);
  const end = Math.min(text.length, match.index + match[0].length + contextChars);
  
  let context = text.slice(start, end);
  
  // Add ellipsis if truncated
  if (start > 0) context = '...' + context;
  if (end < text.length) context = context + '...';
  
  return context;
}

/**
 * Escape special regex characters for literal matching
 * @param str - String to escape
 * @returns Escaped string safe for use in RegExp
 */
function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
